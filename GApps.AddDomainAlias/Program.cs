using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Google.GData.Apps;
using Google.GData.Client;
using Google.GData.Extensions.Apps;
using Mono.Options; 

namespace WolfeReiter.GoogleApps
{
    class Program
    {
        static void Main( string[] args )
        {
            //parse options
            var options = GoogleAppsAddDomainAliasProgramOptions.Parse( args );
            if( options.Help ) { options.PrintHelp( Console.Out ); return; }
            if( options.Incomplete ) { CollectOptionsInteractiveConsole( options ); }
            if( !options.Username.Contains( "@" ) )
            {
                Console.WriteLine( "Whoops. Your username must be an email address." );
                options.Username = null;
                CollectOptionsInteractiveConsole( options );
            }
            //create service object
            var service = new MultiDomainManagementService( options.Domain, null );
            service.setUserCredentials( options.Username, options.Password );

            try
            {
                //add aliases
                if( string.IsNullOrEmpty( options.File ) ) { AddAliasInteractiveConsole( service, options.Domain ); }
                else { AddAliasBatch( Console.Out, service, options.Domain, options.File ); }
            }
            catch( InvalidCredentialsException )
            {
                Console.WriteLine();
                Console.WriteLine( "Invalid Credentials." );
            }
            catch( CaptchaRequiredException )
            {
                Console.WriteLine();
                Console.WriteLine( "Your account has been locked by Google." );
                Console.WriteLine( "Use your browser to unlock your account." );
                Console.WriteLine( "https://www.google.com/accounts/UnlockCaptcha" );
            }
        }

        static string AddAlias( MultiDomainManagementService service, string primaryDomain, string user, string alias )
        {
            try
            {
                var entry = service.CreateDomainUserAlias( primaryDomain, user, alias );
                return string.Format(
                    "Alias {0} created for {1}.",
                    entry.getPropertyValueByName( AppsMultiDomainNameTable.AliasEmail ),
                    user ); //return alias from server
            }
            catch( GDataRequestException ex )
            {
                return string.Format( "Error creating alias: {0}.", ex.ParseError() ); //or else return error message
            }
        }

        static void AddAliasInteractiveConsole( MultiDomainManagementService service, string primaryDomain )
        {
            Console.WriteLine();
            Console.WriteLine( "No batch file option (-f). Entering interactive mode." );
            Console.WriteLine( "Press CTRL+C to quit." );
            Console.WriteLine();
            while( true ) //continue until CTRL+C
            {
                bool confirm = false;
                string user=null, alias=null;
                while( string.IsNullOrEmpty( user ) )
                {
                    Console.Write( "User to alias [user@domain]: " );
                    user = Console.ReadLine();
                }
                while( string.IsNullOrEmpty( alias ) )
                {
                    Console.Write( String.Format( "Alias for {0}: ", user ) );
                    alias = Console.ReadLine();
                }
                
                Console.Write( string.Format("Please confirm: Add alias {0} to {1} (y/n)? ", alias, user) );
                confirm = Console.ReadLine().StartsWith( "y", StringComparison.InvariantCultureIgnoreCase );
                if( !confirm )
                {
                    Console.WriteLine( "Cancelled. Alias not created.");
                }
                else
                {
                    try
                    {
                        Console.WriteLine( AddAlias( service, primaryDomain, user, alias ) );
                    }
                    catch( InvalidCredentialsException )
                    {
                        Console.WriteLine();
                        Console.WriteLine( "Invalid Credentials." );
                        Console.WriteLine();
                        //collect new credentials
                        var options = new GoogleAppsAddDomainAliasProgramOptions() { Domain = primaryDomain }; 
                        CollectOptionsInteractiveConsole( options );
                        service.setUserCredentials( options.Username, options.Password );
                    }
                }
                Console.WriteLine(); 
            }
        }

        static void AddAliasBatch( TextWriter textWriter, MultiDomainManagementService service, string primaryDomain, string file )
        {
            try
            {
                using( var stream = File.OpenRead( file ) )
                using( var reader = new StreamReader( stream ) )
                {
                    string line;
                    long position=0;
                    while( null != (line = reader.ReadLine()) )
                    {
                        position++;
                        line = line.Trim();
                        if( line.Length == 0 || line.StartsWith( "#" ) ) { continue; } //ignore lines that start with '#'
                        
                        string response;
                        var split = line.Split('\t');
                        if( split.Length != 2 )
                        {
                            response = string.Format( "Line {0}: Error row not in the expected format.", position );
                        }
                        else
                        {
                            string user = split[0];
                            string alias = split[1];
                            string result = AddAlias( service, primaryDomain, user, alias );
                            response = string.Format( "Line {0}: {1}", position, result );
                        }
                        textWriter.WriteLine( response );
                    }
                }
                
            }
            catch( IOException ex )
            {
                textWriter.WriteLine();
                textWriter.WriteLine( ex.Message );
            }
        }

        static void CollectOptionsInteractiveConsole( GoogleAppsAddDomainAliasProgramOptions options )
        {
            while( string.IsNullOrEmpty( options.Domain ) )
            {
                Console.Write( "Primary Google Apps Domain: " );
                options.Domain = Console.ReadLine();
            }

            while( string.IsNullOrEmpty( options.Username ) )
            {
                Console.Write( "Admin username: " );
                options.Username = Console.ReadLine();
            }

            while( string.IsNullOrEmpty( options.Password ) )
            {
                Console.Write( "Password: " );
                StringBuilder stringBuilder = new StringBuilder();
                ConsoleKeyInfo keyInfo;
                while( ConsoleKey.Enter != (keyInfo = Console.ReadKey( true )).Key )
                {
                    if( ConsoleKey.Backspace == keyInfo.Key )
                    {
                        if( stringBuilder.Length > 0 )
                        {
                            stringBuilder.Remove( stringBuilder.Length - 1, 1 );
                            //backspace over '*' char, write space to erase and backspace again to position the cursor
                            Console.Write( "\b \b" );
                        }
                    }
                    else
                    {
                        stringBuilder.Append( keyInfo.KeyChar );
                        Console.Write( "*" );
                    }
                }
                Console.WriteLine();
                options.Password = stringBuilder.ToString();
            }
        }
    } 

    public static class GDataExtension
    {
        public static GDataError ParseError(this GDataRequestException exception )
        {
            try
            {
                var doc = XDocument.Parse( exception.ResponseString );
                return doc.Elements( "AppsForYourDomainErrors" ).Elements( "error" )
                        .Select( x => new GDataError( x.Attribute( "errorCode" ).Value, x.Attribute( "reason" ).Value ) )
                        .SingleOrDefault();
            }
            catch //failed to parse XML as expected
            {
                return new GDataError( string.Empty, exception.ResponseString );
            }
        }

        public sealed class GDataError
        {
            public GDataError(string errorCode, string reason)
            {
                ErrorCode = errorCode;
                Reason = reason;
            }
            public string ErrorCode { get; private set; }
            public string Reason { get; private set; }
            public override string ToString()
            {
                if( string.IsNullOrEmpty(ErrorCode) )
                { return Reason; }
                else
                { return string.Format( "{0} - {1}", ErrorCode, Reason ); }
            }
        }
    }

    public sealed class GoogleAppsAddDomainAliasProgramOptions
    {
        public static GoogleAppsAddDomainAliasProgramOptions Parse( string[] args )
        {
            var programOptions = new GoogleAppsAddDomainAliasProgramOptions();

            programOptions.Options = new OptionSet()
            {
                { "d|domain=", "Primary Google Apps {DOMAIN}", x => programOptions.Domain = x },
                { "u|username=", "{USERNAME}", x => programOptions.Username = x },
                { "p|password=", "{PASSWORD}", x => programOptions.Password = x },
                { "f|file=", "Tab-delimited {FILE} of user and alias", x => programOptions.File = x },
                { "?|help", "Show this message and exit.", x => programOptions.Help = (x != null) }
            };

            programOptions.Options.Parse( args );
                
            return programOptions;
        }

        private OptionSet Options { get; set; }
        public string Domain { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string File { get; set;  }
        public bool Help { get; set; }
        public bool Incomplete
        {
            get 
            {
                return string.IsNullOrEmpty( Username )
                    || string.IsNullOrEmpty( Password )
                    || string.IsNullOrEmpty( Domain );
            }
        }
        public void PrintHelp( TextWriter textWriter )
        {
            string appName = System.AppDomain.CurrentDomain.FriendlyName;
            textWriter.WriteLine( string.Format( "Usage: {0} [OPTIONS]+", appName ) );
            textWriter.WriteLine( "Command-line tool to create aliases for users in a multi-domain Google Apps instance." );
            textWriter.WriteLine( "{0} version {1}", appName, Assembly.GetExecutingAssembly().GetName().Version );
            textWriter.WriteLine();
            textWriter.WriteLine( "Options:" );
            Options.WriteOptionDescriptions( textWriter );
        }
    }
}
