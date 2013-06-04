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
            var options = GoogleAppsAddGroupMemberProgramOptions.Parse( args );
            if( options.Help ) { options.PrintHelp( Console.Out ); return; }
            if( options.Incomplete ) { CollectOptionsInteractiveConsole( options ); }
            if( !options.Username.Contains( "@" ) )
            {
                Console.WriteLine( "Whoops. Your username must be an email address." );
                Console.WriteLine();
                options.Username = null;
                CollectOptionsInteractiveConsole( options );
            }
            //create service object
            var service = new AppsService( options.Domain, options.Username, options.Password );
            try
            {
                if( options.List ) { ListGroups( Console.Out, service ); }
                else if( !string.IsNullOrEmpty( options.ListMembers ) ) { ListGroupMembers( Console.Out, service, options.ListMembers ); }
                else if( !string.IsNullOrEmpty( options.ListGroupsForMember) ) { ListGroupsForMember( Console.Out, service, options.ListGroupsForMember ); }
                else
                {
                    //add members to group
                    if( string.IsNullOrEmpty( options.File ) ) { AddGroupMemberInteractiveConsole( service, options.Domain ); }
                    else { AddGroupMemberBatch( Console.Out, service, options.Domain, options.File ); }
                }
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

        static void ListGroups( TextWriter textWriter, AppsService service )
        {
            textWriter.WriteLine( "All groups in {0}:", service.Domain );
            textWriter.WriteLine();
            try
            {
                var feed = service.Groups.RetrieveAllGroups();
                foreach( AppsExtendedEntry entry in feed.Entries ) //cast Entry to AppsExtendedEntry
                {
                    textWriter.WriteLine( entry.getPropertyValueByName( AppsGroupsNameTable.groupId ) );
                }
            }
            catch( GDataRequestException ex )
            {
                textWriter.WriteLine( string.Format( "Error: {0}.", ex.ParseError() ) );
            }
        }

        static void ListGroupMembers( TextWriter textWriter, AppsService service, string group )
        {
            textWriter.WriteLine( "All active members of {0}:", group );
            textWriter.WriteLine();
            try
            {
                var feed = service.Groups.RetrieveAllMembers(group);
                foreach( AppsExtendedEntry entry in feed.Entries ) //cast Entry to AppsExtendedEntry
                {
                    textWriter.WriteLine( entry.getPropertyValueByName( AppsGroupsNameTable.memberId ) );
                }
            }
            catch( GDataRequestException ex )
            {
                textWriter.WriteLine( string.Format( "Error: {0}.", ex.ParseError() ) );
            }
        }

        static void ListGroupsForMember( TextWriter textWriter, AppsService service, string member )
        {
            textWriter.WriteLine( "All groups of which {0} is a direct member:", member );
            textWriter.WriteLine();
            try
            {
                var feed = service.Groups.RetrieveGroups( member, true );
                foreach( AppsExtendedEntry entry in feed.Entries ) //cast Entry to AppsExtendedEntry
                {
                    textWriter.WriteLine( entry.getPropertyValueByName( AppsGroupsNameTable.groupId ) );
                }
            }
            catch( GDataRequestException ex )
            {
                textWriter.WriteLine( string.Format( "Error: {0}.", ex.ParseError() ) );
            }
        }

        static string AddGroupMember( AppsService service, string group, string member )
        {
            try
            {
                var entry = service.Groups.AddMemberToGroup( member, group );
                return string.Format(
                    "{0} added to group, {1}.",
                    entry.getPropertyValueByName( AppsGroupsNameTable.memberId ),
                    group ); //return member from server
            }
            catch( GDataRequestException ex )
            {
                return string.Format( "Error adding group member: {0}.", ex.ParseError() ); //or else return error message
            }
        }

        static void AddGroupMemberInteractiveConsole( AppsService service, string primaryDomain )
        {
            Console.WriteLine();
            Console.WriteLine( "No batch file option (-f). Entering interactive mode." );
            Console.WriteLine( "Press CTRL+C to quit." );
            Console.WriteLine();
            string lastGroup = null;
            while( true ) //continue until CTRL+C
            {
                bool confirm = false;
                string group=null, member=null;
                while( string.IsNullOrEmpty( group ) )
                {
                    if( lastGroup == null ) 
                    { Console.Write( "Group [group@domain]: " ); }
                    else
                    { Console.Write( string.Format( "Group {0} [enter to confirm]: ", lastGroup ) ); }
                    group = Console.ReadLine();
                    if( lastGroup != null && string.IsNullOrEmpty( group ) ) { group = lastGroup; }
                    lastGroup = group;
                }
                while( string.IsNullOrEmpty( member ) )
                {
                    Console.Write( String.Format( "Member to add to {0}: ", group ) );
                    member = Console.ReadLine();
                }
                
                Console.Write( string.Format("Please confirm: Add {0} to {1} (y/n)? ", member, group) );
                confirm = Console.ReadLine().StartsWith( "y", StringComparison.InvariantCultureIgnoreCase );
                if( !confirm )
                {
                    Console.WriteLine( "Cancelled. Group member not added.");
                }
                else
                {
                    try
                    {
                        Console.WriteLine( AddGroupMember( service, group, member ) );
                    }
                    catch( InvalidCredentialsException )
                    {
                        Console.WriteLine();
                        Console.WriteLine( "Invalid Credentials." );
                        Console.WriteLine();
                        //collect new credentials
                        var options = new GoogleAppsAddGroupMemberProgramOptions() { Domain = primaryDomain };
                        CollectOptionsInteractiveConsole( options );
                        service = new AppsService( options.Domain, options.Username, options.Password );
                    }
                }
                Console.WriteLine(); 
            }
        }

        static void AddGroupMemberBatch( TextWriter textWriter, AppsService service, string primaryDomain, string file )
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
                            string group = split[0];
                            string member = split[1];
                            string result = AddGroupMember( service, group, member );
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

        static void CollectOptionsInteractiveConsole( GoogleAppsAddGroupMemberProgramOptions options )
        {
            while( string.IsNullOrEmpty( options.Domain ) )
            {
                Console.Write( "Google Apps Domain or Sub-domain: " );
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

    public sealed class GoogleAppsAddGroupMemberProgramOptions
    {
        public static GoogleAppsAddGroupMemberProgramOptions Parse( string[] args )
        {
            var programOptions = new GoogleAppsAddGroupMemberProgramOptions();

            programOptions.Options = new OptionSet()
            {
                { "d|domain=", "{DOMAIN} containing group", x => programOptions.Domain = x },
                { "u|username=", "{USERNAME}", x => programOptions.Username = x },
                { "p|password=", "{PASSWORD}", x => programOptions.Password = x },
                { "f|file=", "Tab-delimited {FILE} of group and member", x => programOptions.File = x },
                { "l|list", "Print the complete list of groups in the domain and exit.", x => programOptions.List = (x != null ) },
                { "m|list-members=", "Print the list of members in a group and exit.", x => programOptions.ListMembers = x },
                { "g|list-groups-for=", "Print the list of groups for a user and exit", x => programOptions.ListGroupsForMember = x },
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
        public bool List { get; set; }
        public string ListMembers { get; set; }
        public string ListGroupsForMember { get; set; }
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
            textWriter.WriteLine( "Command-line tool to add members to a Google Apps Group." );
            textWriter.WriteLine( "{0} version {1}", appName, Assembly.GetExecutingAssembly().GetName().Version );
            textWriter.WriteLine();
            textWriter.WriteLine( "Options:" );
            Options.WriteOptionDescriptions( textWriter );
        }
    }
}
