using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

using EnvDTE;
using Microsoft.VisualStudio;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace StyleExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class FunctionHeader
    {
        //--------------------------------------------------------------------------
        // Fields
        //--------------------------------------------------------------------------
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("38dfd4d0-7b42-4360-a7f0-2095537a05e2");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Regular expressions to match function signatures
        /// </summary>
        private readonly List<Regex> regexes = new List<Regex>();

        /// <summary>
        /// Regular expression for a partial function signature match, 
        /// used when function spans multiple lines
        /// </summary>
        private readonly Regex startOfFunction;

        //--------------------------------------------------------------------------
        // Constructors
        //--------------------------------------------------------------------------
        /// <summary>
        /// Initializes a new instance of the <see cref="FunctionHeader"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private FunctionHeader(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            // Add the execute callback to the menu item
            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);

            // Function regex
            regexes.Add(new Regex(@"^(?<returnValue>[\w\*]+)\s+(?<className>\w+)::(?<functionName>\w+)\((?<signature>.*)\)\s*(const)?$", RegexOptions.Compiled | RegexOptions.IgnoreCase));

            // Constructor regex
            regexes.Add(new Regex(@"^(?<className>\w+)::(?<functionName>\w+)\((?<signature>.*)\)\s*$", RegexOptions.Compiled | RegexOptions.IgnoreCase));

            // Partial function match (start of a function)
            startOfFunction = new Regex(@"^(?<returnValue>[\w\*]+)\s+(?<className>\w+)::(?<functionName>\w+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        }

        //--------------------------------------------------------------------------
        // Properties
        //--------------------------------------------------------------------------
        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static FunctionHeader Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        //--------------------------------------------------------------------------
        // Methods
        //--------------------------------------------------------------------------
        /// <summary>
        /// Generates a function header block based off of a match object
        /// </summary>
        /// <param name="match">Regex match object, must be a match for a function signature</param>
        /// <returns></returns>
        private string GenerateFunctionHeader(Match match)
        {
            // Extract match groups
            string functionName = match.Groups["functionName"].Value;
            string className    = match.Groups["className"].Value;
            string signature    = match.Groups["signature"].Value;
            string returnValue  = match.Groups["returnValue"].Value == "" ? "void" : match.Groups["returnValue"].Value;

            // Begin header block
            string functionHeader = "/**-------------------------------------------------------------\n";

            // Automatically fill out function brief for specific method
            if(functionName == className)
                functionHeader = functionHeader + "* @brief Default constructor\n";
            else if (functionName.Contains("Tick"))
                functionHeader = functionHeader + "* @brief Called every frame\n";
            else if (functionName == "BeginPlay")
                functionHeader = functionHeader + "* @brief Called once actor has been spawned into world\n";
            else if (functionName == "EndPlay")
                functionHeader = functionHeader + "* @brief Called when actor is being removed from world\n";
            else if (functionName == "OnConstruction")
                functionHeader = functionHeader + "* @brief Called after spawning actor but before play\n";
            else
                functionHeader = functionHeader + "* @brief \n";

            // Add param values, if there are any
            string[] paramValues = signature.Split(',');
            if (paramValues.Length > 1 || (paramValues.Length == 1 && paramValues[0] != ""))
            {
                // Asterisk between brief and params
                functionHeader = functionHeader + "*\n";

                // Loop over each param in the array
                foreach (var param in paramValues)
                {
                    // Remove whitespace
                    string newParam = param.Trim();

                    // Loop over character array (converted from the string) starting from the end,
                    // and break when the first space is found. This should give us the param name
                    int i;
                    char[] charArray = newParam.ToCharArray();
                    for (i = charArray.Length - 1; i >= 0; i--) if (charArray[i] is ' ') break;

                    // Add the param to the header block
                    functionHeader = functionHeader + String.Format("* @param {0} \n", newParam.Substring(i + 1));
                }
            }

            // Add return value, if the return is not void
            if (returnValue != "void")
            {
                functionHeader = functionHeader + "*\n";
                functionHeader = functionHeader + "* @return \n";
            }

            // Finish the header block
            functionHeader = functionHeader + "*/\n";

            return functionHeader;
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in FunctionHeader's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new FunctionHeader(package, commandService);
        }

        /// <summary>
        /// Checks if the line of text matches a function regex, if it 
        /// does, a header block will be inserted for the function
        /// </summary>
        /// <param name="line">Line of text to check</param>
        /// <param name="insertPoint">Point in file to insert header block</param>
        /// <returns></returns>
        private bool ValidateAndInsert(string line, EditPoint insertPoint)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Check if string is a function signature
            foreach (var regex in regexes)
            {
                // Use regex matching to check for function signature
                Match match = regex.Match(line);
                if (match.Success)
                {
                    // Generate function header and insert into the code
                    string functionHeader = GenerateFunctionHeader(match);
                    insertPoint.Insert(functionHeader);
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private async void Execute(object sender, EventArgs e)
        {
            // Switch to UI thread
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            
            // Get the Development Tools Extensibility service
            DTE dte = await ServiceProvider.GetServiceAsync(typeof(DTE)) as DTE;
            if (dte != null && dte.ActiveDocument != null)
            {
                // Operation succeeded
                bool success = false;

                // Get the text selection
                var selection = dte.ActiveDocument.Selection as TextSelection;
                if (selection is null) return;

                // Find the start and end of the line
                EditPoint startPoint = selection.AnchorPoint.CreateEditPoint();
                EditPoint endPoint   = selection.AnchorPoint.CreateEditPoint();
                startPoint.StartOfLine();
                endPoint.EndOfLine();

                // Denote insertion location
                EditPoint insertPoint = startPoint.CreateEditPoint();

                // Get the line as a string
                string line = startPoint.GetText(endPoint);

                // Check if string is a function signature, if it is insert the header
                success = ValidateAndInsert(line, insertPoint);

                // If a function header wasn't inserted, check if the 
                // function signature is spread across multiple lines
                if (!success && startOfFunction.Match(line).Success)
                {
                    // Count all parentheses on the starting line
                    int openingCount = line.Length - line.Replace("(", "").Length;
                    int closingCount = line.Length - line.Replace(")", "").Length;
                    int parenDiff = openingCount - closingCount;

                    // Copy starting position
                    string multiLine = line;

                    // Move the active point to the anchor point
                    selection.MoveToPoint(selection.ActivePoint);
                    while (true)
                    {
                        // Move down the document concatenating the lines onto the starting line
                        selection.LineDown();
                        startPoint = selection.AnchorPoint.CreateEditPoint();
                        endPoint = selection.AnchorPoint.CreateEditPoint();
                        startPoint.StartOfLine();
                        endPoint.EndOfLine();
                        string nextLine = startPoint.GetText(endPoint).Trim();
                        multiLine += nextLine;

                        // Count all parentheses on this line, and add them to the diff
                        openingCount = nextLine.Length - nextLine.Replace("(", "").Length;
                        closingCount = nextLine.Length - nextLine.Replace(")", "").Length;
                        parenDiff += openingCount - closingCount;

                        // Exit if there are more closing parentheses than opening, or if there is a matching set
                        // This will cause the while loop to exit if an opening parenthesis is not detected within
                        // one line after the line we start on
                        if (parenDiff <= 0) break;

                        // Exit if we reached the end of the document
                        if (endPoint.AtEndOfDocument) break;
                    }
        
                    // Check if string is a function signature, if it is insert the header
                    success = ValidateAndInsert(multiLine, insertPoint);
                }

                // The selected line wasn't a function header, notify user in output
                if (!success)
                {
                    // Get the debug output pane
                    IVsOutputWindow outputWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;
                    Guid paneGuid = VSConstants.GUID_OutWindowDebugPane;
                    IVsOutputWindowPane pane;
                    outputWindow.GetPane(ref paneGuid, out pane);
                    if (pane is null) return;

                    // Print the line that the was not a function signature, and then show the debug output
                    pane.OutputString(String.Format("The following line is not a function signature:\n\t{0}\n", line));
                    pane.Activate();
                }
            }
        }
    }
}
