//------------------------------------------------------------------------------
// <copyright file="TextEditorContextMenuCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using System.Web;
using LibGit2Sharp;
using System.IO;

namespace CodeSearch
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class TextEditorContextMenuCommand
    {
        private const string visualStudioBaseURL = "visualstudio.com";
        private const string gitHubBaseURL = "github.com";

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("55ed39de-e8f3-4e9c-84c8-e5d62911446e");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="TextEditorContextMenuCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private TextEditorContextMenuCommand(Package package)
        {
            try
            { 
                if (package == null)
                {
                    throw new ArgumentNullException("package");
                }

                this.package = package;

                OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
                if (commandService != null)
                {
                    var menuCommandID = new CommandID(CommandSet, CommandId);
                    var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                    commandService.AddCommand(menuItem);
                }
            }
            catch (Exception ex)
            {
                string message = string.Format(CultureInfo.CurrentCulture, ex.Message + ex.StackTrace, this.GetType().FullName);
                string title = "Code Search Extension TextEditorContextMenuCommand Error";

                VsShellUtilities.ShowMessageBox(
                    this.ServiceProvider,
                    message,
                    title,
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static TextEditorContextMenuCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new TextEditorContextMenuCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            try
            {
                DTE dte = (DTE)this.ServiceProvider.GetService(typeof(DTE));
                var sourceControlURL = string.Empty;

                // 1.) Get TFS URL
                //Plan A - Team Explorer Reflection
                try
                {
                    /****** Reflection Hacks to access private DLLs for both Visual Studio 2015 and Visual Studio 2017 ******/
                    var ext = dte.GetObject("Microsoft.VisualStudio.TeamFoundation.TeamFoundationServerExt");

                    System.Type type = ext.GetType();
                    System.Reflection.PropertyInfo property = type.GetProperty("ActiveProjectContext");
                    object activeProjectContext = property.GetValue(ext);

                    System.Type type2 = activeProjectContext.GetType();
                    System.Reflection.PropertyInfo property2 = type2.GetProperty("DomainUri");
                    object domainUri = property2.GetValue(activeProjectContext);

                    sourceControlURL = domainUri.ToString();
                    /********************************************************************************************************/
                }
                catch (Exception ex)
                {
                    //Swallowing this exception on purpose
                }

                //Get TFS or GitHub URL - Plan B - LibGit2Sharp
                if (string.IsNullOrEmpty(sourceControlURL))
                {
                    var directory = new DirectoryInfo(dte.ActiveDocument.Path);
                    while (directory.Parent != null && directory.Parent.Name != directory.Root.Name)
                    {
                        if (Repository.IsValid(directory.FullName))
                        {
                            using (var repo = new Repository(directory.FullName))
                            {
                                var origin = repo.Network.Remotes["origin"].Url;
                                var uri = new Uri(origin);
                                var root = uri.GetLeftPart(UriPartial.Authority);
                                var projectCollection = uri.Segments[1];

                                if (root.Contains(visualStudioBaseURL)) //VisualStudioOnline
                                {
                                    sourceControlURL = String.Format("{0}/{1}", root, projectCollection);
                                }
                                else if (root.Contains(gitHubBaseURL)) //GitHub
                                {
                                    var gitHubRepo = uri.Segments[2].Replace(".git", string.Empty);
                                    sourceControlURL = String.Format("{0}/{1}/{2}", root, projectCollection, gitHubRepo);
                                }

                                break;
                            }
                        }

                        directory = directory.Parent;
                    }
                }

                //2.) Open Browser with Search URL
                var selection = (TextSelection)dte.ActiveDocument.Selection;
                var encodedSelection = HttpUtility.UrlEncode(selection.Text);
                var searchURL = string.Empty;

                if (sourceControlURL.Contains(visualStudioBaseURL)) //VisualStudioOnline
                {
                    searchURL = string.Format("{0}/_search?type=Code&lp=search-project&text={1}&preview=1&_a=contents", sourceControlURL, encodedSelection);
                }
                else if (sourceControlURL.Contains(gitHubBaseURL)) //GitHub
                {
                    searchURL = string.Format("{0}/search?q={1}", sourceControlURL, encodedSelection);
                }

                System.Diagnostics.Process.Start(searchURL);
            }
            catch (Exception ex)
            {
                ShowError(string.Format(CultureInfo.CurrentCulture, ex.Message + ex.StackTrace, this.GetType().FullName));
            }
        }

        private void ShowError(string message)
        {
            VsShellUtilities.ShowMessageBox(
                this.ServiceProvider,
                string.Format(CultureInfo.CurrentCulture, message, this.GetType().FullName),
                "Code Search Extension",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
