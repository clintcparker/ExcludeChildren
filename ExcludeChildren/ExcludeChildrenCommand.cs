//------------------------------------------------------------------------------
// <copyright file="ExcludeChildrenCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Microsoft.Internal.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace ExcludeChildren
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ExcludeChildrenCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("07ecea7b-bb77-4a0c-8bdc-eb48b83edd7a");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcludeChildrenCommand"/> class. Adds our
        /// command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private ExcludeChildrenCommand(Package package)
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

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ExcludeChildrenCommand Instance
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
            Instance = new ExcludeChildrenCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">     Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var dte = (DTE2)ServiceProvider.GetService(typeof(DTE));

            var itemsToExclude = GetSelectedItems(dte);

            ExcludeItems(itemsToExclude);
        }

        private static void ExcludeItems(IEnumerable<ProjectItem> itemsToExclude)
        {
            itemsToExclude.ToList().ForEach(x =>
            {
                x.Remove();
            });
        }

        private static IEnumerable<ProjectItem> GetSelectedItems(DTE2 dte)
        {
            var items1 = (Array)dte.ToolWindows.SolutionExplorer.SelectedItems;

            return items1.Cast<UIHierarchyItem>().SelectMany(GetProjectItems);
        }

        private static IEnumerable<ProjectItem> GetProjectItems(UIHierarchyItem uihitem)
        {
            var hasNoSubBranches = uihitem.UIHierarchyItems.Count == 0 ||
                                   !uihitem.UIHierarchyItems.Cast<UIHierarchyItem>().Any(x => x.Object is ProjectItem);

            return hasNoSubBranches ? FakeYield(uihitem.Object as ProjectItem) : uihitem.UIHierarchyItems.Cast<UIHierarchyItem>().SelectMany(GetProjectItems);
        }

        private static IEnumerable<T> FakeYield<T>(T obj)
        {
            yield return obj;
        }
    }
}