// <copyright file="CommandBase.cs" company="Matt Lacey Ltd.">
// Copyright (c) Matt Lacey Ltd. All rights reserved.
// </copyright>

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace GetLiveXamlInfo
{
    internal class CommandBase
    {
        public async Task<List<string>> GetXamlInfoAsync(Microsoft.VisualStudio.Shell.IAsyncServiceProvider serviceProvider, DetailLevel detailLevel)
        {
            var result = new List<string>();

            try
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                // TODO: make more robust - with useful error messages
                // TODO: filter based on requested detail level
                var dte = (DTE2)(await serviceProvider.GetServiceAsync(typeof(DTE)));

                var proWindow = dte.Windows.Item("{31FC2115-5126-4A87-B2F7-77EAAB65048B}");

                var p = typeof(Microsoft.VisualStudio.Platform.WindowManagement.DTE.WindowBase).GetProperty("DockViewElement", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

                var proDve = p.GetValue(proWindow, null);

                var proViewContent = (proDve as Microsoft.VisualStudio.PlatformUI.Shell.View).Content as System.Windows.Controls.Panel;

                var pvcConPrsntr = proViewContent.Children[0] as System.Windows.Controls.ContentPresenter;

                var proCntCtrl = pvcConPrsntr.Content as System.Windows.Controls.ContentControl;

                var pev = proCntCtrl.Content as Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LivePropertyExplorer.View.PropertyExplorerView;

                Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LivePropertyExplorer.ViewModel.PropertyExplorerViewModel proVm = pev.DataContext as Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LivePropertyExplorer.ViewModel.PropertyExplorerViewModel;

                var aevm = typeof(Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LivePropertyExplorer.ViewModel.PropertyExplorerViewModel).GetProperty("ActiveElementViewModel");

                var window = dte.Windows.Item("{A2EAF38F-A0AD-4503-91F8-5F004A69A040}");

                var dve = p.GetValue(window, null);

                Microsoft.VisualStudio.Platform.WindowManagement.ToolWindowView x = dve as Microsoft.VisualStudio.Platform.WindowManagement.ToolWindowView;

                var dc = ((System.Windows.FrameworkElement)x.Content).DataContext;

                var dcView = dc as Microsoft.VisualStudio.PlatformUI.Shell.View;

                var dcViewContent = (System.Windows.Controls.Panel)dcView.Content;

                var contentPresentr = dcViewContent.Children[1] as System.Windows.Controls.ContentPresenter;

                var contentCtrl = contentPresentr.Content as System.Windows.Controls.ContentControl;

                var contentTree = contentCtrl.Content as Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LiveVisualTree.View.ElementTreeView;

                var contentTreeGrid = contentTree.Content as System.Windows.Controls.Grid;

                var treeView = contentTreeGrid.Children[1] as Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LiveVisualTree.View.ElementTreeVirtualizingTreeView;

                var treeVm = treeView.DataContext as Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LiveVisualTree.ViewModel.ElementTreeViewModel;

                var hostField = treeVm.GetType().GetField("host", BindingFlags.Instance | BindingFlags.NonPublic);

                var propExplr = new Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LivePropertyExplorer.Model.PropertyExplorer((Microsoft.VisualStudio.DesignTools.Diagnostics.Model.IXamlDiagnosticsHost)hostField.GetValue(treeVm));

                // TODO: sort properties into alphabetical order and remove duplicates.
                async Task OutputElement(Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LiveVisualTree.ViewModel.IElementViewModel element, int depth)
                {
                    Debug.WriteLine($"{(depth > 0 ? new string(' ', depth - 1) : string.Empty)}{(depth > 0 ? "- " : string.Empty)}{element.DisplayName}");

                    var propObj = await propExplr.ObjectCache.GetElementOrPropertyObjectAsync(element.Element.Handle);

                    foreach (var source in propObj.Sources)
                    {
                        var sname = source.Source.Name + source.Source.Source.ToString();

                        if (sname != "Default")
                        {
                            foreach (var sourceProp in source.Properties)
                            {
                                Debug.WriteLine($"{new string(' ', depth)} {sourceProp.Name} = {sourceProp.Value ?? "[null]"}");
                            }
                        }
                    }

                    var props = await proVm.GetSelectedElementPropertiesAsync();

                    foreach (var child in element.Children)
                    {
                        await OutputElement(child, depth + 1);
                    }
                }

                foreach (var child in treeVm.Children)
                {
                    await OutputElement(child, 0);
                }
            }
            catch (Exception exc)
            {
                // TODO: add useful error details
                Debug.WriteLine(exc);
            }

            return result;
        }
    }
}
