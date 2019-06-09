// <copyright file="CommandBase.cs" company="Matt Lacey Ltd.">
// Copyright (c) Matt Lacey Ltd. All rights reserved.
// </copyright>

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
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
            const string LiveXamlPropertyWindowGuid = "{31FC2115-5126-4A87-B2F7-77EAAB65048B}";
            const string LiveXamlTreeWindowGuid = "{A2EAF38F-A0AD-4503-91F8-5F004A69A040}";

            var result = new List<string>();

            try
            {
                async Task<List<string>> GetXamlInfoInternalAsync()
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                    var internalResult = new List<string>();

#pragma warning disable SA1501 // Statement should not be on a single line
                    var dte = (await serviceProvider.GetServiceAsync(typeof(DTE))) as DTE2;
                    if (dte is null) { return new List<string>() { "Unexpected situation: Unable to access DTE2" }; }

                    var proWindow = dte.Windows.Item(LiveXamlPropertyWindowGuid);
                    if (proWindow is null) { return new List<string>() { "Unexpected situation: Unable to access Live Property Window. Check app is running in debug and the Window is visible." }; }

                    var p = typeof(Microsoft.VisualStudio.Platform.WindowManagement.DTE.WindowBase).GetProperty("DockViewElement", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                    if (p is null) { return new List<string>() { "Unexpected situation: Unable to access DockViewElement type." }; }

                    var proDve = p.GetValue(proWindow, null);
                    if (proDve is null) { return new List<string>() { "Unexpected situation: Unable to access DockViewElement window within the Live Property Window." }; }

                    var proViewContent = (proDve as Microsoft.VisualStudio.PlatformUI.Shell.View).Content as System.Windows.Controls.Panel;
                    if (proViewContent is null) { return new List<string>() { "Unexpected situation: Unable to access the content of the DockViewElement window within the Live Property Window." }; }

                    var pvcConPrsntr = proViewContent.Children[0] as System.Windows.Controls.ContentPresenter;
                    if (pvcConPrsntr is null) { return new List<string>() { "Unexpected situation: Unable to access the ContentPresenter in the content of the DockViewElement window within the Live Property Window." }; }

                    var proCntCtrl = pvcConPrsntr.Content as System.Windows.Controls.ContentControl;
                    if (proCntCtrl is null) { return new List<string>() { "Unexpected situation: Unable to access the Control inside the ContentPresenter in the content of the DockViewElement window within the Live Property Window." }; }

                    var pev = proCntCtrl.Content as Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LivePropertyExplorer.View.PropertyExplorerView;
                    if (pev is null) { return new List<string>() { "Unexpected situation: Unable to access LivePropertyExplorer.View.PropertyExplorerView." }; }

                    var proVm = pev.DataContext as Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LivePropertyExplorer.ViewModel.PropertyExplorerViewModel;
                    if (proVm is null) { return new List<string>() { "Unexpected situation: Unable to access LivePropertyExplorer.ViewModel.PropertyExplorerViewModel." }; }

                    var aevm = typeof(Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LivePropertyExplorer.ViewModel.PropertyExplorerViewModel).GetProperty("ActiveElementViewModel");
                    if (aevm is null) { return new List<string>() { "Unexpected situation: Unable to access ActiveElementViewModel." }; }

                    var window = dte.Windows.Item(LiveXamlTreeWindowGuid);
                    if (window is null) { return new List<string>() { "Unexpected situation: Unable to access Live XAML TreeView Window. Check app is running in debug and the Window is visible." }; }

                    var dve = p.GetValue(window, null);
                    if (dve is null) { return new List<string>() { "Unexpected situation: Unable to access DockViewElement window within the Live XAML Tree Window." }; }

                    var x = dve as Microsoft.VisualStudio.Platform.WindowManagement.ToolWindowView;
                    if (x is null) { return new List<string>() { "Unexpected situation: Unable to access the ToolViewWindow for the Live XAML Tree Window." }; }

                    var dc = ((System.Windows.FrameworkElement)x.Content).DataContext;
                    if (dc is null) { return new List<string>() { "Unexpected situation: Unable to access the DataContext of the ToolViewWindow for the Live XAML Tree Window." }; }

                    var dcView = dc as Microsoft.VisualStudio.PlatformUI.Shell.View;
                    if (dcView is null) { return new List<string>() { "Unexpected situation: Unable to access the DataContext of the ToolViewWindow for the Live XAML Tree Window as a Shell.View." }; }

                    var dcViewContent = (System.Windows.Controls.Panel)dcView.Content;
                    if (dcViewContent is null) { return new List<string>() { "Unexpected situation: Unable to access the Content of the DataContext of the ToolViewWindow for the Live XAML Tree Window." }; }

                    var contentPresentr = dcViewContent.Children[1] as System.Windows.Controls.ContentPresenter;
                    if (contentPresentr is null) { return new List<string>() { "Unexpected situation: Unable to access the ContentPresenter of the Content of the DataContext of the ToolViewWindow for the Live XAML Tree Window." }; }

                    var contentCtrl = contentPresentr.Content as System.Windows.Controls.ContentControl;
                    if (contentCtrl is null) { return new List<string>() { "Unexpected situation: Unable to access the internal ContentControl for the Live XAML Tree Window." }; }

                    var contentTree = contentCtrl.Content as Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LiveVisualTree.View.ElementTreeView;
                    if (contentTree is null) { return new List<string>() { "Unexpected situation: Unable to access the LiveVisualTree.View.ElementTreeView." }; }

                    var contentTreeGrid = contentTree.Content as System.Windows.Controls.Grid;
                    if (contentTreeGrid is null) { return new List<string>() { "Unexpected situation: Unable to access the Grid inside the LiveVisualTree.View.ElementTreeView." }; }

                    var treeView = contentTreeGrid.Children[1] as Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LiveVisualTree.View.ElementTreeVirtualizingTreeView;
                    if (treeView is null) { return new List<string>() { "Unexpected situation: Unable to access the LiveVisualTree.View.ElementTreeVirtualizingTreeView." }; }

                    var treeVm = treeView.DataContext as Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LiveVisualTree.ViewModel.ElementTreeViewModel;
                    if (treeVm is null) { return new List<string>() { "Unexpected situation: Unable to access the LiveVisualTree.ViewModel.ElementTreeViewModel." }; }

                    var hostField = treeVm.GetType().GetField("host", BindingFlags.Instance | BindingFlags.NonPublic);
                    if (hostField is null) { return new List<string>() { "Unexpected situation: Unable to access hostField type." }; }

                    var propExplr = new Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LivePropertyExplorer.Model.PropertyExplorer((Microsoft.VisualStudio.DesignTools.Diagnostics.Model.IXamlDiagnosticsHost)hostField.GetValue(treeVm));
                    if (propExplr is null) { return new List<string>() { "Unexpected situation: Unable to access the LivePropertyExplorer.Model.PropertyExplorer." }; }
#pragma warning restore SA1501 // Statement should not be on a single line

                    async Task<List<string>> OutputElement(Microsoft.VisualStudio.DesignTools.Diagnostics.UI.LiveVisualTree.ViewModel.IElementViewModel element, int depth)
                    {
                        var outputElementResult = new List<string>();

                        outputElementResult.Add($"{(depth > 0 ? new string(' ', depth - 1) : string.Empty)}{(depth > 0 ? "- " : string.Empty)}{element.DisplayName}");

                        if (detailLevel != DetailLevel.Outline)
                        {
                            var propObj = await propExplr.ObjectCache.GetElementOrPropertyObjectAsync(element.Element.Handle);

                            var properties = new List<string>();

                            foreach (var source in propObj.Sources)
                            {
                                var sname = source.Source.Name + source.Source.Source.ToString();

                                if (sname != "Default")
                                {
                                    if (detailLevel == DetailLevel.All || sname == "Local")
                                    {
                                        foreach (var sourceProp in source.Properties)
                                        {
                                            // TODO: Complex objects currently show as NULL - get actual details
                                            properties.Add($"{new string(' ', depth)} {sourceProp.Name} = {sourceProp.Value ?? "[null]"}");
                                        }
                                    }
                                }
                            }

                            properties.Sort();
                            outputElementResult.AddRange(properties.Distinct());
                        }

                        foreach (var child in element.Children)
                        {
                            outputElementResult.AddRange(await OutputElement(child, depth + 1));
                        }

                        return outputElementResult;
                    }

                    foreach (var child in treeVm.Children)
                    {
                        internalResult.AddRange(await OutputElement(child, 0));
                    }

                    return internalResult;
                }

                result.AddRange(await GetXamlInfoInternalAsync());
            }
            catch (Exception exc)
            {
                result.Add("Exception");
                result.Add("*********");
                result.Add(exc.Source);
                result.Add(exc.Message);
                result.Add(exc.StackTrace);
                result.Add(string.Empty);
            }

            return result;
        }
    }
}
