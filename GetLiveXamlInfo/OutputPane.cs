// <copyright file="OutputPane.cs" company="Matt Lacey Ltd.">
// Copyright (c) Matt Lacey Ltd. All rights reserved.
// </copyright>

using System;
using System.Collections.Generic;
using System.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace GetLiveXamlInfo
{
    public class OutputPane
    {
        private static Guid dsPaneGuid = new Guid("F9C8C9D9-BA46-4C38-BB73-76F99E07C2BD");

        private static OutputPane instance;

        private readonly IVsOutputWindowPane pane;
        private bool firstThingBeingWritten = true;

        private OutputPane()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (ServiceProvider.GlobalProvider.GetService(typeof(SVsOutputWindow)) is IVsOutputWindow outWindow
                && (ErrorHandler.Failed(outWindow.GetPane(ref dsPaneGuid, out this.pane)) || this.pane == null))
            {
                if (ErrorHandler.Failed(outWindow.CreatePane(ref dsPaneGuid, "Live XAML Info", 1, 0)))
                {
                    System.Diagnostics.Debug.WriteLine("Failed to create output pane.");
                    return;
                }

                if (ErrorHandler.Failed(outWindow.GetPane(ref dsPaneGuid, out this.pane)) || (this.pane == null))
                {
                    System.Diagnostics.Debug.WriteLine("Failed to get output pane.");
                }
            }
        }

        public static OutputPane Instance => instance ?? (instance = new OutputPane());

        public async Task ActivateAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(CancellationToken.None);

            this.pane?.Activate();
        }

        public async Task WriteAsync(string message)
        {
            if (this.firstThingBeingWritten)
            {
                this.firstThingBeingWritten = false;

                await OutputPane.Instance.WriteAsync("This is an Experimental extension. Learn more at https://github.com/mrlacey/GetLiveXamlInfo ");
                await OutputPane.Instance.WriteAsync("If you have problems with this extension, or suggestions for improvement, report them at https://github.com/mrlacey/GetLiveXamlInfo/issues/new ");
                //// await OutputPane.Instance.WriteAsync("If you like this extension please leave a review at https://marketplace.visualstudio.com/items?itemName=MattLaceyLtd.GetLiveXamlInfo#review-details ");
                await OutputPane.Instance.WriteAsync(string.Empty);
            }

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(CancellationToken.None);

            this.pane?.OutputString($"{message}{Environment.NewLine}");
        }

        public async Task WriteStringsAsync(List<string> messages)
        {
            foreach (var msg in messages)
            {
                await this.WriteAsync(msg);
            }

            await this.WriteAsync(string.Empty);

            await this.ActivateAsync();
        }
    }
}
