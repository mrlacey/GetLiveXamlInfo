// <copyright file="GetLiveXamlInfoPackage.cs" company="Matt Lacey Ltd.">
// Copyright (c) Matt Lacey Ltd. All rights reserved.
// </copyright>

using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace GetLiveXamlInfo
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(GetLiveXamlInfoPackage.PackageGuidString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class GetLiveXamlInfoPackage : AsyncPackage
    {
        public const string PackageGuidString = "42ff7b4b-0ed9-407b-aa25-3397606f6268";

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            await GetLocalPropertiesCommand.InitializeAsync(this);
            await GetAllPropertiesCommand.InitializeAsync(this);
            await GetElementOutlineCommand.InitializeAsync(this);
        }
    }
}
