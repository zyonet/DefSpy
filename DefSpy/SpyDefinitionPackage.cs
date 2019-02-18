﻿//------------------------------------------------------------------------------
// <copyright file="SpyDefinitionCommandPackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace DefSpy
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(SpyDefinitionPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class SpyDefinitionPackage : AsyncPackage
    {
        /// <summary>
        /// SpyDefinitionCommandPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "6bcb02d7-fb93-42c9-8684-6bebe3e3f45a";

        /// <summary>
        /// Initializes a new instance of the <see cref="SetILPathCommand"/> class.
        /// </summary>
        public SpyDefinitionPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        #region Package Members

        protected override Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            SetILPathCommand.Initialize(this);
            ShowDefinitionInILSpyCommand.Initialize(this);
            return base.InitializeAsync(cancellationToken, progress);
        }

        #endregion

        #region helper

        public void ShowMessage(string format, params object[] items)
        {
            ShowMessage(OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST, OLEMSGICON.OLEMSGICON_INFO, format, items);
        }

        public void ShowMessage(OLEMSGICON icon, string format, params object[] items)
        {
            ShowMessage(OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST, icon, format, items);
        }

        public int ShowMessage(OLEMSGBUTTON buttons, OLEMSGDEFBUTTON defaultButton, OLEMSGICON icon, string format, params object[] items)
        {
            IVsUIShell uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            Guid clsid = Guid.Empty;
            int result;
            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(
                uiShell.ShowMessageBox(
                    0,
                    ref clsid,
                    "ILSpy AddIn",
                    string.Format(CultureInfo.CurrentCulture, format, items),
                    string.Empty,
                    0,
                    buttons,
                    defaultButton,
                    icon,
                    0,        // false
                    out result
                )
            );

            return result;
        }



        #endregion
    }
}
