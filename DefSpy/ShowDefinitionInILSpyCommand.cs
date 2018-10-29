//------------------------------------------------------------------------------
// <copyright file="ShowDefinitionInILSpyCommand.cs" company="Zyonet">
//     Copyright (c) Zyonet.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Threading;
using EnvDTE;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Text;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using IServiceProvider = System.IServiceProvider;
using Package = Microsoft.VisualStudio.Shell.Package;
using Process = System.Diagnostics.Process;

namespace DefSpy
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ShowDefinitionInILSpyCommand
    {
        private struct CopyDataStruct
        {
            public IntPtr Padding;

            public int Size;

            public IntPtr Buffer;

            public CopyDataStruct(IntPtr padding, int size, IntPtr buffer)
            {
                Padding = padding;
                Size = size;
                Buffer = buffer;
            }
        }

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0101;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("0B348AFC-8219-41DD-93CE-90AD6396B2EB");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Microsoft.VisualStudio.Shell.Package package;

        private IVsStatusbar _statusBar;

        private static Process _ilSpyProcess;

        [DllImport("user32.dll")]
        private static extern bool SetWindowText(IntPtr hWnd, string text);


        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint msg, IntPtr wParam, ref CopyDataStruct lParam, uint flags, uint timeout, out IntPtr result);

        /// <summary>
        /// Initializes a new instance of the <see cref="ShowDefinitionInILSpyCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private ShowDefinitionInILSpyCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService =
                this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
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
        public static ShowDefinitionInILSpyCommand Instance { get; private set; }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get { return this.package; }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new ShowDefinitionInILSpyCommand(package);
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
                var symbolAtCursor = getChosenSymbol();
                Debug.Write($"Selected Symbol: {symbolAtCursor.Name} - {symbolAtCursor.Kind} - {symbolAtCursor.ContainingAssembly}");
                var id = symbolAtCursor.GetDocumentationCommentId();
                ShowInfo(id);
                INamedTypeSymbol namedTypeSymbol = symbolAtCursor as INamedTypeSymbol;

                var assemblySymbol = (namedTypeSymbol != null)?namedTypeSymbol.ContainingAssembly
                    :symbolAtCursor.ContainingAssembly;

                var assembly = Assembly.Load(assemblySymbol.Identity.Name);
                var args = $"\"{assembly.Location}\" /navigateTo:{id} /singleInstance";
                if (string.IsNullOrEmpty(DefSpy.Default.ILSpyPath))
                {
                    SetILPathCommand.GetILSpyPath();
                }
                if (string.IsNullOrEmpty(DefSpy.Default.ILSpyPath))
                {
                    ShowInfo("Please specify path to ILSpy.exe first!");
                }
                else
                {
                    if (_ilSpyProcess == null || _ilSpyProcess.HasExited)
                    {
                        if (_ilSpyProcess != null)
                        {
                            _ilSpyProcess.Dispose();
                        }
                        Trace.WriteLine($"Launching {args}");
                        _ilSpyProcess = Process.Start(DefSpy.Default.ILSpyPath, args);
                    }
                    else
                    {
                        //send a message to existing instance
                        var msg = $"ILSpy:\r\n{assembly.Location}\r\n/navigateTo:{id}\r\n/singleInstance";
                        Send(_ilSpyProcess.MainWindowHandle, msg);
                    }
                }
            }
            catch (Exception ex)
            {
                ShowInfo(ex.Message);
                Trace.TraceError(ex.Message);
            }
        }

        private static unsafe IntPtr Send(IntPtr hWnd, string message)
        {
            CopyDataStruct lParam = default(CopyDataStruct);
            lParam.Padding = IntPtr.Zero;
            lParam.Size = message.Length * 2;
            fixed (char* value = message)
            {
                IntPtr result = IntPtr.Zero;
                lParam.Buffer = (IntPtr)(void*)value;
                //WM_COPYDATA
                if (SendMessageTimeout(hWnd, 74u, IntPtr.Zero, ref lParam, 0u, 3000u, out result) != IntPtr.Zero)
                {
                    return result;
                }
                return IntPtr.Zero;
            }
        }
        /// <summary>
        /// Get the CodeAnalysis Document object, which is different from VS shell document
        /// </summary>
        /// <returns></returns>
        private IWpfTextView getTextView()
        {
            IVsTextView textView = null;
            var textMgr = ServiceProvider.GetService(typeof(SVsTextManager)) as IVsTextManager;
            textMgr.GetActiveView(1, null, out textView);

            var comModel = ServiceProvider.GetService(typeof(SComponentModel))
                as IComponentModel;
            var editorFactory = comModel.GetService<IVsEditorAdaptersFactoryService>();

            IWpfTextView view2 = editorFactory.GetWpfTextView(textView);

            return view2;
        }

        /// <summary>
        /// Get the Type of the selected text. e.g. for a local variable, we need to get the type of the
        /// variable instead of the variable itself
        /// </summary>
        /// <returns></returns>
        private ISymbol getChosenSymbol()
        {
            ISymbol selected = null;
            var textView = getTextView();
            Microsoft.CodeAnalysis.Document codeDoc = textView.Caret.Position.BufferPosition.Snapshot
               .GetOpenDocumentInCurrentContextWithChanges(); //textfeature extension method
            var pos = textView.Caret.Position.BufferPosition;

            SyntaxNode rootNode = codeDoc.GetSyntaxRootAsync().Result;
            SyntaxToken st = rootNode.FindToken(pos);
            var semanticModel = codeDoc.GetSemanticModelAsync().Result;
            var parentKind = st.Parent.Kind();

            //credit: https://github.com/verysimplenick/GoToDnSpy/blob/master/GoToDnSpy/GoToDnSpy.cs 
            //a SyntaxNode is parent of a SyntaxToken
            if (st.Kind() == SyntaxKind.IdentifierToken && (
                       parentKind == SyntaxKind.PropertyDeclaration
                    || parentKind == SyntaxKind.FieldDeclaration
                    || parentKind == SyntaxKind.MethodDeclaration
                    || parentKind == SyntaxKind.NamespaceDeclaration
                    || parentKind == SyntaxKind.DestructorDeclaration
                    || parentKind == SyntaxKind.ConstructorDeclaration
                    || parentKind == SyntaxKind.OperatorDeclaration
                    || parentKind == SyntaxKind.ConversionOperatorDeclaration
                    || parentKind == SyntaxKind.EnumDeclaration
                    || parentKind == SyntaxKind.EnumMemberDeclaration
                    || parentKind == SyntaxKind.ClassDeclaration
                    || parentKind == SyntaxKind.EventDeclaration
                    || parentKind == SyntaxKind.EventFieldDeclaration
                    || parentKind == SyntaxKind.InterfaceDeclaration
                    || parentKind == SyntaxKind.StructDeclaration
                    || parentKind == SyntaxKind.DelegateDeclaration
                    || parentKind == SyntaxKind.IndexerDeclaration
                    || parentKind == SyntaxKind.VariableDeclarator
                    ))
            {
                selected = semanticModel.LookupSymbols(pos.Position, name: st.Text).FirstOrDefault();
            }
            else
            {
                var symbolInfo = semanticModel.GetSymbolInfo(st.Parent);
                selected = symbolInfo.Symbol ?? (symbolInfo.GetType().GetProperty("CandidateSymbols")
                    .GetValue(symbolInfo) as IEnumerable<ISymbol>)?.FirstOrDefault();
            }

            var localSymbol = selected as ILocalSymbol;

            return (localSymbol == null) ? selected : localSymbol.Type;
        }

        #region helper
        void ShowInfo(string text)
        {
            Dispatcher.CurrentDispatcher.VerifyAccess();
            if (_statusBar == null)
                try
                {
                    _statusBar = ServiceProvider.GetService(typeof(SVsStatusbar)) as IVsStatusbar;
                }
                catch
                {
                    MessageBox.Show(text);
                    return;
                }

            _statusBar.SetText(text);
        }
        #endregion

    }
}