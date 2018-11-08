//------------------------------------------------------------------------------
// <copyright file="ShowDefinitionInILSpyCommand.cs" company="Zyonet">
//     Copyright (c) Zyonet.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
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
using Project = EnvDTE.Project;

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
        private DTE _dte;

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

            _dte = ServiceProvider.GetService(typeof (DTE)) as DTE;
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
                Debug.WriteLine($" *** Selected Symbol: {symbolAtCursor.Name} - {symbolAtCursor.Kind} - {symbolAtCursor.ContainingAssembly}");
                var id = symbolAtCursor.GetDocumentationCommentId();
                //ShowInfo(id);
                INamedTypeSymbol namedTypeSymbol = symbolAtCursor as INamedTypeSymbol;

                var assemblySymbol = (namedTypeSymbol != null)?namedTypeSymbol.ContainingAssembly
                    :symbolAtCursor.ContainingAssembly;

                var textView = getTextView();
                var semanticModel = textView.Caret.Position.BufferPosition.Snapshot
                    .GetOpenDocumentInCurrentContextWithChanges().GetSemanticModelAsync().Result;
                bool isProject;
                var assemPath = GetAssemblyPath(semanticModel, assemblySymbol.Identity.ToString(), out isProject);
                var realPath = assemPath;
                Assembly assembly = null;
                try
                {
                    if (!isProject && assemPath != assemblySymbol.Identity.ToString())
                    {
                        Debug.WriteLine($" !!! - loading assembly {assemblySymbol.Identity.Name} ...");
                        assembly = Assembly.ReflectionOnlyLoad(assemblySymbol.Identity.Name);
                        if (assembly != null)
                        {
                            //replace referecend assemblies path with real path
                            realPath = assembly.Location;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    Debug.WriteLine(ex.StackTrace);
                }

                var args = $"\"{realPath}\" /navigateTo:{id} /singleInstance";

                ShowInfo(args);

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
                        SetWindowText(_ilSpyProcess.MainWindowHandle, $"ILSpy for DefILSpy");

                    }
                    else
                    {
                        //send a message to existing instance
                        var msg = $"ILSpy:\r\n{realPath}\r\n/navigateTo:{id}\r\n/singleInstance";
                        Send(_ilSpyProcess.MainWindowHandle, msg);
                    }
                }
            }
            catch (Exception ex)
            {
                ShowInfo(ex.Message + ":" + ex.StackTrace);
                Debug.WriteLine(ex.Message);
                Debug.WriteLine(ex.StackTrace);
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
            var semanticModel = codeDoc.GetSemanticModelAsync().Result;

            SyntaxNode rootNode = codeDoc.GetSyntaxRootAsync().Result;
            var pos = textView.Caret.Position.BufferPosition;

            var st = rootNode.FindToken(pos);

            var curNode = rootNode.FindNode(new Microsoft.CodeAnalysis.Text.TextSpan(pos.Position, 0));
            if (curNode == null)
            {
                curNode = st.Parent;
            }
            var parentKind = st.Parent.Kind();

            //credit: https://github.com/verysimplenick/GoToDnSpy/blob/master/GoToDnSpy/GoToDnSpy.cs 
            //a SyntaxNode is parent of a SyntaxToken
            if (st.Kind() == SyntaxKind.IdentifierToken )
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

            var rs = (localSymbol == null) ? selected : localSymbol.Type;
            if (rs != null)
                return rs;
            else
            {
                return GetSymbolResolvableByILSpy(semanticModel, curNode);
            }
        }

        #region helper

        //https://github.com/icsharpcode/ILSpy/blob/master/ILSpy.AddIn/Commands/OpenCodeItemCommand.cs
        ISymbol GetSymbolResolvableByILSpy(SemanticModel model, SyntaxNode node)
        {
            var current = node;
            while (current != null)
            {
                var symbol = model.GetSymbolInfo(current).Symbol;
                if (symbol == null)
                {
                    symbol = model.GetDeclaredSymbol(current);
                }

                // ILSpy can only resolve some symbol types, so allow them, discard everything else
                if (symbol != null)
                {
                    switch (symbol.Kind)
                    {
                        case SymbolKind.ArrayType:
                        case SymbolKind.Event:
                        case SymbolKind.Field:
                        case SymbolKind.Method:
                        case SymbolKind.NamedType:
                        case SymbolKind.Namespace:
                        case SymbolKind.PointerType:
                        case SymbolKind.Property:
                            break;
                        default:
                            symbol = null;
                            break;
                    }
                }

                if (symbol != null)
                    return symbol;

                current = current is IStructuredTriviaSyntax
                    ? ((IStructuredTriviaSyntax) current).ParentTrivia.Token.Parent
                    : current.Parent;
            }
            return null;
        }

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

        string GetAssemblyPath(SemanticModel semanticModel, string assemblyDef, out bool isProject)
        {
            IEnumerator<AssemblyIdentity> refAsmNames = semanticModel.Compilation.ReferencedAssemblyNames.GetEnumerator();
            IEnumerator<MetadataReference> refs = semanticModel.Compilation.References.GetEnumerator();

            isProject = false;
            var projName = assemblyDef.Split(',').First().Trim();

            // try find in referenced assemblies first
            MetadataReference metReference = null;
            string displayName = "";
            while (refAsmNames.MoveNext())
            {
                refs.MoveNext();
                if (!string.Equals(refAsmNames.Current.ToString(), assemblyDef, StringComparison.OrdinalIgnoreCase))
                    continue;

                displayName = refs.Current.Display;
                metReference = refs.Current;
                if (!assemblyDef.Contains(displayName))
                    //maybe a package path: 
                    return displayName; 
                    //maybe a reference assembly such as "C:\\Program Files (x86)\\Reference Assemblies\\Microsoft\\Framework\\.NETFramework\\v4.5.1\\mscorlib.dll"
            }

            //var assembly = semanticModel.Compilation.GetAssemblyOrModuleSymbol(metReference);
            //symbols defined in other projects of the same solution
            EnvDTE.Project project = null;
            // try found project
            var prjItem = _dte.Solution.FindProjectItem(projName);
            if (prjItem != null)
            {
                project = prjItem.ContainingProject;
            }
            foreach (EnvDTE.Project proj in _dte.Solution.Projects)
            {
                 project = _findProject(proj, projName);
                if (project != null) break;
            }
            //project reference
            if (project != null)
            {
                isProject = true;
               return getProjectOutputPath(project);
            }
            //symbol defined in current project of active document
            EnvDTE.Project curProject = _dte.ActiveDocument.ProjectItem.ContainingProject as EnvDTE.Project;
            if (curProject != null)
            {
                isProject = true;
                return getProjectOutputPath(curProject);
            }

            return assemblyDef;
        }

        private Project _findProject(Project proj, string projName)
        {
            if (proj == null ||  proj.Name == projName)
                return proj;

            Debug.WriteLine(proj.Name);
            var count = proj.ProjectItems.Count;
            Project found = null;
            for (int i = 1; i <= count; ++i)
            {
                ProjectItem nextItem = proj.ProjectItems.Item(i);
                Debug.WriteLine($"==> {nextItem.Name}");
                if (nextItem.ContainingProject != null && nextItem.Name == projName)
                    return nextItem.SubProject;
                else
                {
                    found = _findProject(nextItem.SubProject, projName);
                    if (found != null)
                        return found;
                }
            }
            return null;
        }

        private static string getProjectOutputPath(Project curProject)
        {
            if (curProject != null)
            {
                var config = curProject.ConfigurationManager.ActiveConfiguration;
                foreach (Property p in config.Properties)
                {
                    Debug.WriteLine($"{p.Name}: {p.Value}");
                }

                var prjFolder = Path.GetDirectoryName(curProject.FullName);
                var outPath = Path.Combine(prjFolder, (config.Properties.Item("CodeAnalysisInputAssembly") as Property).Value.ToString());
                return outPath;
            }
            return null;
        }

        #endregion

    }
}