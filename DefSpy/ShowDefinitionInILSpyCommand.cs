﻿//------------------------------------------------------------------------------
// <copyright file="ShowDefinitionInILSpyCommand.cs" company="Zyonet">
//     Copyright (c) Zyonet.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Threading;
using EnvDTE;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.LanguageServices;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Package = Microsoft.VisualStudio.Shell.Package;
using Process = System.Diagnostics.Process;
using Project = EnvDTE.Project;

namespace DefSpy
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ShowDefinitionInILSpyCommand:ILSpyCommand
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
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Microsoft.VisualStudio.Shell.Package package;

        private IVsStatusbar _statusBar;

        private static Process _ilSpyProcess;
        private DTE _dte;
        private VisualStudioWorkspace workspace;
        //private Microsoft.VisualStudio.LanguageServices.RoslynVisualStudioWorkspace workspace;

        [DllImport("user32.dll")]
        private static extern bool SetWindowText(IntPtr hWnd, string text);


        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint msg, IntPtr wParam, ref CopyDataStruct lParam, uint flags, uint timeout, out IntPtr result);

        /// <summary>
        /// Initializes a new instance of the <see cref="ShowDefinitionInILSpyCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private ShowDefinitionInILSpyCommand(SpyDefinitionPackage package):base(package, CommandId)
        {

            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            //OleMenuCommandService commandService =
            //    this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            //if (commandService != null)
            //{
            //    var menuCommandID = new CommandID(ILSpyCommand.CommandSet, CommandId);
            //    var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
            //    commandService.AddCommand(menuItem);
            //}

            _dte = Package.GetGlobalService(typeof (DTE)) as DTE;

            var componentModel = (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
            this.workspace = componentModel.GetService<VisualStudioWorkspace>() as VisualStudioWorkspace;


        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ShowDefinitionInILSpyCommand Instance { get; private set; }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
       

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(SpyDefinitionPackage package)
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
        private async void MenuItemCallback(object sender, EventArgs e)
        {
            try
            {

                var roslynDoc = GetRoslynDocument();
                var semanticModel = await roslynDoc.GetSemanticModelAsync();

                var symbolAtCursor = getChosenSymbol();
                Debug.WriteLine($" *** Selected Symbol: {symbolAtCursor?.Name} - {symbolAtCursor?.Kind} - {symbolAtCursor?.ContainingAssembly}");

                if (symbolAtCursor == null)
                {
                    ShowInfo("Cannot find symbol selected. Returning...");
                }

                var roslynProject = roslynDoc.Project;

                var id = (symbolAtCursor.OriginalDefinition ?? symbolAtCursor)
                    .GetDocumentationCommentId();
                //ShowInfo(id);

                INamedTypeSymbol namedTypeSymbol = symbolAtCursor as INamedTypeSymbol;
                var alias = symbolAtCursor as IAliasSymbol;
                if (alias != null)
                {
                    namedTypeSymbol = alias.Target as INamedTypeSymbol;
                    //id = namedTypeSymbol.GetDocumentationCommentId();
                }

                var assemblySymbol = (namedTypeSymbol != null)?namedTypeSymbol.ContainingAssembly
                    :symbolAtCursor?.ContainingAssembly;



                bool isProject;
                var fullName = assemblySymbol?.ToDisplayString();
                var assemPath = GetAssemblyPath(semanticModel, fullName, out isProject);
                var realPath = assemPath;
                Assembly assembly = null;
                try
                {
                    if (assemPath.Contains("packages") || isProject)
                    {
                        assembly = Assembly.ReflectionOnlyLoad(assemPath);
                    }
                    else 
                    {
                        Debug.WriteLine($" !!! - loading assembly {assemblySymbol?.Identity.Name} ...");
                        assembly = Assembly.Load(fullName); //such as System
                    }

                    if (assembly != null)
                    {
                        //replace referecend assemblies path with real path
                        //e.g. C:\Windows\Microsoft.NET\Framework\v4.0.30319\mscorlib.dll
                        assemPath = assembly.Location;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    Debug.WriteLine(ex.StackTrace);
                }

                var args = "";
                if (assemPath.Contains("assembly\\GAC_"))
                     args = $"\"{assemPath}\" /navigateTo:{id} /singleInstance";
                else
                {
                    args = $"\"{realPath}\" /navigateTo:{id} /singleInstance";
                }

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
            var textMgr = owner.GetServiceAsync(typeof(SVsTextManager)).Result as IVsTextManager;
            textMgr.GetActiveView(1, null, out textView);

            IVsUserData userData = textView as IVsUserData;
            if (userData == null)
                return null;

            var comModel = owner.GetServiceAsync(typeof(SComponentModel)).Result
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
            Microsoft.CodeAnalysis.Document codeDoc = GetRoslynDocument();
            //textfeature extension method
            var semanticModel = codeDoc.GetSemanticModelAsync().Result;

            SyntaxNode rootNode = codeDoc.GetSyntaxRootAsync().Result;
            var pos = textView.Caret.Position.BufferPosition;

            var st = rootNode.FindToken(pos);

            var curNode = rootNode.FindNode(
                new Microsoft.CodeAnalysis.Text.TextSpan(pos.Position, 0));

            if (curNode == null)
            {
                curNode = st.Parent;
            }

            var sym = GetSymbolResolvableByILSpy(semanticModel, curNode);
            if (sym != null)
                return sym;
            
            var parentKind = st.Parent.RawKind;

            //credit: https://github.com/verysimplenick/GoToDnSpy/blob/master/GoToDnSpy/GoToDnSpy.cs 
            //a SyntaxNode is parent of a SyntaxToken
            if (st.RawKind== (int) SyntaxKind.IdentifierToken )
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
            return rs;
        }

        #region helper
        Microsoft.CodeAnalysis.Document GetRoslynDocument()

        {
            var document = this._dte.ActiveDocument;
            var selection = (EnvDTE.TextPoint)((EnvDTE.TextSelection)document.Selection).ActivePoint;
            var id = this.workspace.CurrentSolution.GetDocumentIdsWithFilePath(document.FullName).FirstOrDefault();
            if (id == null)
                return null;

            return this.workspace.CurrentSolution.GetDocument(id);
        }

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
                    _statusBar = owner.GetServiceAsync(typeof(SVsStatusbar)).Result as IVsStatusbar;
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
            string displayName = assemblyDef;
            while (refAsmNames.MoveNext())
            {
                refs.MoveNext();
                if (!string.Equals(refAsmNames.Current.ToString(), assemblyDef, StringComparison.OrdinalIgnoreCase))
                    continue;

                displayName = refs.Current.Display;
                metReference = refs.Current;

                //maybe a reference assembly such as "C:\\Program Files (x86)\\Reference Assemblies\\Microsoft\\Framework\\.NETFramework\\v4.5.1\\mscorlib.dll"
                if (displayName.Contains("Reference Assemblies"))
                {
                    var fileName  = Path.GetFileName(displayName);
                    displayName = Path.GetFileNameWithoutExtension(fileName);
                    //e.g. C:\windows\Microsoft.Net\assembly\GAC_32\mscorlib\v4.0_4.0.0.0__b77a5c561934e089\mscorlib.dll
                    var gacPath = GacHelper.GetAssemblyPath(displayName);
                    if (gacPath.Item2)
                        return gacPath.Item1;
                    else
                    {
                        return displayName;
                    }
                }

                if (displayName.ToLower().Contains("\\packages\\"))
                {
                    return displayName;
                }

                if (metReference.Properties.Kind == MetadataImageKind.Assembly
                    && File.Exists(displayName))
                {
                    return displayName;
                }
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

        protected override void OnExecute(object sender, EventArgs e)
        {
            MenuItemCallback(sender, e);
        }
    }

}