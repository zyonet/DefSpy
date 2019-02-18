using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Process = System.Diagnostics.Process;

namespace DefSpy
{

    /// <summary>
    /// Credit: The following code based on ILSpy's ILSpyAddin code base
    /// </summary>

    public class ILSpyParameters
    {
        public ILSpyParameters(IEnumerable<string> assemblyFileNames, params string[] arguments)
        {
            this.AssemblyFileNames = assemblyFileNames;
            this.Arguments = arguments;
        }

        public IEnumerable<string> AssemblyFileNames { get; private set; }
        public string[] Arguments { get; private set; }
    }

    public class DetectedReference
    {
        public DetectedReference(string name, string assemblyFile, bool isProjectReference)
        {
            this.Name = name;
            this.AssemblyFile = assemblyFile;
            this.IsProjectReference = isProjectReference;
        }

        public string Name { get; private set; }
        public string AssemblyFile { get; private set; }
        public bool IsProjectReference { get; private set; }
    }
    abstract class ILSpyCommand
    {
        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("0B348AFC-8219-41DD-93CE-90AD6396B2EB");

        protected SpyDefinitionPackage owner;

        protected ILSpyCommand(SpyDefinitionPackage owner, uint id)
        {
            this.owner = owner;
            CommandID menuCommandID = new CommandID(CommandSet, (int)id);
            OleMenuCommand menuItem = new OleMenuCommand(OnExecute, menuCommandID);
            menuItem.BeforeQueryStatus += OnBeforeQueryStatus;

            OleMenuCommandService commandService =
                this.owner.GetServiceAsync(typeof(IMenuCommandService)).Result as OleMenuCommandService;
            if (commandService != null)
            {
                commandService.AddCommand(menuItem);
            }
        }

        protected virtual void OnBeforeQueryStatus(object sender, EventArgs e)
        {
        }


        protected abstract void OnExecute(object sender, EventArgs e);

        protected string GetILSpyPath()
        {
            var basePath = Path.GetDirectoryName(typeof(SpyDefinitionPackage).Assembly.Location);
            return Path.Combine(basePath, "ILSpy.exe");
        }

        protected void OpenAssembliesInILSpy(ILSpyParameters parameters)
        {
            if (parameters == null)
                return;

            foreach (string assemblyFileName in parameters.AssemblyFileNames)
            {
                if (!File.Exists(assemblyFileName))
                {
                    owner.ShowMessage("Could not find assembly '{0}', please ensure the project and all references were built correctly!", assemblyFileName);
                    return;
                }
            }

            string commandLineArguments = ICSharpCode.ILSpy.AddIn.Utils.ArgumentArrayToCommandLine(parameters.AssemblyFileNames.ToArray());
            if (parameters.Arguments != null)
            {
                commandLineArguments = string.Concat(commandLineArguments, " ", ICSharpCode.ILSpy.AddIn.Utils.ArgumentArrayToCommandLine(parameters.Arguments));
            }

            Process.Start(GetILSpyPath(), commandLineArguments);
        }


        protected EnvDTE.Project FindProject(IEnumerable<EnvDTE.Project> projects, string projectFile)
        {
            foreach (var project in projects)
            {
                if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                {
                    // This is a solution folder -> search in sub-projects
                    var subProject = FindProject(
                        project.ProjectItems.OfType<ProjectItem>().Select(pi => pi.SubProject).OfType<EnvDTE.Project>(),
                        projectFile);
                    if (subProject != null)
                        return subProject;
                }
                else
                {
                    if (project.FileName == projectFile)
                        return project;
                }
            }

            return null;
        }
    }
}