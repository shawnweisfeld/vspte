using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio.ExportTemplate;
using vspte.Com;
using vspte.Export;
using vspte.Vsix;
using System.IO.Compression;

namespace vspte
{
    public class VisualStudio : IDisposable
    {
        private IMessageFilter _messageFilter;
        private DTE _dte;

        private TextWriter Log { get; set; }

        private VisualStudio(TextWriter log)
        {
            Log = log ?? TextWriter.Null;
        }

        public static VisualStudio Create(TextWriter logTo = null)
        {
            var vs = new VisualStudio(logTo);
            vs.Log.Write("Loading Visual Studio...");

            vs._messageFilter = new MessageFilter();

            var dteComClassName = Type.GetTypeFromProgID("VisualStudio.DTE", true);
            vs._dte = (DTE) Activator.CreateInstance(dteComClassName);

            vs.Log.WriteLine(" OK");
            return vs;
        }

        public virtual void OpenSolution(string slnPath)
        {
            Log.Write("Loading solution...");

            slnPath = Path.GetFullPath(slnPath);
            _dte.Solution.Open(slnPath);

            Log.WriteLine(" OK");
        }


        public virtual IEnumerable<string> ExportTemplate(bool includeNuGetPackages)
        {
            foreach (var project in _dte.Solution.AllProjects())
            {
                yield return ExportTemplate(project.Name, includeNuGetPackages);
            }
        }

        public virtual string ExportTemplate(string projectName, bool includeNuGetPackages)
        {
            Log.Write("Exporting project template - {0}...", projectName);

            var template = new ExportTemplatePackage();
            //var package = ExportTemplatePackage.PackageInstance;
            typeof(ExportTemplatePackage)
                .GetField("staticPackage", BindingFlags.Static | BindingFlags.NonPublic)
                .SetValue(null, template);

            var project = _dte.Solution.Projects.Cast<Project>().FindByName(projectName);
            var wizard = new StandaloneTemplateWizardForm();
            wizard.SetUserData("DTE", _dte);
            wizard.SetUserData("IsProjectExport", true);
            wizard.SetUserData("Project", project);
            wizard.SetUserData("TemplateName", project.Name);
            wizard.SetUserData("AutoImport", false);
            wizard.SetUserData("ExplorerOnZip", false);
            wizard.SetUserData("IncludeNuGetPackages", includeNuGetPackages);

            //typeof(ExportTemplateWizard)
            //    .GetMethod("GenProjectXMLFile", BindingFlags.Instance | BindingFlags.NonPublic)
            //    .Invoke(wizard, null);
            //wizard.OnFinish();

            Log.WriteLine(" OK");
            return wizard.GetProjectXMLFile();
        }

        public virtual void CreateVsix(string templateName, string vsixProjectName)
        {
            Log.Write("Creating VSIX package...");

            var vsixProject = _dte.Solution.Projects.Cast<Project>().First(p => p.Name == vsixProjectName);
            var myDocsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var templateZipPath = Path.Combine(myDocsPath, "My Exported Templates", templateName + ".zip");

            new VsixTemplateBuilder().Build(vsixProject, templateZipPath);
            Log.WriteLine(" OK");
        }

        public void Dispose()
        {
            if (_dte.Solution != null)
            {
                _dte.Solution.Close();
            }
            if (_dte != null)
            {
                _dte.Quit();
            }
            if (_messageFilter is IDisposable)
            {
                ((IDisposable) _messageFilter).Dispose();
            }
        }

        public void CombineTemplates(string solutionName, IEnumerable<string> exportedTemplates)
        {
            string directory = string.Empty;
            foreach (var templatePath in exportedTemplates)
            {
                directory = Path.GetDirectoryName(templatePath);
                ZipFile.ExtractToDirectory(templatePath, Path.Combine(directory, solutionName, Path.GetFileNameWithoutExtension(templatePath)));
            }

            if (!string.IsNullOrEmpty(directory))
            {
                ZipFile.CreateFromDirectory(Path.Combine(directory, solutionName), Path.Combine(directory, solutionName + ".zip"));

                foreach (var templatePath in exportedTemplates)
                {
                    File.Delete(templatePath);
                }
                Directory.Delete(Path.Combine(directory, solutionName), true);
            }


        }
    }
}
