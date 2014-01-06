using CommandLine;
using CommandLine.Text;
using System;
using System.IO;
using System.Linq;

namespace vspte
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var options = new Options();

            if (!Parser.Default.ParseArguments(args, options))
            {
                Environment.Exit(Parser.DefaultExitCodeFail);
            }

            using (var vs = VisualStudio.Create(logTo: Console.Out))
            {
                vs.OpenSolution(options.SlnPath);

                if (string.IsNullOrEmpty(options.ProjectName))
                {
                    var exportedTemplates = vs.ExportTemplate(options.IncludeNuGetPackages).ToList();

                    if (exportedTemplates.Any())
                    {
                        vs.CombineTemplates(Path.GetFileNameWithoutExtension(options.SlnPath), exportedTemplates);
                    }
                    else
                    {
                        Console.WriteLine("No Projects to export!");
                    }
                }
                else
                {
                    vs.ExportTemplate(options.ProjectName, options.IncludeNuGetPackages);
                }

                if (!string.IsNullOrEmpty(options.VsixProjectName))
                {
                    vs.CreateVsix(options.ProjectName, options.VsixProjectName);
                }
            }

            Console.WriteLine("Everything is OK!");
        }
    }

    class Options
    {
        [Option('s', "sln", Required = true, HelpText = "Path to .sln file containing the project you wish to export a template from")]
        public string SlnPath { get; set; }

        [Option('p', "project", Required = false, HelpText = "The name of a project for template export")]
        public string ProjectName { get; set; }

        [Option("vsix", HelpText = "Create VSIX package with project template")]
        public string VsixProjectName { get; set; }

        [Option("nuget", HelpText = "Include NuGet packages to the project template")]
        public bool IncludeNuGetPackages { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            var help = HelpText.AutoBuild(this, current =>
                {
                    current.MaximumDisplayWidth = int.MaxValue;
                    HelpText.DefaultParsingErrorsHandler(this, current);
                });

            return help;
        }
    }
}
