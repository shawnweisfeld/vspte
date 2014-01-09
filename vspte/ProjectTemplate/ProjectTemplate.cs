using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using EnvDTE;

namespace vspte.ProjectTemplate
{
    public class ProjectTemplateBuilder
    {
        public static void Create(Solution sln, string path)
        {
            var projects = new List<object>();

            foreach (var project in sln.Projects.Cast<Project>())
            {
                //just a regular project
                if (project != null && !string.IsNullOrEmpty(project.FullName))
                {
                    projects.Add(new ProjectTemplateLink()
                    {
                        ProjectName = project.FullName,
                        Value = project.FullName + "\\MyTemplate.vstemplate"
                    });
                }

                //a solution folder
                else if (project != null && !string.IsNullOrEmpty(project.Name))
                {
                    var items = new List<object>();

                    foreach (var subproject in project.ProjectItems.Cast<EnvDTE.ProjectItem>())
                    {
                        if (subproject.SubProject != null
                             && !string.IsNullOrEmpty(subproject.SubProject.FullName))
                        {
                            items.Add(new ProjectTemplateLink()
                            {
                                ProjectName = subproject.SubProject.Name,
                                Value = subproject.SubProject.Name + "\\MyTemplate.vstemplate"
                            });
                        }
                    }

                    projects.Add(new SolutionFolder()
                    {
                        Name = project.Name,
                        Items = items.ToArray()
                    });
                }
            }

            var template = new VSTemplate()
            {
                Version = "2.0.0",
                Type = "ProjectGroup",
                TemplateData = new VSTemplateTemplateData()
                {
                    Name = new NameDescriptionIcon() { Value = "HumanTracker" },
                    Description = new NameDescriptionIcon() { Value = "HumanTracker" },
                    Icon = new NameDescriptionIcon() { Value = "Icon.ico" },
                    ProjectType = "CSharp"
                },
                TemplateContent = new VSTemplateTemplateContent()
                {
                    Items = new object[] 
                    {
                       new VSTemplateTemplateContentProjectCollection()
                       {
                           Items = projects.ToArray()
                       }
                    }
                }
            };

            using (TextWriter writer = new StreamWriter(Path.Combine(path, "MyTemplate.vstemplate")))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(VSTemplate));
                serializer.Serialize(writer, template);
            }
        }
    }
}
