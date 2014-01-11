using EnvDTE;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace vspte
{
    public static class Extensions
    {
        public static Project FindByName(this IEnumerable<Project> projects, string name)
        {
            foreach (var project in projects)
            {
                if (project.Name == name)
                {
                    return project;
                }
                else
                {
                    foreach (var subproject in project.ProjectItems.Cast<ProjectItem>())
                    {
                        if (subproject.Name == name)
                        {
                            return subproject.SubProject;
                        }
                    }
                }
            }
            throw new Exception("Project Not Found");
        }

        public static IEnumerable<Project> AllProjects(this Solution solution)
        {
            foreach (var project in solution.Projects.Cast<Project>())
            {
                if (project != null && !string.IsNullOrEmpty(project.FullName))
                    yield return project;

                foreach (var subproject in project.ProjectItems.Cast<ProjectItem>())
                {
                    if (subproject.SubProject != null
                         && !string.IsNullOrEmpty(subproject.SubProject.FullName))
                        yield return subproject.SubProject;
                }
            }
        }

        public static IEnumerable<ProjectItem> SolutionItems(this Solution solution)
        {
            foreach (var project in solution.Projects.Cast<Project>())
            {
                foreach (var subproject in project.ProjectItems.Cast<ProjectItem>())
                {
                    if (subproject != null 
                        && subproject.SubProject == null
                        && !string.IsNullOrEmpty(subproject.Name))
                    {
                        yield return subproject;
                    }
                }
            }
        }
        
    }
}
