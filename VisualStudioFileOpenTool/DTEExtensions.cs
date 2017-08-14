namespace VisualStudioFileOpenTool
{
    using System.Collections.Generic;
    using System.Linq;
    using EnvDTE;

    /// <summary>
    /// Based on http://www.wwwlicious.com/2011/03/29/envdte-getting-all-projects-html/
    /// </summary>
    public static class DTEExtensions
    {
        // we want to avoid referencing EnvDTE80 so we will include the constant right here
        private const string VsProjectKindSolutionFolder = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";

        public static IEnumerable<Project> GetAllProjects(this Solution sln)
        {
            foreach (Project projectRef in sln.Projects)
            {
                foreach (var project in GetProjects(projectRef))
                {
                    yield return project;
                }
            }
        }

        private static IEnumerable<Project> GetProjects(Project project)
        {
            if (project.Kind == VsProjectKindSolutionFolder)
            {
                return project.ProjectItems
                              .Cast<ProjectItem>()
                              .Select(p => p.SubProject)
                              .Where(p => p != null)
                              .SelectMany(GetProjects);
            }

            return new[] { project };
        }
    }
}