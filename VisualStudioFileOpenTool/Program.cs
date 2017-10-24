namespace VisualStudioFileOpenTool
{
    #region

    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using EnvDTE;
    using Process = System.Diagnostics.Process;

    #endregion

    internal class Program
    {
        private static bool FindInstance(out DTE foundDte, string filePath)
        {
            var allInstances = ComHelper2.GetRunningInstances().Cast<DTE>().Reverse().ToList();
            if (allInstances.Count == 0)
            {
                // no VS instances present
                foundDte = null;
                return false;
            }

            if (allInstances.Count == 1)
            {
                // only one VS instance present - open file in it
                foundDte = allInstances.First();
                return true;
            }

            // More than one VS instance is active.
            // Check each one and find which one contains the file and return it.
            // If multiple solutions has this file, prefer one with the fewer projects count.
            foundDte = null;
            var foundDteProjectsCount = int.MaxValue;

            foreach (var instance in allInstances)
            {
                Debug.WriteLine("Checking solution: " + instance.Solution.FullName);
                try
                {
                    var foundProjectItem = instance.Solution.FindProjectItem(filePath);
                    if (foundProjectItem != null)
                    {
#if DEBUG // log all solution projects
                        foreach (var proj in instance.Solution.GetAllProjects())
                        {
                            Debug.WriteLine("Project: " + proj.Name);
                        }
#endif

                        var instanceProjectsCount = instance.Solution.GetAllProjects().Count();
                        if (instanceProjectsCount < foundDteProjectsCount)
                        {
                            foundDte = instance;
                            foundDteProjectsCount = instanceProjectsCount;
                        }
                    }
                }
                catch
                {
                    // faulty DTE instance - ignore exception
                }
            }

            if (foundDte != null)
            {
                // found an VS instance containing the file
                return true;
            }

            // didn't found an VS instance containing the file - return the first available instance
            foundDte = allInstances[0];
            return true;
        }

        [STAThread]
        private static void Main(string[] args)
        {
            try
            {
                if (args == null
                    || args.Length < 1
                    || args.Length > 2)
                {
                    MessageBox.Show(
                        "Visual Studio 2015/2017 file open tool by AtomicTorch Studio."
                        + Environment.NewLine
                        + "Finds an VS instance containing a solution with the specified file to open it at the specified line."
                        + Environment.NewLine
                        + "If there are multiple VS instances, prefer one with the fewer projects count."
                        + Environment.NewLine
                        + "Otherwise runs a new VS instance with this file as an argument."
                        + Environment.NewLine
                        + Environment.NewLine
                        + "usage: <file path> <line number>"
                        + Environment.NewLine
                        + Environment.NewLine);
                    return;
                }

                var filePath = args[0];
                filePath = filePath.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);

                int.TryParse(args[1], out var fileline);

                if (!FindInstance(out var foundDte, filePath))
                {
                    Process.Start(filePath);
                    return;
                }

                // ensure the main window is focused
                foundDte.MainWindow.Activate();
                // apply workaround in case another application steal focus (in our case - MonoGame application do this)
                SetForegroundWindow(new IntPtr(foundDte.MainWindow.HWnd));

                // open file
                var window = foundDte.ItemOperations.OpenFile(filePath);
                // and move caret to the required file line
                var textSelection = (TextSelection)window.Selection;
                textSelection.MoveToLineAndOffset(fileline, 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Visual Studio 2015/2017 file open tool cannot open the file:"
                    + Environment.NewLine
                    + ex.GetType().FullName
                    + Environment.NewLine
                    + ex.Message);
            }
        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}