﻿using CSScriptLibrary;
using OpenBots.Core.Enums;
using OpenBots.Core.Project;
using OpenBots.Core.Script;
using OpenBots.Core.Utilities.CommonUtilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace OpenBots.Utilities
{
    public class ExecutionManager
    {
        private const int MAX_PATH = 260;

        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, SetLastError = false)]
        public static extern bool PathFindOnPath([In, Out] StringBuilder pszFile, [In] string[] ppszOtherDirs);

        public static void RunTextEditorProject(string configPath, List<ProjectArgument> scriptArgs)
        {
            Project project = Project.OpenProject(configPath);
            string mainPath = Path.Combine(new FileInfo(configPath).DirectoryName, project.Main);
            switch (project.ProjectType)
            {
                case ProjectType.Python:
                    RunPythonAutomation(mainPath, scriptArgs);
                    break;
                case ProjectType.TagUI:
                    RunTagUIAutomation(mainPath, new FileInfo(configPath).DirectoryName, scriptArgs);
                    break;
                case ProjectType.CSScript:
                    RunCSharpAutomation(mainPath, scriptArgs);
                    break;
            }
        }

        public static void RunPythonAutomation(string scriptPath, List<ProjectArgument> scriptArgs)
        {
            string error = "";
            string version = scriptArgs.Where(x => x.ArgumentName == "--PythonVersion").FirstOrDefault().ArgumentValue?.ToString();
            if (version == null)
                throw new ArgumentNullException("--PythonVersion");

            string pythonExecutable = CommonMethods.GetPythonPath(Environment.UserName, version);

            string strScriptArgs = string.Join(" ", scriptArgs.Select(x => $"{x.ArgumentName} {x.ArgumentValue}".Trim())
                                                              .ToList());

            if (!string.IsNullOrEmpty(pythonExecutable))
            {
                Process scriptProc = new Process();

                scriptProc.StartInfo = new ProcessStartInfo()
                {
                    FileName = pythonExecutable,
                    Arguments = $"\"{scriptPath}\" {strScriptArgs}",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                scriptProc.Start();
                scriptProc.WaitForExit();

                scriptProc.StandardOutput.ReadToEnd();
                error = scriptProc.StandardError.ReadToEnd();
                scriptProc.Close();
            }

            if (!string.IsNullOrEmpty(error))
                throw new Exception(error);
        }

        public static void RunCSharpAutomation(string scriptPath, List<ProjectArgument> scriptArgs)
        {
            string code = File.ReadAllText(scriptPath);
            MethodInfo mainMethod = CSScript.LoadCode(code).CreateObject("*").GetType().GetMethod("Main");

            object[] scriptArgsArray;
            if (scriptArgs.Count == 0)
                scriptArgsArray = new object[] { null };
            else
                scriptArgsArray = new object[] { scriptArgs.Select(x => x.ArgumentValue).ToArray() };
            
            if (mainMethod.GetParameters().Length == 0)
                mainMethod.Invoke(null, null);
            else
                mainMethod.Invoke(null, scriptArgsArray);
        }

        public static void RunTagUIAutomation(string scriptPath, string projectPath, List<ProjectArgument> scriptArgs)
        {
            string error = "";
            string tagUIexePath = GetFullPathFromWindows("tagui");
            if (tagUIexePath == null)
                throw new Exception("TagUI installation was not detected on the machine. Please perform the installation as outlined in the official documentation.\n"+
                                    "https://tagui.readthedocs.io/en/latest/setup.html");

            // Copy Script Folder/Files to ".\tagui\flows" Directory
            string destinationDirectory = Path.Combine(new DirectoryInfo(tagUIexePath).Parent.Parent.FullName, "flows", new DirectoryInfo(projectPath).Name);
            if (Directory.Exists(destinationDirectory))
                Directory.Delete(destinationDirectory, true);

            Directory.CreateDirectory(destinationDirectory);
            DirectoryCopy(projectPath, destinationDirectory, true);

            string newScriptPath = destinationDirectory + scriptPath.Replace(projectPath, "");

            string strScriptArgs = string.Join(" ", scriptArgs.Select(x => $"{x.ArgumentName} {x.ArgumentValue}".Trim())
                                                              .ToList());

            if (!string.IsNullOrEmpty(tagUIexePath))
            {
                Process scriptProc = new Process();

                scriptProc.StartInfo = new ProcessStartInfo()
                {
                    FileName = tagUIexePath + ".cmd",
                    Arguments = $"\"{newScriptPath}\" {strScriptArgs}",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                scriptProc.Start();
                scriptProc.WaitForExit();

                scriptProc.StandardOutput.ReadToEnd();
                error = scriptProc.StandardError.ReadToEnd();
                scriptProc.Close();
            }

            if (!string.IsNullOrEmpty(error))
                throw new Exception(error);
        }

        private static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            DirectoryInfo[] dirs = dir.GetDirectories();
            // If the destination directory doesn't exist, create it.
            Directory.GetParent(destDirName);
            if (!Directory.GetParent(destDirName).Exists)
            {
                throw new DirectoryNotFoundException(
                    "Destination directory does not exist or could not be found: "
                    + Directory.GetParent(destDirName));
            }

            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, false);
            }

            // If copying subdirectories, copy them and their contents to new location.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }

        private static string GetFullPathFromWindows(string exeName)
        {
            if (exeName.Length >= MAX_PATH)
                throw new ArgumentException($"The executable name '{exeName}' must have less than {MAX_PATH} characters.",
                    nameof(exeName));

            StringBuilder sb = new StringBuilder(exeName, MAX_PATH);
            var exePath = PathFindOnPath(sb, null) ? sb.ToString() : null;

            if (exePath != null)
                return exePath;

            // Get User Environment Variable "Path"
            var envPathValue = Environment.GetEnvironmentVariable("Path", EnvironmentVariableTarget.User);

            exePath = FindAppPath(envPathValue, exeName);
            if (!string.IsNullOrEmpty(exePath))
                return exePath;

            // Get System Environment Variable "Path"
            envPathValue = Environment.GetEnvironmentVariable("Path", EnvironmentVariableTarget.Machine);
            exePath = FindAppPath(envPathValue, exeName);
            if (!string.IsNullOrEmpty(exePath))
                return exePath;
            else
                return null;

        }

        private static string FindAppPath(string envPathValue, string exeName)
        {
            string appFullPath = string.Empty;
            if (envPathValue != null)
            {
                var pathValues = envPathValue.ToString().Split(Path.PathSeparator);

                foreach (var path in pathValues)
                {
                    if (File.Exists(Path.Combine(path, exeName)))
                    {
                        appFullPath = Path.Combine(path, exeName);
                        break;
                    }
                }
            }

            return appFullPath;
        }
    }
}
