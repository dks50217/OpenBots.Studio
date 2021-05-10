﻿using OpenBots.Core.Enums;
using OpenBots.Core.Settings;
using System;
using System.IO;
using System.Windows.Forms;

namespace OpenBots.Core.IO
{
    public static class Folders
    {
        public static string GetFolder(FolderType folderType)
        {
            switch (folderType)
            {
                case FolderType.RootFolder:
                    //return root folder from settings
                    var rootSettings = new ApplicationSettings().GetOrCreateApplicationSettings();
                    var rootFolder = rootSettings.ClientSettings.RootFolder;
                    return rootFolder;
                case FolderType.AttendedTasksFolder:
                    //return attended tasks folder from settings
                    var attendedSettings = new ApplicationSettings().GetOrCreateApplicationSettings();
                    var attentedTasksFolder = attendedSettings.ClientSettings.AttendedTasksFolder;
                    return attentedTasksFolder;
                case FolderType.SettingsFolder:
                    //return app data OpenBots folder
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OpenBots Inc");
                case FolderType.StudioFolder:
                    //return app data OpenBots.Studio folder
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OpenBots Inc", "OpenBots.Studio");
                case FolderType.ScriptsFolder:
                    //return scripts folder
                    return Path.Combine(GetFolder(FolderType.RootFolder), "My Scripts");
                case FolderType.PublishedFolder:
                    //return scripts folder
                    return Path.Combine(GetFolder(FolderType.RootFolder), "Published");
                case FolderType.LogFolder:
                    //return logs folder
                    return Path.Combine(GetFolder(FolderType.RootFolder), "Logs");
                case FolderType.TempFolder:
                    //return temp folder
                    return Path.Combine(Path.GetTempPath(), "OpenBotsStudio");
                case FolderType.LocalAppDataPackagesFolder:
                    //return local app data packages folder
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OpenBots Inc", "packages");
                case FolderType.ProgramFilesPackagesFolder:
                    //return program files packages folder
                    return Path.Combine(new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent.FullName, "packages", Application.ProductVersion);
                default:
                    //enum is not implemented
                    throw new NotImplementedException("FolderType " + folderType.ToString() + " Not Supported");
            }
        }
    }
}
