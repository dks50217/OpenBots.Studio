﻿using System;
using System.IO;
using System.Reflection;

namespace OpenBots.NativeServer.Library
{
    public static class Utils
    {
        public static string AssemblyLoadDirectory()
        {
            string codeBase = Assembly.GetEntryAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            return Path.GetDirectoryName(path);
        }
    }
}
