﻿using System.Linq;
using System.Net;
using System.Net.NetworkInformation;

namespace OpenBots.Core.Server_Documents.User
{
    public class SystemInfo
    {
        public static string MachineName = Dns.GetHostName();

        public static string GetMacAddress()
        {
            const int MIN_MAC_ADDR_LENGTH = 12;
            string macAddress = string.Empty;
            long maxSpeed = -1;

            foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                string tempMac = nic.GetPhysicalAddress().ToString();
                if (nic.Speed > maxSpeed &&
                    !string.IsNullOrEmpty(tempMac) &&
                    tempMac.Length >= MIN_MAC_ADDR_LENGTH)
                {
                    maxSpeed = nic.Speed;
                    macAddress = tempMac;
                }
            }
            return string.Join("-", Enumerable.Range(0, 6).Select(i => macAddress.Substring(i * 2, 2)));
        }
    }    
}
