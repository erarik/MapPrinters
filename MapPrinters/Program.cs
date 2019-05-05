using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Net.NetworkInformation;
using System.Net;
using System.DirectoryServices;
using System.Management;



namespace MapPrinters
{
    class Program
    {
[DllImport("Netapi32.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "DsAddressToSiteNamesEx", CharSet = CharSet.Unicode)]
        private static extern int DsAddressToSiteNamesEx(
            [In] string computerName,
            [In] Int32 EntryCount,
            [In] SOCKET_ADDRESS[] SocketAddresses,
            [Out] out IntPtr SiteNames,
            [Out] out IntPtr SubnetNames);

[DllImport("Netapi32.dll")]
private static extern int NetApiBufferFree([In] IntPtr buffer);

private const int ERROR_SUCCESS = 0;

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
internal struct SockAddr_In
{
    public Int16 sin_family;
    public Int16 sin_port;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
    public byte[] sin_addr;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
    public byte[] sin_zero;
}

[StructLayout(LayoutKind.Sequential)]
internal struct SOCKET_ADDRESS
{
    public IntPtr lpSockaddr;
    public Int32  iSockaddrLength;
}

        static public string GetUsage()
        {
            var text = @"MapPrinters v1.0.0
Usage: MapPrinters.exe
";
            return text;
        }

        static void Main(string[] args)
        {

            if (args[0]=="-?" || args[0] == "-help")
            {
                Console.WriteLine(GetUsage());
                return;
            }
              

            int minimumSpeed = 10000000;
            foreach (NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
            {
                // discard because of standard reasons
                if ((ni.OperationalStatus != OperationalStatus.Up) ||
                    (ni.NetworkInterfaceType == NetworkInterfaceType.Loopback) ||
                    (ni.NetworkInterfaceType == NetworkInterfaceType.Tunnel))
                    continue;


                // this allow to filter modems, serial, etc.
                // I use 10000000 as a minimum speed for most cases
                if (ni.Speed < minimumSpeed)
                    continue;

                // discard virtual cards (virtual box, virtual pc, etc.)
                if ((ni.Description.IndexOf("virtual", StringComparison.OrdinalIgnoreCase) >= 0) ||
                    (ni.Name.IndexOf("virtual", StringComparison.OrdinalIgnoreCase) >= 0))
                    continue;

                // discard "Microsoft Loopback Adapter", it will not show as NetworkInterfaceType.Loopback but as Ethernet Card.
                if (ni.Description.Equals("Microsoft Loopback Adapter", StringComparison.OrdinalIgnoreCase))
                    continue;

                // Ignore loopback addresses (e.g., 127.0.0.1) 
                /*if (IPAddress.IsLoopback(address.Address)) 
                    continue; */
         
                //Console.WriteLine(ni.Name);
                IPInterfaceProperties ipInterface = ni.GetIPProperties();
                foreach (UnicastIPAddressInformation unicastAddress in ipInterface.UnicastAddresses)
                {
                    //Console.WriteLine(unicastAddress.IPv4Mask == null ? "No subnet defined" : unicastAddress.IPv4Mask.ToString());
                    if (unicastAddress.IPv4Mask != null && unicastAddress.IPv4Mask.ToString() != "0.0.0.0")
                    {
                        FindPrinters(GetLocation(GetSubnet(unicastAddress.Address.ToString(), System.Environment.GetEnvironmentVariable("logonserver"))));
                    }
                }
            }
            Console.ReadLine();
        }


        static void MapPrinter(string sShare)
        {
            using (ManagementClass win32Printer = new ManagementClass("Win32_Printer"))
            {
                using (ManagementBaseObject inputParam =
                   win32Printer.GetMethodParameters("AddPrinterConnection"))
                {
                    // Replace <server_name> and <printer_name> with the actual server and
                    // printer names.
                    inputParam.SetPropertyValue("Name", sShare);//"\\\\<server_name>\\<printer_name>"

                    using (ManagementBaseObject result =
                        (ManagementBaseObject)win32Printer.InvokeMethod("AddPrinterConnection", inputParam, null))
                    {
                        uint errorCode = (uint)result.Properties["returnValue"].Value;

                        switch (errorCode)
                        {
                            case 0:
                                Console.Out.WriteLine("[MapPrinter] Successfully connected printer.");
                                break;
                            case 5:
                                Console.Out.WriteLine("[MapPrinter] Access Denied.");
                                break;
                            case 123:
                                Console.Out.WriteLine("[MapPrinter] The filename, directory name, or volume label syntax is incorrect.");
                                break;
                            case 1801:
                                Console.Out.WriteLine("[MapPrinter] Invalid Printer Name.");
                                break;
                            case 1930:
                                Console.Out.WriteLine("[MapPrinter] Incompatible Printer Driver.");
                                break;
                            case 3019:
                                Console.Out.WriteLine("[MapPrinter] The specified printer driver was not found on the system and needs to be downloaded.");
                                break;
                            default:
                                Console.Out.WriteLine("[MapPrinter] Error mapping printer " + sShare + " errorCode=" + errorCode);
                                break;
                        }
                    }
                }
            }
        }



        static void FindPrinters(string sLocation)
        {
            if (sLocation == null)
                return;

            var ds = new DirectorySearcher { Filter = "(&(objectClass=printqueue)(location=" + sLocation + "))" };

            Console.WriteLine("[FindPrinters] Searching printers on location: " + sLocation);

            foreach (SearchResult sr in ds.FindAll())
            {
                //some properties: "name", "servername", "printername", "drivername", "shortservername", "location" , "printShareName"
                string suNCName = sr.Properties["uNCName"][0].ToString();
                string sPrintername = sr.Properties["printername"][0].ToString();

                Console.WriteLine("[FindPrinters] Printer: " + suNCName);
                MapPrinter(suNCName);
            }   
        }


        static string GetLocation(string sSubnet)
        {
            if (sSubnet == null)
                return null;

            DirectoryEntry rootDSE = new DirectoryEntry("LDAP://RootDSE");
            string sNC = rootDSE.Properties["rootDomainNamingContext"].Value.ToString();
            //CN=10.0.216.0/30,CN=Subnets,CN=Sites,CN=Configuration,.....

            string sPath = @"LDAP://CN=" + sSubnet.Replace("/", "\\/") + ",CN=Subnets,CN=Sites,CN=Configuration," + sNC;

            DirectoryEntry de = new DirectoryEntry(sPath);
            
            Console.WriteLine("[GetLocation] Location found: " + de.Properties["location"].Value.ToString());
            return de.Properties["location"].Value.ToString();

        }

        static string GetSubnet(string sIP, string sDC)
        {
            
            IntPtr pSiteNames;
            IntPtr pSubnetNames;
            SOCKET_ADDRESS[] SocketAddresses = new SOCKET_ADDRESS[1];

            String[] ipArray = sIP.Split('.');

            if (ipArray.Length < 4)
                return null;

            SockAddr_In sockAddr;
            sockAddr.sin_family = 2;
            sockAddr.sin_port = 0;
            sockAddr.sin_addr = new byte[4] { byte.Parse(ipArray[0]), byte.Parse(ipArray[1]), byte.Parse(ipArray[2]), byte.Parse(ipArray[3]) };
            sockAddr.sin_zero = new byte[8] { 0, 0, 0, 0, 0, 0, 0, 0 };

            SocketAddresses[0].iSockaddrLength = Marshal.SizeOf(sockAddr);

            IntPtr pSockAddr = Marshal.AllocHGlobal(16);
            Marshal.StructureToPtr(sockAddr, pSockAddr, true);
            SocketAddresses[0].lpSockaddr = pSockAddr;

            if (DsAddressToSiteNamesEx(
                    sDC,
                    1,//EntryCount
                    SocketAddresses,
                    out pSiteNames,
                    out pSubnetNames) == ERROR_SUCCESS)
            {
                string sSite = Marshal.PtrToStringAuto(Marshal.ReadIntPtr(pSiteNames, 0));
                string sSubnet = Marshal.PtrToStringAuto(Marshal.ReadIntPtr(pSubnetNames, 0));

                NetApiBufferFree(pSubnetNames);
                NetApiBufferFree(pSiteNames);
                Console.WriteLine("[GetSubnet] Subnet found: " + sSubnet);
                Console.WriteLine("[GetSubnet] Site found: " + sSite);
                return sSubnet;
            }
            else
            {
                Console.WriteLine("[GetSubnet] No subnet found!");
                return null;
            }
        }
    }

}

