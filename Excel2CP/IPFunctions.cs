using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;

namespace Excel2CP
{
    public static class IPFunctions
    {
        public static bool IsValidIP(string addr)
        {
            if (addr == "0.0.0.0")
            {
                return true;
            }
            else if (addr.Contains(" "))
            {
                return false;
            }
            else if (addr.Replace(".", "").Any(char.IsLetter))
            {
                return false;
            }
            else
            {
                if (regIP.IsMatch(addr))
                {
                    //MessageBox.Show("Valid IP Address entered.", "Valid IP Address", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return true;
                }
                //MessageBox.Show(strIPAddressField);
                return false;
            }
        }

        public static bool IsNetSubnet(string addr)
        {
            if (addr.Contains(" "))
            {
                if (IsValidIP(addr.Substring(0, addr.IndexOf(" "))))
                {
                    if(IsValidSubnet(addr.Substring(addr.IndexOf(' ')+1)))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        //function to sort subnets from the smalles to largest
        public static DataTable Sort_Subnets(DataTable dtSubnetList)
        {
            DataTable dtSorted = new DataTable();
            dtSorted.Columns.Add("IP");
            dtSorted.Columns.Add("Subnet");
            dtSorted.Columns.Add("Color");

            DataRow[] drSubnet32 = dtSubnetList.Select("Subnet = '255.255.255.255'");
            foreach (DataRow dr in drSubnet32)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }

            DataRow[] drSubnet31 = dtSubnetList.Select("Subnet = '255.255.255.254'");
            foreach (DataRow dr in drSubnet31)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet30 = dtSubnetList.Select("Subnet = '255.255.255.252'");
            foreach (DataRow dr in drSubnet30)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet29 = dtSubnetList.Select("Subnet = '255.255.255.248'");
            foreach (DataRow dr in drSubnet29)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet28 = dtSubnetList.Select("Subnet = '255.255.255.240'");
            foreach (DataRow dr in drSubnet28)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet27 = dtSubnetList.Select("Subnet = '255.255.255.224'");
            foreach (DataRow dr in drSubnet27)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet26 = dtSubnetList.Select("Subnet = '255.255.255.192'");
            foreach (DataRow dr in drSubnet26)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet25 = dtSubnetList.Select("Subnet = '255.255.255.128'");
            foreach (DataRow dr in drSubnet25)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet24 = dtSubnetList.Select("Subnet = '255.255.255.0'");
            foreach (DataRow dr in drSubnet24)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet23 = dtSubnetList.Select("Subnet = '255.255.254.0'");
            foreach (DataRow dr in drSubnet23)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet22 = dtSubnetList.Select("Subnet = '255.255.252.0'");
            foreach (DataRow dr in drSubnet22)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet21 = dtSubnetList.Select("Subnet = '255.255.248.0'");
            foreach (DataRow dr in drSubnet21)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet20 = dtSubnetList.Select("Subnet = '255.255.240.0'");
            foreach (DataRow dr in drSubnet20)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet19 = dtSubnetList.Select("Subnet = '255.255.224.0'");
            foreach (DataRow dr in drSubnet19)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet18 = dtSubnetList.Select("Subnet = '255.255.192.0'");
            foreach (DataRow dr in drSubnet18)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet17 = dtSubnetList.Select("Subnet = '255.255.128.0'");
            foreach (DataRow dr in drSubnet17)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet16 = dtSubnetList.Select("Subnet = '255.255.0.0'");
            foreach (DataRow dr in drSubnet16)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet15 = dtSubnetList.Select("Subnet = '255.254.0.0'");
            foreach (DataRow dr in drSubnet15)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet14 = dtSubnetList.Select("Subnet = '255.252.0.0'");
            foreach (DataRow dr in drSubnet14)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet13 = dtSubnetList.Select("Subnet = '255.248.0.0'");
            foreach (DataRow dr in drSubnet13)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet12 = dtSubnetList.Select("Subnet = '255.240.0.0'");
            foreach (DataRow dr in drSubnet12)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet11 = dtSubnetList.Select("Subnet = '255.224.0.0'");
            foreach (DataRow dr in drSubnet11)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet10 = dtSubnetList.Select("Subnet = '255.192.0.0'");
            foreach (DataRow dr in drSubnet10)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet9 = dtSubnetList.Select("Subnet = '255.128.0.0'");
            foreach (DataRow dr in drSubnet9)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet8 = dtSubnetList.Select("Subnet = '255.0.0.0'");
            foreach (DataRow dr in drSubnet8)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet7 = dtSubnetList.Select("Subnet = '254.0.0.0'");
            foreach (DataRow dr in drSubnet7)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet6 = dtSubnetList.Select("Subnet = '252.0.0.0'");
            foreach (DataRow dr in drSubnet6)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet5 = dtSubnetList.Select("Subnet = '248.0.0.0'");
            foreach (DataRow dr in drSubnet5)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet4 = dtSubnetList.Select("Subnet = '240.0.0.0'");
            foreach (DataRow dr in drSubnet4)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet3 = dtSubnetList.Select("Subnet = '224.0.0.0'");
            foreach (DataRow dr in drSubnet3)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet2 = dtSubnetList.Select("Subnet = '192.0.0.0'");
            foreach (DataRow dr in drSubnet2)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }
            DataRow[] drSubnet1 = dtSubnetList.Select("Subnet = '128.0.0.0'");
            foreach (DataRow dr in drSubnet1)
            {
                DataRow dtRow = dtSorted.NewRow();
                dtRow[0] = dr[0];
                dtRow[1] = dr[1];
                dtRow[2] = dr[2];
                dtSorted.Rows.Add(dtRow);
            }

            return dtSorted;
        }

        //converts subnet short notation to long
        public static string Subnet_Short_To_Long(string subnet)
        {

            if (subnet == "32")
                return "255.255.255.255";
            else if (subnet == "31")
                return "255.255.255.254";
            else if (subnet == "30")
                return "255.255.255.252";
            else if (subnet == "29")
                return "255.255.255.248";
            else if (subnet == "28")
                return "255.255.255.240";
            else if (subnet == "27")
                return "255.255.255.224";
            else if (subnet == "26")
                return "255.255.255.192";
            else if (subnet == "25")
                return "255.255.255.128";
            else if (subnet == "24")
                return "255.255.255.0";
            else if (subnet == "23")
                return "255.255.254.0";
            else if (subnet == "22")
                return "255.255.252.0";
            else if (subnet == "21")
                return "255.255.248.0";
            else if (subnet == "20")
                return "255.255.240.0";
            else if (subnet == "19")
                return "255.255.224.0";
            else if (subnet == "18")
                return "255.255.192.0";
            else if (subnet == "17")
                return "255.255.128.0";
            else if (subnet == "16")
                return "255.255.0.0";
            else if (subnet == "15")
                return "255.254.0.0";
            else if (subnet == "14")
                return "255.252.0.0";
            else if (subnet == "13")
                return "255.248.0.0";
            else if (subnet == "12")
                return "255.240.0.0";
            else if (subnet == "11")
                return "255.224.0.0";
            else if (subnet == "10")
                return "255.192.0.0";
            else if (subnet == "9")
                return "255.128.0.0";
            else if (subnet == "8")
                return "255.0.0.0";
            else if (subnet == "7")
                return "254.0.0.0";
            else if (subnet == "6")
                return "252.0.0.0";
            else if (subnet == "5")
                return "248.0.0.0";
            else if (subnet == "4")
                return "240.0.0.0";
            else if (subnet == "3")
                return "224.0.0.0";
            else if (subnet == "2")
                return "192.0.0.0";
            else if (subnet == "1")
                return "128.0.0.0";
            else
                return "";
        }

        //check if subnet mask is valid
        public static bool IsValidSubnet(string addr)
        {
            if (addr == "255.255.255.255")
            { return true; }
            if (addr == "255.255.255.254")
            { return true; }
            if (addr == "255.255.255.252")
            { return true; }
            if (addr == "255.255.255.248")
            { return true; }
            if (addr == "255.255.255.240")
            { return true; }
            if (addr == "255.255.255.224")
            { return true; }
            if (addr == "255.255.255.192")
            { return true; }
            if (addr == "255.255.255.128")
            { return true; }
            if (addr == "255.255.255.0")
            { return true; }
            if (addr == "255.255.254.0")
            { return true; }
            if (addr == "255.255.252.0")
            { return true; }
            if (addr == "255.255.248.0")
            { return true; }
            if (addr == "255.255.240.0")
            { return true; }
            if (addr == "255.255.224.0")
            { return true; }
            if (addr == "255.255.192.0")
            { return true; }
            if (addr == "255.255.128.0")
            { return true; }
            if (addr == "255.255.0.0")
            { return true; }
            if (addr == "255.254.0.0")
            { return true; }
            if (addr == "255.252.0.0")
            { return true; }
            if (addr == "255.248.0.0")
            { return true; }
            if (addr == "255.240.0.0")
            { return true; }
            if (addr == "255.224.0.0")
            { return true; }
            if (addr == "255.192.0.0")
            { return true; }
            if (addr == "255.128.0.0")
            { return true; }
            if (addr == "255.0.0.0")
            { return true; }
            if (addr == "254.0.0.0")
            { return true; }
            if (addr == "252.0.0.0")
            { return true; }
            if (addr == "248.0.0.0")
            { return true; }
            if (addr == "240.0.0.0")
            { return true; }
            if (addr == "224.0.0.0")
            { return true; }
            if (addr == "192.0.0.0")
            { return true; }
            if (addr == "128.0.0.0")
            { return true; }
            if (addr == "0.0.0.0")
            { return true; }
            return false;
        }

        //check if subnet mask is valid
        public static int GetSubnetMask(string addr)
        {
            if (addr == "255.255.255.255")
            { return 32; }
            if (addr == "255.255.255.254")
            { return 31; }
            if (addr == "255.255.255.252")
            { return 30; }
            if (addr == "255.255.255.248")
            { return 29; }
            if (addr == "255.255.255.240")
            { return 28; }
            if (addr == "255.255.255.224")
            { return 27; }
            if (addr == "255.255.255.192")
            { return 26; }
            if (addr == "255.255.255.128")
            { return 25; }
            if (addr == "255.255.255.0")
            { return 24; }
            if (addr == "255.255.254.0")
            { return 23; }
            if (addr == "255.255.252.0")
            { return 22; }
            if (addr == "255.255.248.0")
            { return 21; }
            if (addr == "255.255.240.0")
            { return 20; }
            if (addr == "255.255.224.0")
            { return 19; }
            if (addr == "255.255.192.0")
            { return 18; }
            if (addr == "255.255.128.0")
            { return 17; }
            if (addr == "255.255.0.0")
            { return 16; }
            if (addr == "255.254.0.0")
            { return 15; }
            if (addr == "255.252.0.0")
            { return 14; }
            if (addr == "255.248.0.0")
            { return 13; }
            if (addr == "255.240.0.0")
            { return 12; }
            if (addr == "255.224.0.0")
            { return 11; }
            if (addr == "255.192.0.0")
            { return 10; }
            if (addr == "255.128.0.0")
            { return 9; }
            if (addr == "255.0.0.0")
            { return 8; }
            if (addr == "254.0.0.0")
            { return 7; }
            if (addr == "252.0.0.0")
            { return 6; }
            if (addr == "248.0.0.0")
            { return 5; }
            if (addr == "240.0.0.0")
            { return 4; }
            if (addr == "224.0.0.0")
            { return 3; }
            if (addr == "192.0.0.0")
            { return 2; }
            if (addr == "128.0.0.0")
            { return 1; }
            return 0;
        }
        private static Regex regIPold = new Regex(
            @"(?<First>2[0-4]\d|25[0-5]|[01]?\d\d?)\.(?<Second>2[0-4]\d|25"
            + @"[0-5]|[01]?\d\d?)\.(?<Third>2[0-4]\d|25[0-5]|[01]?\d\d?)\.(?"
            + @"<Fourth>2[0-4]\d|25[0-5]|[01]?\d\d?)+$",
            RegexOptions.IgnoreCase
            | RegexOptions.CultureInvariant
            | RegexOptions.IgnorePatternWhitespace
            | RegexOptions.Compiled
            );

        private static Regex regIP = new Regex(
            @"([1-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])(\.([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])){3}$",
            RegexOptions.IgnoreCase
            | RegexOptions.CultureInvariant
            | RegexOptions.IgnorePatternWhitespace
            | RegexOptions.Compiled
            );


        public static DataTable Subnets_To_Datatables(string SourceType, string Filename)
        {
            DataTable dtRoutes = new DataTable();
            dtRoutes.Columns.Add("Route");
            dtRoutes.Columns.Add("Subnet");

            if (SourceType == "File")
            {
                FileStream fileStream = new FileStream(Filename, FileMode.Open, FileAccess.Read);
                string StreamLine = "";

                try
                {
                    StreamReader SR = new StreamReader(fileStream);
                    while ((StreamLine = SR.ReadLine()) != null)
                    {
                        string[] lines = Regex.Split(StreamLine, " ");
                        if (lines[0] != "0.0.0.0" && lines[0] != "127.0.0.0")
                        {
                            DataRow dtRow = dtRoutes.NewRow();
                            dtRow[0] = lines[0];
                            dtRow[1] = lines[1];
                            dtRoutes.Rows.Add(dtRow);
                        }
                    }
                }
                finally
                {
                    fileStream.Close();
                }

            }
            return dtRoutes;
        }

    }
    public static class SubnetMask
    {
        public static readonly IPAddress ClassA = IPAddress.Parse("255.0.0.0");
        public static readonly IPAddress ClassB = IPAddress.Parse("255.255.0.0");
        public static readonly IPAddress ClassC = IPAddress.Parse("255.255.255.0");

        public static IPAddress CreateByHostBitLength(int hostpartLength)
        {
            int hostPartLength = hostpartLength;
            int netPartLength = 32 - hostPartLength;

            if (netPartLength < 2)
                throw new ArgumentException("Number of hosts is to large for IPv4");

            Byte[] binaryMask = new byte[4];

            for (int i = 0; i < 4; i++)
            {
                if (i * 8 + 8 <= netPartLength)
                    binaryMask[i] = (byte)255;
                else if (i * 8 > netPartLength)
                    binaryMask[i] = (byte)0;
                else
                {
                    int oneLength = netPartLength - i * 8;
                    string binaryDigit =
                        String.Empty.PadLeft(oneLength, '1').PadRight(8, '0');
                    binaryMask[i] = Convert.ToByte(binaryDigit, 2);
                }
            }
            return new IPAddress(binaryMask);
        }

        public static IPAddress CreateByNetBitLength(int netpartLength)
        {
            int hostPartLength = 32 - netpartLength;
            return CreateByHostBitLength(hostPartLength);
        }

        public static IPAddress CreateByHostNumber(int numberOfHosts)
        {
            int maxNumber = numberOfHosts + 1;

            string b = Convert.ToString(maxNumber, 2);

            return CreateByHostBitLength(b.Length);
        }
    }

    public static class IPAddressExtensions
    {
        public static IPAddress GetBroadcastAddress(this IPAddress address, IPAddress subnetMask)
        {
            byte[] ipAdressBytes = address.GetAddressBytes();
            byte[] subnetMaskBytes = subnetMask.GetAddressBytes();

            if (ipAdressBytes.Length != subnetMaskBytes.Length)
                throw new ArgumentException("Lengths of IP address and subnet mask do not match.");

            byte[] broadcastAddress = new byte[ipAdressBytes.Length];
            for (int i = 0; i < broadcastAddress.Length; i++)
            {
                broadcastAddress[i] = (byte)(ipAdressBytes[i] | (subnetMaskBytes[i] ^ 255));
            }
            return new IPAddress(broadcastAddress);
        }

        public static IPAddress GetNetworkAddress(this IPAddress address, IPAddress subnetMask)
        {
            byte[] ipAdressBytes = address.GetAddressBytes();
            byte[] subnetMaskBytes = subnetMask.GetAddressBytes();

            if (ipAdressBytes.Length != subnetMaskBytes.Length)
                throw new ArgumentException("Lengths of IP address and subnet mask do not match.");

            byte[] broadcastAddress = new byte[ipAdressBytes.Length];
            for (int i = 0; i < broadcastAddress.Length; i++)
            {
                broadcastAddress[i] = (byte)(ipAdressBytes[i] & (subnetMaskBytes[i]));
            }
            return new IPAddress(broadcastAddress);
        }

        public static bool IsInSameSubnet(this IPAddress address2, IPAddress address, IPAddress subnetMask)
        {
            IPAddress network1 = address.GetNetworkAddress(subnetMask);
            IPAddress network2 = address2.GetNetworkAddress(subnetMask);

            return network1.Equals(network2);
        }
    }
}
