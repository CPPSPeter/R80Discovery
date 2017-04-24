using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace Excel2CP
{
    class funcPolParser
    {
        public static void ParseRule(DataRow drRule, int RuleNumber)
        {
            if (RuleNumber == 46)
            { }
            //rmMain.Instance.RTBWriteState = "\nParsing Rule:";
            string RuleHeader = drRule[1].ToString().Trim();
            string SourceField = drRule[2].ToString().Trim();
            string DestinationField = drRule[3].ToString().Trim();

            string ProtocolField = drRule[4].ToString().Trim();
            string ServiceField = drRule[5].ToString().Trim();

            string ActionField = drRule[6].ToString().Trim();
            string CommentField = drRule[7].ToString().Trim();

            string[] SourceFIleds = SourceField.Split(';');
            string RuleSource ="";
            foreach (string Src in SourceFIleds)
            {
                string SrcFieldData = Src;
                if (Src.Contains("/"))
                {
                    string Subnet = Src.Substring(Src.IndexOf("/") + 1, Src.Length - Src.IndexOf("/")-1);
                    string LongSubnet = IPFunctions.Subnet_Short_To_Long(Subnet);
                    SrcFieldData = Src.Replace("/" + Subnet, " " + LongSubnet);
                }
                RuleSource = RuleSource + ParseObject(SrcFieldData.Trim(), RuleNumber) + ";";
            }

            string[] DestinationFields = DestinationField.Split(';');
            string RuleDestination = "";
            foreach (string Dst in DestinationFields)
            {
                string DstFieldData = Dst;
                if (Dst.Contains("/"))
                {
                    string Subnet = Dst.Substring(Dst.IndexOf("/") + 1, Dst.Length - Dst.IndexOf("/") - 1);
                    string LongSubnet = IPFunctions.Subnet_Short_To_Long(Subnet);
                    DstFieldData = Dst.Replace("/" + Subnet, " " + LongSubnet);
                }
                RuleDestination = RuleDestination + ParseObject(DstFieldData.Trim(), RuleNumber) + ";";
            }

            string[] ProtocolFileds = ProtocolField.Split(';');
            string RuleProto = "";
            string RuleService = "";
            foreach (string proto in ProtocolFileds)
            {
                string ProtoInfo = proto.Trim();
                string[] ServiceFields = ServiceField.Split(';');
                
                foreach (string Srv in ServiceFields)
                {
                    string SrvFieldData = Srv;
                    RuleService = RuleService + ParseService(ProtoInfo.ToLower(), SrvFieldData.Trim(), RuleNumber) + ";";
                }
                RuleProto = RuleProto + ProtoInfo.ToLower() + ";";
            }

            string RuleFlag = "";
            if (RuleSource == "" || RuleDestination == "" || RuleService == "" || RuleService.Contains("Unknown") || RuleSource.Contains("Unknown") || RuleDestination.Contains("Unknown"))
            {
                RuleFlag = "Warning";
            }



            //load to datatable
            DataRow drPolRule = frmMain.dtPolicy.NewRow();
            drPolRule[0] = RuleNumber;
            drPolRule[1] = RuleHeader;
            drPolRule[2] = RuleSource;
            drPolRule[3] = RuleDestination;
            drPolRule[4] = RuleService;
            drPolRule[5] = ActionField;
            drPolRule[6] = CommentField;
            drPolRule[7] = RuleProto;
            drPolRule[8] = RuleFlag;
            frmMain.dtPolicy.Rows.Add(drPolRule);
             

        }

        private static string ParseObject(string ObjectField, int RuleNumber)
        {
            string ReturnedObject = "";
            //remove stupid formatings
            if(ObjectField.Contains("(*)"))
            {
                ObjectField=ObjectField.Replace("(*)","").Trim();
            }

            if (ObjectField.ToLower() == "any")
            {
                ReturnedObject = "Any";
            }
            else if (ObjectField.StartsWith("host ") || IPFunctions.IsValidIP(ObjectField))
            {
                string IPData = ObjectField.Replace("host ", "");
                if (IPFunctions.IsValidIP(IPData))
                {
                    frmMain.AddObjectsToDT("host_" + IPData, "host", IPData, "255.255.255.255", "", "");

                }
                else
                {
                    frmMain.Instance.RTBWriteState = "\nRule " + RuleNumber + " has invalid object defined: " + IPData;
                }
                ReturnedObject = "host_" + IPData;
            }
            else if (IPFunctions.IsNetSubnet(ObjectField))
            { 
                string IP = ObjectField.Substring(0,ObjectField.IndexOf(" "));
                string Subnet = ObjectField.Substring(ObjectField.IndexOf(" "), ObjectField.Length - ObjectField.IndexOf(" "));
                Subnet = Subnet.Trim();
                
                if (IPFunctions.IsValidSubnet(Subnet))
                {
                    int CIDR = IPFunctions.GetSubnetMask(Subnet);
                    if (Subnet == "255.255.255.255")
                    {
                        frmMain.AddObjectsToDT("host_" + IP, "host", IP, "255.255.255.255", "", "");
                        ReturnedObject = "host_" + IP;
                    }
                    else
                    {
                        frmMain.AddObjectsToDT("net_" + IP + "_" + CIDR.ToString(), "network", IP, Subnet, "", "");
                        ReturnedObject = "net_" + IP + "_" + CIDR.ToString();
                    }
                }
                else
                {
                    frmMain.Instance.RTBWriteState = "\nRule " + RuleNumber + " has invalid object defined: " + IP + " " + Subnet;
                }
            }
            else if (ObjectField.StartsWith("object-group "))
            {
                ReturnedObject = ObjectField.Replace("object-group ", "");
                DataRow[] drCheckGroupExist = frmMain.dtObjects.Select("Name_Orig = '" + ReturnedObject + "'");
                if (drCheckGroupExist.Length == 0)
                {
                    frmMain.Instance.RTBWriteError = "\nMissing referenced object: " + ReturnedObject;
                    ReturnedObject = "";
                }
            }
            else
            {
                frmMain.Instance.RTBWriteError = "\nUnrecognized referenced object: " + ObjectField + " Rule:" + RuleNumber;
            }




            if (ReturnedObject == "")
            {
                ReturnedObject = "Unknown";
            }

            return ReturnedObject;
        }

        private static string ParseService(string protocol, string ServiceField, int RuleNumber)
        {
            string ReturnedService = "";
            if (ServiceField == "")
            {
                if (protocol == "ip")
                {
                    ReturnedService = "Any";
                }
                else if (protocol == "icmp")
                {
                    ReturnedService = "icmp";
                }
                else if (protocol == "object-group tcpudp")
                {
                    ReturnedService = "Any";
                }
                else if (protocol == "tcp")
                {
                    frmMain.AddServiceToDT("tcp_all", "", "tcp", "1-65535", "", "", "");
                    ReturnedService=  "tcp_all";
                }
                else if (protocol == "udp")
                {
                    frmMain.AddServiceToDT("udp_all", "", "udp", "1-65535", "", "", "");
                    ReturnedService = "udp_all";
                }
                else
                {
                    frmMain.Instance.RTBWriteError = "\nUnrecognized referenced service(protocol): " + protocol + " Rule:" + RuleNumber;
                }
            }
            else if (protocol.ToLower() == "icmp")
            {
                ReturnedService = "icmp";
            }
            else if (ServiceField.StartsWith("eq "))
            {
                string ServData = ServiceField.Replace("eq ", "");
                if (!funcShared.IsNumeric(ServData))
                {
                    //this is a cisco predefined service
                    ReturnedService = ServData.Trim();
                }
                else
                {
                    if (protocol == "tcp")
                    {
                        frmMain.AddServiceToDT("tcp_" + ServData, "", "tcp", ServData, "", "", "");
                        ReturnedService = ReturnedService + "tcp_" + ServData;
                    }
                    else if (protocol == "udp")
                    {
                        frmMain.AddServiceToDT("udp_" + ServData, "", "udp", ServData, "", "", "");
                        ReturnedService = ReturnedService + "udp_" + ServData;
                    }
                    else if (protocol == "tcp-udp" || protocol == "tcp;udp")
                    {
                        frmMain.AddServiceToDT("tcp_" + ServData, "", "tcp", ServData, "", "", "");
                        ReturnedService = ReturnedService + "tcp_" + ServData + ";";
                        frmMain.AddServiceToDT("udp_" + ServData, "", "udp", ServData, "", "", "");
                        ReturnedService = ReturnedService + "udp_" + ServData;
                    }
                    else
                    {
                        frmMain.Instance.RTBWriteError = "\nService - Invalid Group protocol type : " + protocol + " Rule: " + RuleNumber;
                    }
                }
            }
            else if (ServiceField.StartsWith("object-group "))
            {
                ReturnedService = ServiceField.Replace("object-group ", "");
                DataRow[] drCheckGroupExist = frmMain.dtServices.Select("Name_Orig = '" + ReturnedService + "'");
                if (drCheckGroupExist.Length == 0)
                {
                    frmMain.Instance.RTBWriteError = "\nMissing referenced object: " + ReturnedService;
                    ReturnedService = "";
                }
            }
            else
            {
                //these are the manual defined odd things
                if (ServiceField.Contains("(") && ServiceField.Contains(")"))
                {
                    //we ignore the description and only worry about the port that is defined
                    //get the port number from string
                    string ServPort = ServiceField.Remove(0, ServiceField.IndexOf('(') + 1);
                    ServPort = ServPort.Remove(ServPort.IndexOf(')'));
                    ServPort = ServPort.Trim();
                    if (funcShared.IsNumeric(ServPort))
                    {
                        //lets try to find if there is a predefined value here
                        DataRow[] drPredefServ = frmMain.dtDefaultServices.Select("proto = '" + protocol + "' and port = '" + ServPort + "'");
                        if (drPredefServ.Length == 1)
                        {
                            ReturnedService = drPredefServ[0][0].ToString();
                        }
                        else if (drPredefServ.Length > 1)
                        {
                            frmMain.Instance.RTBWriteError = "\nMultiple predefined services match " + ServiceField + " Rule: " + RuleNumber;
                        }
                        else
                        {
                            //we need to create this port
                            frmMain.AddServiceToDT(protocol.ToLower() + "_" + ServPort, "", protocol.ToLower(), ServPort, "", "", "");
                            ReturnedService = ReturnedService + protocol.ToLower() + "_" + ServPort;
                        }
                    }
                    else
                    {
                        frmMain.Instance.RTBWriteError = "\nCannot fetch service port value: " + ServiceField + " Rule: " + RuleNumber;
                    }
                }
                else
                {
                    if (ServiceField.Contains(' '))
                    {
                        ServiceField = ServiceField.ToLower();
                        string ProtoPart = ServiceField.Substring(0, ServiceField.IndexOf(' '));
                        string PortPart = ServiceField.Replace(ProtoPart, "").Trim();
                        if (funcShared.IsNumeric(PortPart))
                        {
                            //lets try to find if there is a predefined value here
                            DataRow[] drPredefServ = frmMain.dtDefaultServices.Select("proto = '" + ProtoPart + "' and port = '" + PortPart + "'");
                            if (drPredefServ.Length == 1)
                            {
                                ReturnedService = drPredefServ[0][0].ToString();
                            }
                            else if (drPredefServ.Length > 1)
                            {
                                frmMain.Instance.RTBWriteError = "\nMultiple predefined services match " + ServiceField + " Rule: " + RuleNumber;
                            }
                            else
                            {
                                //lets create the port
                                frmMain.AddServiceToDT(ProtoPart.ToLower() + "_" + PortPart, "", ProtoPart.ToLower(), PortPart, "", "", "");
                                ReturnedService = ReturnedService + ProtoPart.ToLower() + "_" + PortPart;
                            }
                        }
                        else if (PortPart.Contains('-'))
                        { 
                            //this is a range
                            //we need to create this port
                            frmMain.AddServiceToDT(protocol.ToLower() + "_" + PortPart, "", protocol.ToLower(), PortPart, "", "", "");
                            ReturnedService = ReturnedService + protocol.ToLower() + "_" + PortPart;
                        }
                        else
                        {
                            frmMain.Instance.RTBWriteError = "\nThe service definition is not numeric: " + ServiceField + " Rule: " + RuleNumber;
                        }
                    }
                    else
                    {                        
                        DataRow[] drPredefServ = frmMain.dtDefaultServices.Select("proto = '" + protocol + "' and Name = '" + ServiceField + "'");
                        if (drPredefServ.Length == 1)
                        {
                            ReturnedService = drPredefServ[0][0].ToString();
                        }
                        else 
                        {
                            frmMain.Instance.RTBWriteError = "\nNo service is matching this definition " + ServiceField + " Rule: " + RuleNumber;
                        }    
                    }
                }
            }

            if (ReturnedService == "")
            {
                ReturnedService = "Unknown";
            }

            return ReturnedService;
        }
    }
}
