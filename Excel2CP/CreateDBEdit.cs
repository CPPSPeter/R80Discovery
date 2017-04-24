using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2CP
{
    class CreateDBEdit
    {
        public static bool CreatePolicyHeader(string WorkPath, string FileName, string PolicyName)
        {

            funcFileIO.Write_to_File(WorkPath, FileName, "create policies_collection " + PolicyName);
            funcFileIO.Write_to_File(WorkPath, FileName, "create firewall_policy ##" + PolicyName);
            funcFileIO.Write_to_File(WorkPath, FileName, "modify fw_policies ##" + PolicyName + " collection policies_collections:" + PolicyName);
            return true;
        }

        public static bool CreatePolicyFooter(string WorkPath, string FileName, string PolicyName)
        {
            funcFileIO.Write_to_File(WorkPath, FileName, "update policies_collections " + PolicyName);
            funcFileIO.Write_to_File(WorkPath, FileName, "update fw_policies ##" + PolicyName);
            return true;
        }

        public static bool CreateRule(string WorkPath, string FileName, string PolicyName, int RuleIndex, string RuleName, string RuleSRC, string RuleDST, string RuleSRV, string RuleLog, string RuleAction,
            string RuleComment, bool RuleDisabled, bool srcNegated, bool dstNegated)
        {
            //lets write a rule
            funcFileIO.Write_to_File(WorkPath, FileName, "addelement fw_policies " + PolicyName + " rule security_rule");

            //disabled
            if (RuleDisabled)
            {
                funcFileIO.Write_to_File(WorkPath, FileName, "modify fw_policies " + PolicyName + " rule:" + RuleIndex + ":disabled true");
            }

            //comment
            if (RuleComment != "")
            {
                funcFileIO.Write_to_File(WorkPath, FileName, "modify fw_policies " + PolicyName + " rule:" + RuleIndex + ":comments '" + RuleComment.Replace("\r\n", ";") + "'");
            }

            //name
            if (RuleName != "")
            {
                funcFileIO.Write_to_File(WorkPath, FileName, "modify fw_policies " + PolicyName + " rule:" + RuleIndex + ":name '" + RuleName.Replace("\r\n", "") + "'");
            }

            //src
            string[] SourceList = RuleSRC.Split(';');
            foreach (string SrcObjects in SourceList)
            {
                if (SrcObjects != "" && SrcObjects != ";")
                {
                    if (SrcObjects == "Any")
                    {
                        funcFileIO.Write_to_File(WorkPath, FileName, "addelement fw_policies " + PolicyName + " rule:" + RuleIndex + ":src:'' globals:Any");
                    }
                    else
                    {
                        funcFileIO.Write_to_File(WorkPath, FileName, "addelement fw_policies " + PolicyName + " rule:" + RuleIndex + ":src:'' network_objects:" + SrcObjects);
                    }

                    if (srcNegated)
                    {
                        funcFileIO.Write_to_File(WorkPath, "LOG-" + FileName, "!!!!!! Rule #" + RuleIndex.ToString() + " source is negated");
                    }
                }//fi valid objec check
            }//loop source objects

            //dst
            string[] DestList = RuleDST.Split(';');
            foreach (string DstObjects in DestList)
            {
                if (DstObjects != "" && DstObjects != ";")
                {
                    if (DstObjects == "Any")
                    {
                        funcFileIO.Write_to_File(WorkPath, FileName, "addelement fw_policies " + PolicyName + " rule:" + RuleIndex + ":dst:'' globals:Any");
                    }
                    else
                    {
                        funcFileIO.Write_to_File(WorkPath, FileName, "addelement fw_policies " + PolicyName + " rule:" + RuleIndex + ":dst:'' network_objects:" + DstObjects);
                    }

                    if (dstNegated)
                    {
                        funcFileIO.Write_to_File(WorkPath, "LOG-" + FileName, "!!!!!! Rule #" + RuleIndex.ToString() + " destination is negated");
                    }
                }//fi valid objec check
            }//loop destination objects

            //srv
            string[] ServList = RuleSRV.Split(';');
            foreach (string SrvObjects in ServList)
            {
                if (SrvObjects != "" && SrvObjects != ";")
                {
                    if (SrvObjects == "Any")
                    {
                        funcFileIO.Write_to_File(WorkPath, FileName, "addelement fw_policies " + PolicyName + " rule:" + RuleIndex + ":services:'' globals:Any");
                    }
                    else
                    {
                        funcFileIO.Write_to_File(WorkPath, FileName, "addelement fw_policies " + PolicyName + " rule:" + RuleIndex + ":services:'' services:" + SrvObjects);
                    }
                }//fi valid objec check
            }//loop destination objects

            //logging
            if (RuleLog == "Log")
            {
                funcFileIO.Write_to_File(WorkPath, FileName, "rmbyindex fw_policies " + PolicyName + " rule:" + RuleIndex + ":track 0");
                funcFileIO.Write_to_File(WorkPath, FileName, "addelement fw_policies " + PolicyName + " rule:" + RuleIndex + ":track tracks:Log");
            }

            //action
            if (RuleAction == "accept")
            {
                funcFileIO.Write_to_File(WorkPath, FileName, "addelement fw_policies " + PolicyName + " rule:" + RuleIndex + ":action accept_action:accept");
            }
            else if (RuleAction == "drop")
            {
                funcFileIO.Write_to_File(WorkPath, FileName, "addelement fw_policies " + PolicyName + " rule:" + RuleIndex + ":action drop_action:drop");
            }
            else if (RuleAction == "reject")
            {
                funcFileIO.Write_to_File(WorkPath, FileName, "addelement fw_policies " + PolicyName + " rule:" + RuleIndex + ":action reject_action:reject");
            }

            return true;
        }

        public static bool CreateHost(string WorkPath, string FileName, string Name, string IP, string Comment, bool GenerateDelete)
        {
            try
            {
                funcFileIO.Write_to_File(WorkPath, FileName, "create host_plain " + Name);
                funcFileIO.Write_to_File(WorkPath, FileName, "modify network_objects " + Name + " ipaddr " + IP);
                //funcFileIO.Write_to_File(WorkPath, FileName, "modify network_objects " + Name + " color 'red'");
                if (Comment != "")
                {
                    funcFileIO.Write_to_File(WorkPath, FileName, "modify network_objects " + Name + " comments '" + Comment.Replace("'", "").Trim() + "'");
                }
                funcFileIO.Write_to_File(WorkPath, FileName, "update network_objects " + Name);
                if (GenerateDelete)
                {
                    funcFileIO.Write_to_File(WorkPath, "Del-" + FileName, "delete network_objects " + Name);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool CreateNetwork(string WorkPath, string FileName, string Name, string IP, string Subnet, string Comment, bool GenerateDelete)
        {
            try
            {
                funcFileIO.Write_to_File(WorkPath, FileName, "create network " + Name);
                funcFileIO.Write_to_File(WorkPath, FileName, "modify network_objects " + Name + " ipaddr " + IP);
                funcFileIO.Write_to_File(WorkPath, FileName, "modify network_objects " + Name + " netmask " + Subnet);
                //funcFileIO.Write_to_File(WorkPath, FileName, "modify network_objects " + Name + " color 'red'");
                if (Comment != "")
                {
                    funcFileIO.Write_to_File(WorkPath, FileName, "modify network_objects " + Name + " comments '" + Comment.Replace("'", "").Trim() + "'");
                }
                funcFileIO.Write_to_File(WorkPath, FileName, "update network_objects " + Name);

                if (GenerateDelete)
                {
                    funcFileIO.Write_to_File(WorkPath, "Del-" + FileName, "delete network_objects " + Name);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool CreateObjRange(string WorkPath, string FileName, string Name, string StartIP, string EndIP, string Comment, bool GenerateDelete)
        {
            try
            {
                funcFileIO.Write_to_File(WorkPath, FileName, "create address_range " + Name);
                funcFileIO.Write_to_File(WorkPath, FileName, "modify network_objects " + Name + " ipaddr_first " + StartIP);
                funcFileIO.Write_to_File(WorkPath, FileName, "modify network_objects " + Name + " ipaddr_last " + EndIP);
                if (Comment != "")
                {
                    funcFileIO.Write_to_File(WorkPath, FileName, "modify network_objects " + Name + " comments '" + Comment.Replace("'", "").Trim() + "'");
                }
                funcFileIO.Write_to_File(WorkPath, FileName, "update network_objects " + Name);

                if (GenerateDelete)
                {
                    funcFileIO.Write_to_File(WorkPath, "Del-" + FileName, "delete network_objects " + Name);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool CreateObjectGroup(string WorkPath, string FileName, string Name, string Members, string Comment, bool GenerateDelete)
        {
            try
            {
                funcFileIO.Write_to_File(WorkPath, FileName, "create network_object_group " + Name);
                //funcFileIO.Write_to_File(WorkPath, FileName, "modify network_objects " + Name + " color '" + drGroupCreate[3].ToString() + "'");
                if (Comment != "")
                {
                    funcFileIO.Write_to_File(WorkPath, FileName, "modify network_objects " + Name + " comments '" + Comment.Replace("'", "").Trim() + "'");
                }
                funcFileIO.Write_to_File(WorkPath, FileName, "update network_objects " + Name);

                if (Members != "")
                {
                    string[] GrpMembers = Members.Split(';');
                    foreach (string GrpMember in GrpMembers)
                    {
                        if (GrpMember != "")
                        {
                            funcFileIO.Write_to_File(WorkPath, FileName, "addelement network_objects " + Name + " '' network_objects:" + GrpMember);
                        }
                    }
                    funcFileIO.Write_to_File(WorkPath, FileName, "update network_objects " + Name);
                }

                if (GenerateDelete)
                {
                    funcFileIO.Write_to_File(WorkPath, "Del-" + FileName, "delete network_objects " + Name);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool CreateSectionTitle(string WorkPath, string FileName, string PolicyName, int RuleIndex, string SectionName)
        {
            funcFileIO.Write_to_File(WorkPath, FileName, "addelement fw_policies " + PolicyName + " rule security_header_rule");
            funcFileIO.Write_to_File(WorkPath, FileName, "addelement fw_policies " + PolicyName + " rule:" + RuleIndex + ":action drop_action:drop");
            funcFileIO.Write_to_File(WorkPath, FileName, "modify fw_policies " + PolicyName + " rule:" + RuleIndex + ":disabled true");
            funcFileIO.Write_to_File(WorkPath, FileName, "modify fw_policies " + PolicyName + " rule:" + RuleIndex + ":header_text \"" + SectionName + "\"");
            funcFileIO.Write_to_File(WorkPath, FileName, "modify fw_policies " + PolicyName + " rule:" + RuleIndex + ":state expanded");

            return true;
        }

        public static bool CreateService_Simple(string WorkPath, string FileName, string Name, string Type, string Port, string Comment, bool GenerateDelete)
        {
            if (Type.ToLower() == "tcp")
            { Type = "tcp_service"; }
            else if (Type.ToLower() == "udp")
            { Type = "udp_service"; }
            else
            {
                //wtf
            }

            if (Port.EndsWith(";"))
            {
                Port = Port.Substring(0, Port.Length - 1);
            }

            //lets write the dbedit file
            if (Name.Length > 0 && (Type == "udp_service" || Type == "tcp_service") & Port.Length > 0)
            {
                funcFileIO.Write_to_File(WorkPath, FileName, "create " + Type + " " + Name);
                funcFileIO.Write_to_File(WorkPath, FileName, "modify services " + Name + " port " + Port);
                if (Port.Contains("-"))
                {
                    funcFileIO.Write_to_File(WorkPath, FileName, "modify services " + Name + " include_in_any false");
                }
                if (Comment != "")
                {
                    funcFileIO.Write_to_File(WorkPath, FileName, "modify services " + Name + " comments '" + Comment.Replace("'", "").Trim() + "'");
                }
                funcFileIO.Write_to_File(WorkPath, FileName, "update services " + Name);
                if (GenerateDelete)
                {
                    funcFileIO.Write_to_File(WorkPath, "Del-" + FileName, "delete services " + Name);
                }
                return true;
            }
            else
            {
                funcFileIO.Write_to_File(WorkPath, "ServicesDbEditLog.txt", Name + " Type: " + Type + " Port: " + Port + "- Error Creating DBEdit invalid Name");
                return false;
            }
        }

        public static bool CreateService_Other(string WorkPath, string FileName, string Name, string Protocol, string Comment, bool GenerateDelete)
        {
            if (Name.Length > 0 && Protocol.Length > 0)
            {
                if (Protocol.EndsWith(";"))
                {
                    Protocol = Protocol.Substring(0, Protocol.Length - 1);
                }
                funcFileIO.Write_to_File(WorkPath, FileName, "create other_service " + Name);
                funcFileIO.Write_to_File(WorkPath, FileName, "modify services " + Name + " protocol " + Protocol);
                if (Comment != "")
                {
                    funcFileIO.Write_to_File(WorkPath, FileName, "modify services " + Name + " comments '" + Comment.Replace("'", "").Trim() + "'");
                }
                //funcFileIO.Write_to_File(frmMain.WorkPath, "services.dbedit", "modify services " + Name + " include_in_any " + drService[10].ToString());
                //funcFileIO.Write_to_File(frmMain.WorkPath, "services.dbedit", "modify services " + Name + " replies " + drService[15].ToString());
                funcFileIO.Write_to_File(WorkPath, FileName, "update services " + Name);
                if (GenerateDelete)
                {
                    funcFileIO.Write_to_File(WorkPath, "Del-" + FileName, "delete services " + Name);
                }
                return true;
            }
            else
            {
                funcFileIO.Write_to_File(WorkPath, "ServicesDbEditLog.txt", Name + " Type: Other Protocol: " + Protocol + " - Error Creating DBEdit invalid Name/Protocol");
                return false;
            }


        }

        public static bool CreateService_Group(string WorkPath, string FileName, string Name, string Members, string Comment, bool GenerateDelete)
        {
            if (Name.Length > 0 && Members.Length > 0)
            {
                funcFileIO.Write_to_File(WorkPath, FileName, "create service_group " + Name);
                if (Comment != "")
                {
                    funcFileIO.Write_to_File(WorkPath, FileName, "modify services " + Name + " comments '" + Comment.Replace("'", "").Trim() + "'");
                }
                funcFileIO.Write_to_File(WorkPath, FileName, "update services " + Name);

                //add members to the service
                string[] GrpMembers = Members.Split(';');
                foreach (string GrpMember in GrpMembers)
                {
                    if (GrpMember != "")
                    {
                        funcFileIO.Write_to_File(WorkPath, FileName, "addelement services " + Name + " '' services:" + GrpMember);
                    }
                }
                funcFileIO.Write_to_File(WorkPath, FileName, "update services " + Name);
                if (GenerateDelete)
                {
                    funcFileIO.Write_to_File(WorkPath, "Del-" + FileName, "delete services " + Name);
                }
                return true;
            }
            else
            {
                funcFileIO.Write_to_File(WorkPath, "ServicesDbEditLog.txt", Name + " Type: Group - Error Creating DBEdit invalid Name/Members");
                return false;
            }
        }
    }
}
