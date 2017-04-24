using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Windows.Forms;

namespace Excel2CP
{
    class funcDBEDITGeneration
    {

        public static bool Generate_ACL_Rule(string PolicyName, bool SimpleCommentsOnly)
        {
            bool Success = true;
            int RuleIndex = 0;
            Application.DoEvents();

            CreateDBEdit.CreatePolicyHeader(frmMain.WorkDir, "6-Policy.dbedit", PolicyName);

            string SectionTitle = "";

            frmMain.Instance.RTBWriteStateDBEGen = "\nGenerating rules";
            frmMain.Instance.RTBWriteStateDBEGen = "\nNumber of rules to generate: " + frmMain.dtPolicy.Rows.Count.ToString();
            
            DataRow[] drRules = frmMain.dtPolicy.Select("1=1");
            foreach (DataRow drRule in drRules)
            {
                Application.DoEvents();
                //check if thee is a new section title needed
                if (drRule[1].ToString() != SectionTitle)
                {
                    frmMain.Instance.RTBWriteStateDBEGen = "\nGenerating rules for header : " + drRule[1].ToString();
                    CreateDBEdit.CreateSectionTitle(frmMain.WorkDir, "6-Policy.dbedit", "##" + PolicyName, RuleIndex, drRule[1].ToString());
                    RuleIndex++;
                    SectionTitle=drRule[1].ToString();
                }
                //lets start the actual rule generation
                string RuleName = drRule[8].ToString();
                string RuleSRC = drRule[2].ToString();
                string RuleDST = drRule[3].ToString();
                string RuleSRV = drRule[4].ToString();
                string RuleAction = drRule[5].ToString();
                string RuleLog = "Log";
                string RuleComment = drRule[6].ToString();                
                string RuleProtocol = drRule[7].ToString();
                bool RuleDisabl = false;

                //source
                string CPSRC ="";
                if (RuleSRC.Contains(';'))
                {
                    string[] RuleSources = RuleSRC.Split(';');
                    foreach (string Source in RuleSources)
                    {
                        if (Source != "")
                        {
                            CPSRC = CPSRC + GetCPObject(Source);
                        }
                    }
                }
                else
                {
                    CPSRC = GetCPObject(RuleSRC);
                }

                //destination
                string CPDST = "";
                if (RuleDST.Contains(';'))
                {
                    string[] RuleDestinationss = RuleDST.Split(';');
                    foreach (string Destination in RuleDestinationss)
                    {
                        if (Destination != "")
                        {
                            CPDST = CPDST + GetCPObject(Destination);
                        }
                    }
                }
                else
                {
                    CPDST = GetCPObject(RuleDST);
                }

                //service
                string CPSRV = "";
                if (RuleSRV.Contains(";"))
                {
                    string[] RuleServices = RuleSRV.Split(';');
                    foreach (string Service in RuleServices)
                    {
                        if (Service != "")
                        {
                            CPSRV = CPSRV + GetCPService(Service, RuleProtocol, false) + ";";
                        }
                    }
                }
                else
                {
                    CPSRV = GetCPService(RuleSRV, RuleProtocol, false);
                }
                string CPAction = "";
                string CPLOG = RuleLog;

                DataRow drCPRule = frmMain.dtCPPolicy.NewRow();
                drCPRule[0] = RuleIndex;
                drCPRule[1] = CPSRC;
                drCPRule[2] = CPDST;
                drCPRule[3] = CPSRV;
                if (RuleAction == "permit")
                {
                    drCPRule[4] = "accept";
                    CPAction = "accept";
                }
                else if (RuleAction == "deny")
                {
                    drCPRule[4] = "drop";
                    CPAction = "drop";
                }
                else
                {
                    frmMain.Instance.RTBWriteErrorDBEGen = "\nUnknown Actions:" + RuleAction;
                }

               
                drCPRule[6] = RuleComment;
                drCPRule[7] = "";
                drCPRule[8] = RuleName;
                frmMain.dtCPPolicy.Rows.Add(drCPRule);

                //do the dbedit for the rule
                CreateDBEdit.CreateRule(frmMain.WorkDir, "6-Policy.dbedit", "##" + PolicyName, RuleIndex, RuleName.Trim(), CPSRC, CPDST, CPSRV, CPLOG, CPAction, RuleComment.Trim(), RuleDisabl, false, false);
                RuleIndex++;
            }


            //cleanup rule
            CreateDBEdit.CreateSectionTitle(frmMain.WorkDir, "6-Policy.dbedit", "##" + PolicyName, RuleIndex, "Cleanup Rule");
            RuleIndex++;
            CreateDBEdit.CreateRule(frmMain.WorkDir, "6-Policy.dbedit", "##" + PolicyName, RuleIndex, "Cleanup Rule", "Any", "Any", "Any", "Log", "drop", "", false, false, false);
            //commit policy
            CreateDBEdit.CreatePolicyFooter(frmMain.WorkDir, "6-Policy.dbedit", PolicyName);


            return Success;
        }


        public static string GetCPService(string ACLService, string Protocol, bool isMember)
        {
            string CPService = "";
            if (ACLService == "Unknown")
            {
                if (!frmMain.DBEditServceCreated.Contains(";Dummy_Placeholder_Serv;"))
                {
                    CPService = "Dummy_Placeholder_Serv";
                    CreateDBEdit.CreateService_Simple(frmMain.WorkDir, "3-services.dbedit", "Dummy_Placeholder_Serv", "tcp", "11111", "Dummy placeholder service for manual fixup", frmMain.GenerateDeleteFiles);
                    frmMain.DBEditServceCreated = frmMain.DBEditServceCreated + "Dummy_Placeholder_Serv;";
                }
                else
                {
                    CPService = "Dummy_Placeholder_Serv";
                }
            }
            else if (ACLService == "Any")
            {
                CPService = "Any";
            }
            else
            {   
                DataRow[] drServices = frmMain.dtServices.Select("Name_Orig='" + ACLService + "'");                
                if (drServices.Length > 0)
                {
                    bool OKToProceed = false;
                    if (drServices.Length == 1)
                    {
                        OKToProceed = true;
                    }
                    else
                    {
                        OKToProceed = funcShared.Compare_Services(drServices);
                    }

                    if (OKToProceed)
                    {
                        if (frmMain.DBEditServceCreated.Contains(";" + drServices[0][1].ToString() + ";"))
                        {
                            //object already exists
                            CPService = drServices[0][1].ToString();
                        }
                        else
                        {
                            //lets see if the object is in cpinfo
                            //CPService = CheckIfCPServiceAlreadyExistsinCpinfo(drServices[0][4].ToString(), drServices[0][5].ToString(), drServices[0][3].ToString(), CurrentForm, drServices[0][2].ToString());
                            if (CPService == "")
                            {
                                if (drServices[0][3].ToString() == "group")
                                {
                                    string[] Members = drServices[0][6].ToString().Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                                    string GroupMembers = ";";
                                    foreach (string Member in Members)
                                    {
                                        if (Member == ACLService)
                                        {
                                            //this is a nasty exception where a member is the same object as the parent                                            
                                            string CPPredefined = funcShared.Get_CP_Service_Predefined(Member);
                                            if (CPPredefined != "")
                                            {
                                                GroupMembers = GroupMembers + Member + ";";
                                            }
                                            else
                                            {
                                                //lets notify the user about this
                                                frmMain.Instance.RTBWriteErrorDBEGen = "\nThe service parent has the same name as the member - circular reference. Service: " + ACLService;
                                            }

                                        }
                                        else
                                        {
                                            string CPmemberServiceName = GetCPService(Member, "*", true);
                                            GroupMembers = GroupMembers + CPmemberServiceName + ";";
                                        }
                                    }
                                    //create the group object
                                    //CreateDBEdit.CreateObjectGroup(frmMain.WorkDir, "2-groups.dbedit", drInterfaceGrp[0][2].ToString(), GroupMembers, drInterfaceGrp[0][7].ToString(), frmMain.GenerateDeleteFiles);
                                    if (frmMain.ReplaceASDMName)
                                    {
                                        if (drServices[0][1].ToString().StartsWith("DM_INLINE_"))
                                        {
                                            CPService = GroupMembers;
                                        }
                                        else
                                        {
                                            CPService = drServices[0][1].ToString();
                                            CreateDBEdit.CreateService_Group(frmMain.WorkDir, "4-service-grp.dbedit", CPService, GroupMembers, drServices[0][7].ToString(), frmMain.GenerateDeleteFiles);
                                        }
                                    }
                                    else
                                    {
                                        CPService = drServices[0][1].ToString();
                                        CreateDBEdit.CreateService_Group(frmMain.WorkDir, "4-service-grp.dbedit", CPService, GroupMembers, drServices[0][7].ToString(), frmMain.GenerateDeleteFiles);
                                    }
                                }
                                else if (!funcShared.IsNumeric(drServices[0][5].ToString()) && !drServices[0][5].ToString().Contains('-'))
                                {
                                    DataRow[] PredefServices = frmMain.dtDefaultServices.Select("Name= '" + ACLService + "'");
                                    if (PredefServices.Length == 1)
                                    {
                                        CPService = PredefServices[0][3].ToString().Trim();
                                    }
                                    else
                                    {
                                        frmMain.Instance.RTBWriteErrorDBEGen = "\nCannot find definition for service " + ACLService;
                                        if (!frmMain.DBEditServceCreated.Contains(";Dummy_Placeholder_Serv;"))
                                        {
                                            CPService = "Dummy_Placeholder_Serv";
                                            CreateDBEdit.CreateService_Simple(frmMain.WorkDir, "3-services.dbedit", "Dummy_Placeholder_Serv", "tcp", "11111", "Dummy placeholder service for manual fixup", frmMain.GenerateDeleteFiles);
                                            frmMain.DBEditServceCreated = frmMain.DBEditServceCreated + "Dummy_Placeholder_Serv;";
                                        }
                                        else
                                        {
                                            CPService = "Dummy_Placeholder_Serv";
                                        }
                                    } 
                                }
                                else if (drServices[0][4].ToString() == "tcp" || drServices[0][4].ToString() == "udp")
                                {
                                    if (drServices[0][5].ToString() == "")
                                    {
                                        frmMain.Instance.RTBWriteErrorDBEGen = "\nProtocol defined but no port found " + ACLService;
                                    }
                                    else
                                    {
                                        CPService = drServices[0][1].ToString();
                                        CreateDBEdit.CreateService_Simple(frmMain.WorkDir, "3-services.dbedit", CPService, drServices[0][4].ToString(), drServices[0][5].ToString(), drServices[0][7].ToString(), frmMain.GenerateDeleteFiles);
                                    }
                                }
                                else if (drServices[0][3].ToString() == "proto")
                                {
                                    if (drServices[0][4].ToString() == "")
                                    {
                                        frmMain.Instance.RTBWriteErrorDBEGen = "\nProtocol defined but no protocol ID found " + ACLService;
                                    }
                                    else
                                    {
                                        CPService = drServices[0][1].ToString();
                                        CreateDBEdit.CreateService_Other(frmMain.WorkDir, "3-services.dbedit", CPService, drServices[0][4].ToString(), drServices[0][7].ToString(), frmMain.GenerateDeleteFiles);
                                    }
                                }
                                else if (drServices[0][1].ToString() == "TCPUDP" && drServices[0][5].ToString() == "Any")
                                {
                                    CPService = drServices[0][5].ToString();
                                }
                                else
                                {
                                    frmMain.Instance.RTBWriteErrorDBEGen = "\nUnhandled service " + ACLService;
                                }

                            }
                        }
                    }
                    else
                    {
                        frmMain.Instance.RTBWriteErrorDBEGen = "\nMultiple matches for service . this will have to be fixed up manually " + ACLService;
                        if (!frmMain.DBEditServceCreated.Contains(";Dummy_Placeholder_Serv;"))
                        {
                            CPService = "Dummy_Placeholder_Serv";
                            CreateDBEdit.CreateService_Simple(frmMain.WorkDir, "3-services.dbedit", "Dummy_Placeholder_Serv", "tcp", "11111", "Dummy placeholder service for manual fixup", frmMain.GenerateDeleteFiles);
                            frmMain.DBEditServceCreated = frmMain.DBEditServceCreated + "Dummy_Placeholder_Serv;";
                        }
                        else
                        {
                            CPService = "Dummy_Placeholder_Serv";
                        }
                    }
                }
                else
                {
                    //we dont have this object parsed!!!!!!!!!!!

                    //lets look at predefined services                    
                    DataRow[] PredefServices = frmMain.dtDefaultServices.Select("Name= '" + ACLService + "'");
                    if (PredefServices.Length == 1)
                    {
                        CPService = PredefServices[0][3].ToString().Trim();
                    }
                    else
                    {
                        frmMain.Instance.RTBWriteErrorDBEGen = "\nCannot find definition for service " + ACLService;
                        if (!frmMain.DBEditServceCreated.Contains(";Dummy_Placeholder_Serv;"))
                        {
                            CPService = "Dummy_Placeholder_Serv";
                            CreateDBEdit.CreateService_Simple(frmMain.WorkDir, "3-services.dbedit", "Dummy_Placeholder_Serv", "tcp", "11111", "Dummy placeholder service for manual fixup", frmMain.GenerateDeleteFiles);
                            frmMain.DBEditServceCreated = frmMain.DBEditServceCreated + "Dummy_Placeholder_Serv;";
                        }
                        else
                        {
                            CPService = "Dummy_Placeholder_Serv";
                        }
                    }                    
                }//fi object not defined and not predefined

            }
            if (CPService != "")
            {
                frmMain.DBEditServceCreated = frmMain.DBEditServceCreated + CPService + ";";
            }
            return CPService;
        }


        //the ismember flag serves to differentiate between looking at the object based on cpname or based on original name. If we talking members then this flag is set to true
        public static string GetCPObject(string ACLObject)
        {
            string CPObject = "";
            if (ACLObject == "Unknown")
            {
                if (!frmMain.DBEditObjectsCreated.Contains(";Dummy_Placeholder;"))
                {
                    CPObject = "Dummy_Placeholder";
                    CreateDBEdit.CreateHost(frmMain.WorkDir, "1-objects.dbedit", "Dummy_Placeholder", "1.1.1.1", "Placeholder to fix up manually", frmMain.GenerateDeleteFiles);
                    frmMain.DBEditObjectsCreated = frmMain.DBEditObjectsCreated + "Dummy_Placeholder;";
                }
                else
                {
                    CPObject = "Dummy_Placeholder";
                }
            }
            else if (ACLObject == "Any")
            {
                CPObject = "Any";      
            }
            else
            {
                //we have a defined object
                string Query = "";
                Query = "Name_Orig='" + ACLObject + "'";                
                
                DataRow[] drObjects = frmMain.dtObjects.Select(Query);
                if (drObjects.Length == 1)
                {                    
                    if (CPObject == "")
                    {
                        CPObject = drObjects[0][1].ToString();
                        //check if this object was already created
                        if (frmMain.DBEditObjectsCreated.Contains(";" + CPObject + ";"))
                        {
                            //do nothing as we already created dbedit for this
                        }
                        else
                        {
                            if (drObjects[0][4].ToString().Contains("-"))
                            {
                                string[] Range = drObjects[0][4].ToString().Split('-');
                                CreateDBEdit.CreateObjRange(frmMain.WorkDir, "1-objects.dbedit", drObjects[0][1].ToString(), Range[0].Trim(), Range[1].Trim(), drObjects[0][7].ToString(), frmMain.GenerateDeleteFiles);
                            }
                            else if (drObjects[0][5].ToString() == "255.255.255.255")
                            {
                                //this is a host object
                                CreateDBEdit.CreateHost(frmMain.WorkDir, "1-objects.dbedit", drObjects[0][1].ToString(), drObjects[0][4].ToString(), drObjects[0][7].ToString(), frmMain.GenerateDeleteFiles);
                            }
                            else if (IPFunctions.IsValidSubnet(drObjects[0][5].ToString()))
                            {
                                //this is a network
                                CreateDBEdit.CreateNetwork(frmMain.WorkDir, "1-objects.dbedit", drObjects[0][1].ToString(), drObjects[0][4].ToString(), drObjects[0][5].ToString(), drObjects[0][7].ToString(), frmMain.GenerateDeleteFiles);
                            }
                            else if (drObjects[0][3].ToString() == "group")
                            {
                                string[] Members = drObjects[0][6].ToString().Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                                string GroupMembers = ";";
                                foreach (string Member in Members)
                                {
                                    string CPmemberObjectName = GetCPObject(Member);
                                    GroupMembers = GroupMembers + CPmemberObjectName + ";";
                                }
                                //create the group object                                
                                if (frmMain.ReplaceASDMName)
                                {
                                    if (drObjects[0][1].ToString().StartsWith("DM_INLINE_"))
                                    {
                                        CPObject = GroupMembers;
                                    }
                                    else
                                    {
                                        CreateDBEdit.CreateObjectGroup(frmMain.WorkDir, "2-groups.dbedit", drObjects[0][1].ToString(), GroupMembers, drObjects[0][7].ToString(), frmMain.GenerateDeleteFiles);
                                    }
                                }
                                else
                                {
                                    CreateDBEdit.CreateObjectGroup(frmMain.WorkDir, "2-groups.dbedit", drObjects[0][1].ToString(), GroupMembers, drObjects[0][7].ToString(), frmMain.GenerateDeleteFiles);
                                }


                            }
                            else
                            {
                                //wtf object
                                frmMain.Instance.RTBWriteErrorDBEGen = " \nUnknown Object Type " + CPObject;
                            }
                            //CPObject = drObjects[0][2].ToString();
                            frmMain.DBEditObjectsCreated = frmMain.DBEditObjectsCreated + CPObject + ";";
                        }
                    }
                }
                else
                {
                    if (funcShared.Compare_Objects(drObjects))
                    {
                        CPObject = drObjects[0][1].ToString();
                        //check if we already have this object defined in teh cpinfo
                        //CPObject = CheckIfCPObjectAlreadyExistsinCpinfo(drObjects[0][4].ToString(), drObjects[0][5].ToString(), drObjects[0][3].ToString(), CurrentForm, drObjects[0][2].ToString());                    
                        //to be completed

                        //check if this object was already created
                        if (frmMain.DBEditObjectsCreated.Contains(";" + CPObject + ";"))
                        {
                            //do nothing as we already created dbedit for this
                        }
                        else
                        {
                            if (drObjects[0][4].ToString().Contains("-"))
                            {
                                string[] Range = drObjects[0][4].ToString().Split('-');
                                CreateDBEdit.CreateObjRange(frmMain.WorkDir, "1-objects.dbedit", drObjects[0][1].ToString(), Range[0].Trim(), Range[1].Trim(), drObjects[0][7].ToString(), frmMain.GenerateDeleteFiles);
                            }
                            else if (drObjects[0][5].ToString() == "255.255.255.255")
                            {
                                //this is a host object
                                CreateDBEdit.CreateHost(frmMain.WorkDir, "1-objects.dbedit", drObjects[0][1].ToString(), drObjects[0][4].ToString(), drObjects[0][7].ToString(), frmMain.GenerateDeleteFiles);
                            }
                            else if (IPFunctions.IsValidSubnet(drObjects[0][5].ToString()))
                            {
                                //this is a network
                                CreateDBEdit.CreateNetwork(frmMain.WorkDir, "1-objects.dbedit", drObjects[0][1].ToString(), drObjects[0][4].ToString(), drObjects[0][5].ToString(), drObjects[0][7].ToString(), frmMain.GenerateDeleteFiles);
                            }
                            else if (drObjects[0][3].ToString() == "group")
                            {
                                string[] Members = drObjects[0][6].ToString().Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                                string GroupMembers = ";";
                                foreach (string Member in Members)
                                {
                                    string CPmemberObjectName = GetCPObject(Member);
                                    GroupMembers = GroupMembers + CPmemberObjectName + ";";
                                }
                                //create the group object                                
                                if (frmMain.ReplaceASDMName)
                                {
                                    if (drObjects[0][1].ToString().StartsWith("DM_INLINE_"))
                                    {
                                        CPObject = GroupMembers;
                                    }
                                    else
                                    {
                                        CreateDBEdit.CreateObjectGroup(frmMain.WorkDir, "2-groups.dbedit", drObjects[0][1].ToString(), GroupMembers, drObjects[0][7].ToString(), frmMain.GenerateDeleteFiles);
                                    }
                                }
                                else
                                {
                                    CreateDBEdit.CreateObjectGroup(frmMain.WorkDir, "2-groups.dbedit", drObjects[0][1].ToString(), GroupMembers, drObjects[0][7].ToString(), frmMain.GenerateDeleteFiles);
                                }
                            }
                            else
                            {
                                //wtf object
                                frmMain.Instance.RTBWriteErrorDBEGen = " \nUnknown Object Type " + CPObject;
                            }
                            //CPObject = drObjects[0][2].ToString();
                            frmMain.DBEditObjectsCreated = frmMain.DBEditObjectsCreated + CPObject + ";";
                        }
                    }
                    else
                    {
                        //throw error
                        frmMain.Instance.RTBWriteErrorDBEGen = " \nMulitple objects with different properties are matching name: " + ACLObject;

                    }

                }
            }

            return CPObject;
        }
    }
}
