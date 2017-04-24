using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace Excel2CP
{
    public partial class frmMain : Form
    {
        private static frmMain instance;

        public static frmMain Instance
        {
            get { return instance; }
        }

        public static string ConfigFile = "";
        public static string WorkDir;
        string Version = "";
        public static string IgnoredLinesFile = "IgnoredLines.log";
        public static string ErrorLogFile = "ParserError.log";
        public static string ParsedLogFile = "ParsedFiles.log";
        public static bool Debugging = true;
        public static bool LogParsedFiles = false;
        public static bool LogIgnoredLines = true;
        public static string ErrorDetail = "";
        public static bool DisplayWarnings = false;

        public static int ObjID = 0;
        public static int ServID = 0;

        //datatables
        public static DataTable dtDefaultServices = new DataTable();
        public static DataTable dtRawPolicy = new DataTable();
        public static DataTable dtPolicy = new DataTable();
        public static DataTable dtObjects = new DataTable();
        public static DataTable dtServices = new DataTable();
        public static DataTable dtCPPolicy = new DataTable();

        public static bool STperACG = true;
        public static bool LogEachhRule = true;
        public static bool ReplaceASDMName = true;
        public static string DBEditObjectsCreated = "";
        public static string DBEditServceCreated = "";

        public static bool GenerateDeleteFiles = false;

        public frmMain()
        {
            InitializeComponent();
            instance = this;
            dtDefaultServices = funcShared.csvToDataTable(System.Environment.CurrentDirectory + @"\" + "Cisco-serv.csv", true);
            if (dtDefaultServices.Rows.Count == 0)
            {
                rtbProgressBox.AppendText("Error: Failed to load predefined services");
            }
            else
            {
                BindPredefServices();
            }
        }

        public string RTBWriteState
        {
            set
            {
                Font fBold = new Font("Tahoma", 8, FontStyle.Bold);
                Font fNorm = new Font("Tahoma", 8);

                rtbProgressBox.SelectionFont = fBold;
                rtbProgressBox.SelectionColor = Color.Blue;
                rtbProgressBox.AppendText(value);
                rtbProgressBox.SelectionColor = Color.Black;
                rtbProgressBox.SelectionFont = fNorm;
            }
        }

        public string RTBWriteError
        {
            set
            {
                rtbProgressBox.SelectionColor = Color.Red;
                rtbProgressBox.AppendText(value);
                rtbProgressBox.SelectionColor = Color.Black;
            }
        }

        public string RTBWriteWarning
        {
            set
            {
                rtbProgressBox.SelectionColor = Color.Orange;
                rtbProgressBox.AppendText(value);
                rtbProgressBox.SelectionColor = Color.Black;
            }
        }

        public string RTBWriteStateDBEGen
        {
            set
            {
                Font fBold = new Font("Tahoma", 8, FontStyle.Bold);
                Font fNorm = new Font("Tahoma", 8);

                rtbDBEditProc.SelectionFont = fBold;
                rtbDBEditProc.SelectionColor = Color.Blue;
                rtbDBEditProc.AppendText(value);
                rtbDBEditProc.SelectionColor = Color.Black;
                rtbDBEditProc.SelectionFont = fNorm;
            }
        }

        public string RTBWriteErrorDBEGen
        {
            set
            {
                rtbDBEditProc.SelectionColor = Color.Red;
                rtbDBEditProc.AppendText(value);
                rtbDBEditProc.SelectionColor = Color.Black;
            }
        }

        public string RTBWriteTextDBEGen
        {
            set
            {
                rtbDBEditProc.AppendText(value);
            }
        }

        private void cmdCfgFile_Click(object sender, EventArgs e)
        {
            Version = "";
            OpenFileDialog OpenFile = new OpenFileDialog();
            OpenFile.Filter = "Cisco Config File|*.*";
            OpenFile.Title = "Select Cisco Config File to convert";

            if (OpenFile.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = OpenFile.FileName;
                ConfigFile = OpenFile.FileName;
                WorkDir = System.IO.Path.GetDirectoryName(ConfigFile) + @"\";
                clsDataTables.InitDataTables();

                //lets open the excel file
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                Excel.Range range;

                string str;
                int rCnt = 0;
                int cCnt = 0;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(ConfigFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                //range = xlWorkSheet.UsedRange;

                //for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                //{
                //    for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                //    {
                //        str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                //        MessageBox.Show(str);
                //    }
                //}
                //releaseObject(xlWorkSheet);

                //parse objects
                Excel.Worksheet xlWSGroups;
                xlWSGroups = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(4);

                range = xlWSGroups.UsedRange;

                int rowGCnt = 0;
                
                string GroupName = "";
                string GroupType = "";
                string GroupProtocol = "";
                string GroupDescription = "";
                string GroupMembers = "";
                bool GrpService = false;
                bool GrpObject = false;

                for (rowGCnt = 1; rCnt <= range.Rows.Count; rowGCnt++)
                {
                    Application.DoEvents();
                    //lets start the parsing
                    string ColA = (string)(range.Cells[rowGCnt, 1] as Excel.Range).Value2;
                    string ColB = (string)(range.Cells[rowGCnt, 2] as Excel.Range).Value2;
                    object ColCObj = (object)(range.Cells[rowGCnt, 3] as Excel.Range).Value2;
                    object ColDObj = (object)(range.Cells[rowGCnt, 4] as Excel.Range).Value2;
                    object ColEObj = (object)(range.Cells[rowGCnt, 5] as Excel.Range).Value2;
                    string ColC = "";
                    string ColD = "";
                    string ColE = "";
                    if (ColCObj != null)
                    {
                        ColC = ColCObj.ToString();
                    }
                    if (ColDObj != null)
                    { 
                        ColD = ColDObj.ToString();
                    }
                    if (ColEObj != null)
                    {
                        ColE = ColEObj.ToString();
                    }

                    if(ColA==null && ColB==null)
                    {
                        rtbProgressBox.AppendText("\nThe group worksheet contained " + rowGCnt.ToString() + " lines of config.");
                        //write to datatable
                        if (GrpObject)
                        {
                            AddObjectsToDT(GroupName, "group", "", "", GroupMembers, GroupDescription);
                        }
                        if (GrpService)
                        {
                            AddServiceToDT(GroupName, "group", "", "", GroupMembers, GroupDescription, GroupProtocol);
                        }
                        break;
                    }
                    if (ColA == "object-group")
                    {
                        //write to datatable the last group parsed
                        if (GrpObject)
                        {
                            AddObjectsToDT(GroupName, "group", "", "", GroupMembers, GroupDescription);                      
                        }
                        if (GrpService)
                        {
                            AddServiceToDT(GroupName, "group", "", "", GroupMembers, GroupDescription, GroupProtocol);
                        }
                        //reset all values
                        GroupName = "";
                        GroupType = "";
                        GroupProtocol = "";
                        GroupDescription = "";
                        GroupMembers = "";


                        GrpService = false;
                        GrpObject = false;
                        //we have a new object definition
                        GroupType = (string)(range.Cells[rowGCnt, 2] as Excel.Range).Value2;
                        GroupName = (string)(range.Cells[rowGCnt, 3] as Excel.Range).Value2;
                        GroupProtocol = (string)(range.Cells[rowGCnt, 4] as Excel.Range).Value2;
                        if (GroupType == "service")
                        {
                            GrpService = true;
                        }
                        else if (GroupType == "network")
                        {
                            GrpObject = true;
                            if (GroupProtocol != null)
                            {
                                rtbProgressBox.AppendText("\nGrooup object " + GroupName + " network seem to have some additional info defined: " + GroupProtocol);
                            }
                        }
                        else
                        {
                            rtbProgressBox.AppendText("\nInvalid Group Type: " + GroupName + " type: " + GroupType);   
                        }                        
                    }
                    else if(ColA==null && ColB!=null)
                    {
                        //get description
                        if (ColB == "description")
                        {
                            for (int grpcCnt = 3; grpcCnt <= range.Columns.Count; grpcCnt++)
                            {
                                object DescColString = (object)(range.Cells[rowGCnt, grpcCnt] as Excel.Range).Value2;
                                if (DescColString != null)
                                {
                                    GroupDescription = GroupDescription + DescColString.ToString() + " ";
                                }
                            }
                            GroupDescription = GroupDescription.Trim();
                        }
                        else
                        {
                            //get type related info
                            if (GrpObject)
                            {
                                if (ColB == "network-object")
                                {
                                    if (ColC == "host")
                                    {
                                        if (IPFunctions.IsValidIP(ColD))
                                        {
                                            AddObjectsToDT("host_" + ColD, "host", ColD, "255.255.255.255", "", "");
                                            GroupMembers = GroupMembers + "host_" + ColD + ";";
                                        }
                                        else
                                        {
                                            rtbProgressBox.AppendText("\nNetwork - Invalid Group member definition : " + GroupName + " host member IP : " + ColD);
                                        }
                                    }
                                    else if (IPFunctions.IsValidIP(ColC.Trim()))
                                    {
                                        //this is most likely a network
                                        if (IPFunctions.IsValidSubnet(ColD))
                                        {
                                            int CIDR = IPFunctions.GetSubnetMask(ColD);
                                            if (ColD == "255.255.255.255")
                                            {
                                                AddObjectsToDT("host_" + ColC, "host", ColC, "255.255.255.255", "", "");
                                                GroupMembers = GroupMembers + "host_" + ColC + ";";
                                            }
                                            else
                                            {
                                                AddObjectsToDT("net_" + ColC + "_" + CIDR.ToString(), "network", ColC, ColD, "", "");
                                                GroupMembers = GroupMembers + "net_" + ColC + "_" + CIDR.ToString() + ";";
                                            }
                                        }
                                        else
                                        {
                                            rtbProgressBox.AppendText("\nNetwork - Invalid Group member definition : " + GroupName + " host member subnet : " + ColD);
                                        }
                                    }
                                    else
                                    {
                                        rtbProgressBox.AppendText("\nNetwork - Invalid Group member definition : " + GroupName + " host member ip/subnet : " + ColC + "/" + ColD);
                                    }
                                }
                                else if (ColB == "group-object")
                                {
                                    GroupMembers = GroupMembers + ColC + ";";   
                                }
                                else
                                {
                                    rtbProgressBox.AppendText("\nNetwork - Invalid Group member definition : " + GroupName + " definition: " + ColB);
                                }
                            }
                            else if (GrpService)
                            {
                                if (ColB == "port-object")
                                {
                                    if (ColC == "eq")
                                    {
                                        if (funcShared.IsNumeric(ColD))
                                        {
                                            if (GroupProtocol == "tcp")
                                            {
                                                AddServiceToDT("tcp_" + ColD, "", "tcp", ColD, "", "", "");
                                                GroupMembers = GroupMembers + "tcp_" + ColD + ";";
                                            }
                                            else if (GroupProtocol == "udp")
                                            {
                                                AddServiceToDT("udp_" + ColD, "", "udp", ColD, "", "", "");
                                                GroupMembers = GroupMembers + "udp_" + ColD + ";";
                                            }
                                            else if (GroupProtocol == "tcp-udp" || GroupProtocol == "tcp;udp")
                                            {
                                                AddServiceToDT("tcp_" + ColD, "", "tcp", ColD, "", "", "");
                                                GroupMembers = GroupMembers + "tcp_" + ColD + ";";
                                                AddServiceToDT("udp_" + ColD, "", "udp", ColD, "", "", "");
                                                GroupMembers = GroupMembers + "udp_" + ColD + ";";
                                            }
                                            else
                                            {
                                                rtbProgressBox.AppendText("\nService - Invalid Group protocol type : " + GroupName + " type: " + GroupProtocol);
                                            }
                                        }
                                        else
                                        {
                                            if (GroupProtocol == "tcp")
                                            {
                                                AddServiceToDT(ColD, "", "tcp", ColD, "", "", "");
                                                GroupMembers = GroupMembers + ColD + ";";
                                            }
                                            else if (GroupProtocol == "udp")
                                            {
                                                AddServiceToDT(ColD, "", "udp", ColD, "", "", "");
                                                GroupMembers = GroupMembers + ColD + ";";
                                            }
                                            else if (GroupProtocol == "tcp-udp" || GroupProtocol == "tcp;udp")
                                            {
                                                AddServiceToDT(ColD, "", "tcp", ColD, "", "", "");
                                                GroupMembers = GroupMembers + ColD + ";";
                                                AddServiceToDT(ColD, "", "udp", ColD, "", "", "");
                                                GroupMembers = GroupMembers + ColD + ";";
                                            }
                                            else
                                            {
                                                rtbProgressBox.AppendText("\nService - Invalid Group protocol type : " + GroupName + " type: " + GroupProtocol);
                                            }
                                        }
                                    }
                                    else if (ColC == "range")
                                    {
                                        if (ColD == ColE)
                                        {
                                            if (GroupProtocol == "tcp")
                                            {
                                                AddServiceToDT(ColD, "", "tcp", ColD, "", "", "");
                                                GroupMembers = GroupMembers + ColD + ";";
                                            }
                                            else if (GroupProtocol == "udp")
                                            {
                                                AddServiceToDT(ColD, "", "udp", ColD, "", "", "");
                                                GroupMembers = GroupMembers + ColD + ";";
                                            }
                                            else if (GroupProtocol == "tcp-udp" || GroupProtocol == "tcp;udp")
                                            {
                                                AddServiceToDT(ColD, "", "tcp", ColD, "", "", "");
                                                GroupMembers = GroupMembers + ColD + ";";
                                                AddServiceToDT(ColD, "", "udp", ColD, "", "", "");
                                                GroupMembers = GroupMembers + ColD + ";";
                                            }
                                            else
                                            {
                                                rtbProgressBox.AppendText("\nService - Invalid Group protocol type : " + GroupName + " type: " + GroupProtocol);
                                            }
                                        }
                                        else
                                        {
                                            if (GroupProtocol == "tcp")
                                            {
                                                AddServiceToDT("tcp_" + ColD + "-" + ColE, "", "tcp", ColD + "-" + ColE, "", "", "");
                                                GroupMembers = GroupMembers + "tcp_" + ColD + "-" + ColE + ";";
                                            }
                                            else if (GroupProtocol == "udp")
                                            {
                                                AddServiceToDT("udp_" + ColD + "-" + ColE, "", "udp", ColD + "-" + ColE, "", "", "");
                                                GroupMembers = GroupMembers + "udp_" + ColD + "-" + ColE + ";";
                                            }
                                            else if (GroupProtocol == "tcp-udp" || GroupProtocol == "tcp;udp")
                                            {
                                                AddServiceToDT("tcp_" + ColD + "-" + ColE, "", "tcp", ColD + "-" + ColE, "", "", "");
                                                GroupMembers = GroupMembers + "tcp_" + ColD + "-" + ColE + ";";
                                                AddServiceToDT("udp_" + ColD + "-" + ColE, "", "udp", ColD + "-" + ColE, "", "", "");
                                                GroupMembers = GroupMembers + "udp_" + ColD + "-" + ColE + ";";
                                            }
                                            else
                                            {
                                                rtbProgressBox.AppendText("\nService - Invalid Group protocol type : " + GroupName + " type: " + GroupProtocol);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        rtbProgressBox.AppendText("\nService - Invalid Group member: " + GroupName + " member type: " + ColC);
                                    }
                                }
                                else if (ColB == "group-object")
                                {
                                    GroupMembers = GroupMembers + ColC + ";";   
                                }
                                else
                                {
                                    rtbProgressBox.AppendText("\nService - Invalid Group member: " + GroupName + " member type: " + ColB);
                                }

                            }
                            else
                            {
                                rtbProgressBox.AppendText("\nInvalid Group Type: " + GroupName + " type: " + GroupType);
                            }
                        }
                    }

                }







                //parse Policy
                Excel.Worksheet xlWSPolicy;
                xlWSPolicy = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

                
                int rowCnt = 0;
                
                range = xlWSPolicy.UsedRange;
                if (range.Columns.Count != 7)
                {
                    rtbProgressBox.AppendText("\nThe policy worksheet should only have 7 columns.\nNumber of columns found = " + range.Columns.Count.ToString());
                }

                int PolicyID = 0;
                string Heading = "";
                string Source = "";
                string Destination = "";
                string Protocol = "";
                string Port = "";
                string Action = "";
                string Comment = "";
                
                for (rowCnt = 4; rCnt <= range.Rows.Count; rowCnt++)
                {
                    Application.DoEvents();
                    Heading = (string)(range.Cells[rowCnt, 1] as Excel.Range).Value2;
                    Source = (string)(range.Cells[rowCnt, 2] as Excel.Range).Value2;
                    Destination = (string)(range.Cells[rowCnt, 3] as Excel.Range).Value2;
                    Protocol = (string)(range.Cells[rowCnt, 4] as Excel.Range).Value2;
                    Port = (string)(range.Cells[rowCnt, 5] as Excel.Range).Value2;
                    Action = (string)(range.Cells[rowCnt, 6] as Excel.Range).Value2;
                    Comment = (string)(range.Cells[rowCnt, 7] as Excel.Range).Value2;

                    if (Heading == null && Source == null && Destination == null && Protocol == null && Port == null && Action == null && Comment == null)
                    {
                        rtbProgressBox.AppendText("\nThe policy worksheet contained " + rowCnt.ToString() + " lines of config.");
                        break;
                    }

                    //load to datatable
                    DataRow drPolRow = frmMain.dtRawPolicy.NewRow();
                    drPolRow[0] = PolicyID;
                    drPolRow[1] = Heading;
                    drPolRow[2] = Source;
                    drPolRow[3] = Destination;
                    drPolRow[4] = Protocol;
                    drPolRow[5] = Port;
                    drPolRow[6] = Action;
                    drPolRow[7] = Comment;
                    frmMain.dtRawPolicy.Rows.Add(drPolRow);
                    PolicyID++;                    
                }
                
                //parse the policy to right form
                int RuleNumber = 0;
                foreach (DataRow drRule in dtRawPolicy.Rows)
                {
                    funcPolParser.ParseRule(drRule, RuleNumber);
                    RuleNumber++;
                }




                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                releaseObject(xlWSPolicy);
                //releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                //remove duplicates from the services table
                List<string> keyColumns = new List<string>();
                keyColumns.Add("Name_Orig");
                keyColumns.Add("Name_CP");
                keyColumns.Add("Type");
                keyColumns.Add("Proto");
                keyColumns.Add("Port");
                keyColumns.Add("Members");
                keyColumns.Add("Comment");
                keyColumns.Add("ProtocolGroup");
                funcShared.RemoveDuplicatesFromDataTable(ref dtServices, keyColumns);

                
                ShowObjects();
                ShowServices();
                ShowPolicyRaw();
                ShowPolicyParsed();
            }

        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void AddObjectsToDT(string Name, string type, string IP, string Subnet, string Members, string Comment)
        {
            DataRow drObj = frmMain.dtObjects.NewRow();
            drObj[0] = frmMain.ObjID;
            drObj[1] = Name;
            drObj[2] = "";
            drObj[3] = type;
            drObj[4] = IP;
            drObj[5] = Subnet;
            drObj[6] = Members;
            drObj[7] = Comment;
            frmMain.dtObjects.Rows.Add(drObj);
            frmMain.ObjID++;              
        }

        public static void AddServiceToDT(string Name, string type, string proto, string port, string Members, string Comment, string GroupProto)
        {
            DataRow drServ = frmMain.dtServices.NewRow();
            drServ[0] = frmMain.ServID;
            drServ[1] = Name;
            drServ[2] = "";
            drServ[3] = type;
            drServ[4] = proto;
            drServ[5] = port;
            drServ[6] = Members;
            drServ[7] = Comment;
            drServ[8] = GroupProto;
            frmMain.dtServices.Rows.Add(drServ);
            frmMain.ServID++;
        }

        private void BindPredefServices()
        {
            try
            {
                BindingSource myBindingSource = new BindingSource();
                dgvPredefServices.DataSource = myBindingSource;
                DataView myDataView = new DataView(dtDefaultServices);
                myBindingSource.DataSource = myDataView;

            }
            catch
            {
                MessageBox.Show("Error displaying the parsed predefined services", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowObjects()
        {
            try
            {
                BindingSource myBindingSource = new BindingSource();
                dgvObjects.DataSource = myBindingSource;
                DataView myDataView = new DataView(dtObjects);
                myBindingSource.DataSource = myDataView;
            }
            catch
            {
                MessageBox.Show("Error displaying the parsed Object info", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowServices()
        {
            try
            {
                BindingSource myBindingSource = new BindingSource();
                dgvServices.DataSource = myBindingSource;
                DataView myDataView = new DataView(dtServices);
                myBindingSource.DataSource = myDataView;
            }
            catch
            {
                MessageBox.Show("Error displaying the parsed Services info", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowPolicyRaw()
        {
            try
            {
                BindingSource myBindingSource = new BindingSource();
                dgvPolicyRaw.DataSource = myBindingSource;
                DataView myDataView = new DataView(dtRawPolicy);
                myBindingSource.DataSource = myDataView;
            }
            catch
            {
                MessageBox.Show("Error displaying the parsed Raw Policy info", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowPolicyParsed()
        {
            try
            {
                BindingSource myBindingSource = new BindingSource();
                dgvPolicy.DataSource = myBindingSource;
                DataView myDataView = new DataView(dtPolicy);
                myBindingSource.DataSource = myDataView;
            }
            catch
            {
                MessageBox.Show("Error displaying the parsed Policy info", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowCPPolicy()
        {
            try
            {
                BindingSource myBindingSource = new BindingSource();
                dgvCPPolicy.DataSource = myBindingSource;
                DataView myDataView = new DataView(dtCPPolicy);
                myBindingSource.DataSource = myDataView;
            }
            catch
            {
                MessageBox.Show("Error displaying the resulting CP Policy", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDBEdit_Click(object sender, EventArgs e)
        {
            frmMain.Instance.RTBWriteStateDBEGen = "\n\nGenerating DBEdits based on required policy";
            frmMain.DBEditObjectsCreated = ";";
            frmMain.DBEditServceCreated = ";";
            rtbDBEditProc.Text = "";
            clsDataTables.InitDBEditDataTables();
            bool success = funcDBEDITGeneration.Generate_ACL_Rule(txtPolNameDBEdit.Text, chkReplaceDM_Inline.Checked);
            ShowCPPolicy();
        }
        
    }
}
