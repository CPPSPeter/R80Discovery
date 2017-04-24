using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace Excel2CP
{
    class clsDataTables
    {
        public static void InitDataTables()
        {
            //defina the holding datatables
            frmMain.dtRawPolicy.Reset();
            frmMain.dtRawPolicy.Columns.Add("ID", typeof(int));
            frmMain.dtRawPolicy.Columns.Add("Heading");
            frmMain.dtRawPolicy.Columns.Add("Source");
            frmMain.dtRawPolicy.Columns.Add("Destination");
            frmMain.dtRawPolicy.Columns.Add("Protocol");
            frmMain.dtRawPolicy.Columns.Add("Port");
            frmMain.dtRawPolicy.Columns.Add("Action");
            frmMain.dtRawPolicy.Columns.Add("Comment");

            //parsed policy datatable
            frmMain.dtPolicy.Reset();
            frmMain.dtPolicy.Columns.Add("ID", typeof(int));
            frmMain.dtPolicy.Columns.Add("Heading");
            frmMain.dtPolicy.Columns.Add("Source");
            frmMain.dtPolicy.Columns.Add("Destination");            
            frmMain.dtPolicy.Columns.Add("Service");
            frmMain.dtPolicy.Columns.Add("Action");
            frmMain.dtPolicy.Columns.Add("Comment");
            frmMain.dtPolicy.Columns.Add("protocol");
            frmMain.dtPolicy.Columns.Add("flag");

            frmMain.dtObjects.Reset();
            frmMain.dtObjects.Columns.Add("ID", typeof(int));
            frmMain.dtObjects.Columns.Add("Name_Orig");
            frmMain.dtObjects.Columns.Add("Name_CP");
            frmMain.dtObjects.Columns.Add("Type");
            frmMain.dtObjects.Columns.Add("IP");
            frmMain.dtObjects.Columns.Add("Subnet");
            frmMain.dtObjects.Columns.Add("Members");
            frmMain.dtObjects.Columns.Add("Comment");

            frmMain.dtServices.Reset();
            frmMain.dtServices.Columns.Add("ID", typeof(int));
            frmMain.dtServices.Columns.Add("Name_Orig");
            frmMain.dtServices.Columns.Add("Name_CP");
            frmMain.dtServices.Columns.Add("Type");
            frmMain.dtServices.Columns.Add("Proto");
            frmMain.dtServices.Columns.Add("Port");
            frmMain.dtServices.Columns.Add("Members");            
            frmMain.dtServices.Columns.Add("Comment");
            frmMain.dtServices.Columns.Add("ProtocolGroup");
        }

        public static void InitDBEditDataTables()
        {
            frmMain.dtCPPolicy.Reset();
            frmMain.dtCPPolicy.Columns.Add("ID", typeof(int));
            frmMain.dtCPPolicy.Columns.Add("SRC");
            frmMain.dtCPPolicy.Columns.Add("DST");
            frmMain.dtCPPolicy.Columns.Add("SRV");
            frmMain.dtCPPolicy.Columns.Add("Action");
            frmMain.dtCPPolicy.Columns.Add("Log");
            frmMain.dtCPPolicy.Columns.Add("Comment");
            frmMain.dtCPPolicy.Columns.Add("Disabled");
            frmMain.dtCPPolicy.Columns.Add("Name");
        }
        

    }
}
