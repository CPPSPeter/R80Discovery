using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;

namespace Excel2CP
{
    class funcShared
    {
        public static DataTable csvToDataTable(string file, bool isRowOneHeader)
        {
            DataTable dtEmpty = new DataTable();
            try
            {
                DataTable csvDataTable = new DataTable();

                //no try/catch - add these in yourselfs or let exception happen
                String[] csvData = File.ReadAllLines(file);


                String[] headings = csvData[0].Split(',');
                int index = 0; //will be zero or one depending on isRowOneHeader

                if (isRowOneHeader) //if first record lists headers
                {
                    index = 1; //so we won’t take headings as data

                    //for each heading
                    for (int i = 0; i < headings.Length; i++)
                    {
                        //replace spaces with underscores for column names
                        headings[i] = headings[i].Replace(" ", "_").Trim();

                        //add a column for each heading
                        csvDataTable.Columns.Add(headings[i], typeof(string));
                    }
                }
                else //if no headers just go for col1, col2 etc.
                {
                    for (int i = 0; i < headings.Length; i++)
                    {
                        //create arbitary column names
                        csvDataTable.Columns.Add("col" + (i + 1).ToString(), typeof(string));
                    }
                }

                //populate the DataTable
                for (int i = index; i < csvData.Length; i++)
                {
                    //create new rows
                    DataRow row = csvDataTable.NewRow();
                    if (csvData[i].StartsWith("#"))
                    {
                        //this is a comment
                    }
                    else
                    {
                        for (int j = 0; j < headings.Length; j++)
                        {
                            //fill them
                            row[j] = csvData[i].Split(',')[j].Trim();
                        }
                        //add rows to over DataTable
                        csvDataTable.Rows.Add(row);
                    }
                }
                //return the CSV DataTable
                return csvDataTable;
            }
            catch (Exception)
            {
                return dtEmpty;
            }
        }

        public static bool IsNumeric(string StringValue)
        {
            int i = 0;
            bool result = int.TryParse(StringValue, out i);
            return result;
        }

        //Method to remove Duplicate value from DataTable

        public static void RemoveDuplicatesFromDataTable(ref DataTable table, List<string> keyColumns)
        {

            Dictionary<string, string> uniquenessDict = new Dictionary<string, string>(table.Rows.Count);
            StringBuilder stringBuilder = null;
            int rowIndex = 0;
            DataRow row;
            DataRowCollection rows = table.Rows;
            while (rowIndex < rows.Count - 1)
            {
                row = rows[rowIndex];
                stringBuilder = new StringBuilder();
                foreach (string colname in keyColumns)                
                {
                    stringBuilder.Append(((string)row[colname]));
                }
                if (uniquenessDict.ContainsKey(stringBuilder.ToString()))
                {
                    rows.Remove(row);
                }
                else
                {
                    uniquenessDict.Add(stringBuilder.ToString(), string.Empty);
                    rowIndex++;
                }
            }
        }

        public static bool Compare_Objects(DataRow[] drObjects)
        {
            bool AllSame = true;
            string IP = drObjects[0][4].ToString().ToLower();
            string Subnet = drObjects[0][5].ToString().ToLower();
            string Members = drObjects[0][6].ToString().ToLower();

            foreach (DataRow drObject in drObjects)
            {
                if (drObject[4].ToString().ToLower() == IP && drObject[5].ToString().ToLower() == Subnet && drObject[6].ToString().ToLower() == Members)
                {
                    //nothing to see
                }
                else
                {
                    AllSame = false;
                }
            }
            return AllSame;
        }

        public static bool Compare_Services(DataRow[] drServices)
        {
            bool AllSame = true;
            string SType = drServices[0][3].ToString();
            string Proto = drServices[0][4].ToString();
            string Port = drServices[0][5].ToString();

            foreach (DataRow drService in drServices)
            {
                if (drService[4].ToString() != Proto && drService[5].ToString() != Port && drService[3].ToString() != SType)
                {
                    AllSame = false;
                }
            }
            return AllSame;
        }

        public static string Get_CP_Service_Predefined(string ServiceName)
        {
            string DefServiceName = "";
            DataRow[] drPredefServ = frmMain.dtDefaultServices.Select("CPEquivalent ='" + ServiceName + "'");
            if (drPredefServ.Length > 0)
            {
                DefServiceName = drPredefServ[0][3].ToString();
            }
            return DefServiceName;
        }
    }
}
