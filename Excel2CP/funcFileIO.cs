using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Excel2CP
{
    class funcFileIO
    {
        public static void Write_to_File(string Folder, string FileName, string LineToWrite)
        {
            //get the file ready for append
            StreamWriter SW;
            SW = File.AppendText(@Folder + @"\" + FileName);
            SW.WriteLine(LineToWrite);
            SW.Close();
        }
    }
}
