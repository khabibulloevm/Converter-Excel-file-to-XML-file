using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Bytescout.Spreadsheet;

namespace ConvertXLS_to_XML
{
    class ConvertExceltoXml
    {
        static void Main(string[] args)
        {
            Spreadsheet document = new Spreadsheet();
            document.LoadFromFile("D:\\Viggggi.xlsx");

            // delete output file if exists already
            if (File.Exists("Viggggi.xml"))
            {
                File.Delete("Viggggi.xml");
            }

            document.Workbook.Worksheets[0].SaveAsXML("Viggggi.xml");
            document.Close();
        }
    }
}
