using System.Data;

namespace OpenXmlAbstractions.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExcelGenerationLibrary.PopulateSpreadSheetWithDataTableTemplate(
            //    "ProjectStatusReportTemplate.xlsx", "test.xlsx", "N&S", "A3", GetTestDataTable(), null);
            ExcelGenerationLibrary.PopulateSpreadSheetWithXMLTemplate("xmltest.xlsx", "test2.xlsx", "Sheet1", "A1", "<excelexport><columns><column>t1</column></columns><row><cell>asdf</cell></row></excelexport>");
        }

        public static DataTable GetTestDataTable()
        {
            var dt = new DataTable("TestTable");

            var c = 1;
            while (c <= 13)
            {
                var dc = new DataColumn();
                dc.ColumnName = "C" + c;
                dt.Columns.Add(dc);
                c++;
            }

            var i = 0;
            while (i < 1000)
            {
                var dr3 = dt.NewRow();
                dr3["C1"] = null;
                dr3["C2"] = "";
                dr3["C3"] = "Baldwin";
                dr3["C4"] = "0200302a";
                dr3["C5"] = "0200302a";
                dr3["C6"] = "0200302a";
                dr3["C7"] = "0200302a";
                dr3["C8"] = "";
                dr3["C9"] = "G " + (char)10+"asdf"+(char)10+"werer";
                dr3["C10"] = "0200302a";
                dr3["C11"] = "R 0200302a";
                dr3["C12"] = "0200302a";
                dr3["C13"] = "0200302a";
                dt.Rows.Add(dr3);
                i++;
            }

            return dt;
        }
    }
}