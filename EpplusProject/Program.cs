// See https://aka.ms/new-console-template for more information
using static System.Console;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;
using System.Text;
using System.Data;
using System.Drawing;

using (ExcelPackage Pkg = new ExcelPackage()){


ExcelWorksheet WSheet1 = Pkg.Workbook.Worksheets.Add("Locations");

using (ExcelRange Rng = WSheet1.Cells[1,1,11,11])
{
    ExcelTable table = WSheet1.Tables.Add(Rng, "Location");

    table.ShowFilter = false;
    Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
    Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
    Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
    Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
}

WSheet1.Cells[1, 1].Value = "LOCATION NAME";
WSheet1.Cells[1, 2].Value = "TAX RATE 1";
WSheet1.Cells[1, 3].Value = "TAX RATE 2";
WSheet1.Cells[1, 4].Value = "LABOUR TAX";
WSheet1.Cells[1, 5].Value = "CATEGORY";
WSheet1.Cells[1, 6].Value = "STREET";
WSheet1.Cells[1, 7].Value = "CITY";
WSheet1.Cells[1, 8].Value = "STATE";
WSheet1.Cells[1, 9].Value = "COUNTRY";
WSheet1.Cells[1, 10].Value = "POST CODE";
WSheet1.Cells[1, 11].Value = "CAMS WAREHOUSE LIST";

WSheet1.Cells["A2"].Value = "string";
WSheet1.Cells["A3"].Value = "number2";
WSheet1.Cells["A4"].Value = "string";
WSheet1.Cells["A5"].Value = "Camp site";
WSheet1.Cells["A6"].Value = "string";
WSheet1.Cells["A7"].Value = "string";
WSheet1.Cells["A8"].Value = "string";
WSheet1.Cells["A9"].Value = "string";
WSheet1.Cells["A10"].Value = "string";
WSheet1.Cells["A11"].Value = "string";

WSheet1.Cells["B2"].Value = 0;
WSheet1.Cells["B3"].Value = 0;
WSheet1.Cells["B4"].Value = 0;
WSheet1.Cells["B5"].Value = 0;
WSheet1.Cells["B6"].Value = 0;
WSheet1.Cells["B7"].Value = 0;
WSheet1.Cells["B8"].Value = 0;
WSheet1.Cells["B9"].Value = 0;
WSheet1.Cells["B10"].Value = 0;
WSheet1.Cells["B11"].Value = 0;

WSheet1.Cells["C2"].Value = 0;
WSheet1.Cells["C3"].Value = 0;
WSheet1.Cells["C4"].Value = 0;
WSheet1.Cells["C5"].Value = 0;
WSheet1.Cells["C6"].Value = 0;
WSheet1.Cells["C7"].Value = 0;
WSheet1.Cells["C8"].Value = 0;
WSheet1.Cells["C9"].Value = 0;
WSheet1.Cells["C10"].Value = 0;
WSheet1.Cells["C11"].Value = 0;

WSheet1.Cells["D2"].Value = 0;
WSheet1.Cells["D3"].Value = 0;
WSheet1.Cells["D4"].Value = 0;
WSheet1.Cells["D5"].Value = 0;
WSheet1.Cells["D6"].Value = 0;
WSheet1.Cells["D7"].Value = 0;
WSheet1.Cells["D8"].Value = 0;
WSheet1.Cells["D9"].Value = 0;
WSheet1.Cells["D10"].Value = 0;
WSheet1.Cells["D11"].Value = 0;

WSheet1.Cells["E2"].Value = "string";
WSheet1.Cells["E3"].Value = "string";
WSheet1.Cells["E4"].Value = "string";
WSheet1.Cells["E5"].Value = "string";
WSheet1.Cells["E6"].Value = "string";
WSheet1.Cells["E7"].Value = "string";
WSheet1.Cells["E8"].Value = "string";
WSheet1.Cells["E9"].Value = "string";
WSheet1.Cells["E10"].Value = "string";
WSheet1.Cells["E11"].Value = "string";

WSheet1.Cells["F2"].Value = "string";
WSheet1.Cells["F3"].Value = "string";
WSheet1.Cells["F4"].Value = "string";
WSheet1.Cells["F5"].Value = "string";
WSheet1.Cells["F6"].Value = "string";
WSheet1.Cells["F7"].Value = "string";
WSheet1.Cells["F8"].Value = "string";
WSheet1.Cells["F9"].Value = "string";
WSheet1.Cells["F10"].Value = "string";
WSheet1.Cells["F11"].Value = "string";

WSheet1.Cells["G2"].Value = "string";
WSheet1.Cells["G3"].Value = "string";
WSheet1.Cells["G4"].Value = "string";
WSheet1.Cells["G5"].Value = "string";
WSheet1.Cells["G6"].Value = "string";
WSheet1.Cells["G7"].Value = "string";
WSheet1.Cells["G8"].Value = "string";
WSheet1.Cells["G9"].Value = "string";
WSheet1.Cells["G10"].Value = "string";
WSheet1.Cells["G11"].Value = "string";

WSheet1.Cells["H2"].Value = "string";
WSheet1.Cells["H3"].Value = "string";
WSheet1.Cells["H4"].Value = "string";
WSheet1.Cells["H5"].Value = "string";
WSheet1.Cells["H6"].Value = "string";
WSheet1.Cells["H7"].Value = "string";
WSheet1.Cells["H8"].Value = "string";
WSheet1.Cells["H9"].Value = "string";
WSheet1.Cells["H10"].Value = "string";
WSheet1.Cells["H11"].Value = "string";

WSheet1.Cells["I2"].Value = "string";
WSheet1.Cells["I3"].Value = "string";
WSheet1.Cells["I4"].Value = "string";
WSheet1.Cells["I5"].Value = "string";
WSheet1.Cells["I6"].Value = "string";
WSheet1.Cells["I7"].Value = "string";
WSheet1.Cells["I8"].Value = "string";
WSheet1.Cells["I9"].Value = "string";
WSheet1.Cells["I10"].Value = "string";
WSheet1.Cells["I11"].Value = "string";

WSheet1.Cells["J2"].Value = "string";
WSheet1.Cells["J3"].Value = "string";
WSheet1.Cells["J4"].Value = "string";
WSheet1.Cells["J5"].Value = "string";
WSheet1.Cells["J6"].Value = "string";
WSheet1.Cells["J7"].Value = "string";
WSheet1.Cells["J8"].Value = "string";
WSheet1.Cells["J9"].Value = "string";
WSheet1.Cells["J10"].Value = "string";
WSheet1.Cells["J11"].Value = "string";

WSheet1.Cells["K2"].Value = "string";
WSheet1.Cells["K3"].Value = "string";
WSheet1.Cells["K4"].Value = "string";
WSheet1.Cells["K5"].Value = "string";
WSheet1.Cells["K6"].Value = "string";
WSheet1.Cells["K7"].Value = "string";
WSheet1.Cells["K8"].Value = "string";
WSheet1.Cells["K9"].Value = "string";
WSheet1.Cells["K10"].Value = "string";
WSheet1.Cells["K11"].Value = "string";

WSheet1.Cells.AutoFitColumns(30);


WSheet1.Protection.IsProtected = false;
WSheet1.Protection.AllowSelectLockedCells = false;
Pkg.SaveAs(new FileInfo(@"C:\Users\HUMBLE\Desktop\Locations.Xlsx"));
}
