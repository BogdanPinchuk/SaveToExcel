// https://docs.microsoft.com/ru-ru/dotnet/csharp/programming-guide/interop/how-to-access-office-onterop-objects
// https://docs.microsoft.com/ru-ru/dotnet/csharp/programming-guide/interop/walkthrough-office-programming

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;

namespace SaveToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // case 1, for save data into excel file
        private void Button1_Click(object sender, EventArgs e)
        {
            var bankA = new List<Account>
            {
                new Account
                {
                    ID = 345678,
                    Balance = 541.2700051515131
                },
                new Account
                {
                    ID = 1230221,
                    Balance = -127.4400051515131
                }
            };

            DisplayInExcel(bankA);
        }

        private void DisplayInExcel(IEnumerable<Account> accounts)
        {
            var eApp = new Excel.Application
            {
                Visible = false
            };

            var eWbook = eApp.Workbooks.Add();
            var eWsheet = eWbook.ActiveSheet;
            var eWsheet1 = eWbook.Worksheets.Add(After: eWsheet);

            //Excel._Worksheet eWsheet = (Excel.Worksheet)eApp.ActiveSheet;

            eWsheet.Name = "My_sheet";
            eWsheet1.Name = "My_sheet_2";

            //eWsheet.Cells[1, "A"] = "ID Number";
            //eWsheet.Cells[1, "B"] = "Current Balance";
            eWsheet.Cells[1, 1] = "ID Number";
            eWsheet.Cells[1, 2] = "Current Balance";

            // do active list
            eWsheet.Activate();

            var row = 1;
            foreach (var acct in accounts)
            {
                row++;
                eWsheet.Cells[row, "A"] = acct.ID;
                eWsheet.Cells[row, "B"] = acct.Balance;
            }

            eWsheet.Columns.AutoFit();
            //eWsheet.Columns.NumberFormat = "0.000";

            //eWsheet.Range[1, 1].NumberFormat = Excel.XlColumnDataType.xlTextFormat;
            //eWsheet.Range[1, 2].NumberFormat = Excel.XlColumnDataType.xlTextFormat;

            string path = Directory.GetCurrentDirectory(),
                tPath = @"\save data\",
                name = "Case_1";

            int id = 1;

            // choose path
            if (saveFD.ShowDialog() == DialogResult.OK)
            {
                id = saveFD.FilterIndex;
                path = saveFD.FileName;
            }
            else
            {
                return;
            }

            // create new path
            //Directory.CreateDirectory(path + tPath);
            switch (id)
            {
                case 1: // xlsx
                    eWbook.SaveAs(Filename: path, FileFormat: Excel.XlFileFormat.xlWorkbookDefault);
                    break;
                case 2: // xls
                    eWbook.SaveAs(Filename: path, FileFormat: Excel.XlFileFormat.xlWorkbookNormal);
                    break;
                default:
                    break;
            }

            eWbook.Close();
            eApp.Quit();
        }
    }
}
