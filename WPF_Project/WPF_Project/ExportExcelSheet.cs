using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using static WPF_Project.ControlVoice;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace WPF_Project
{
    public class ExportExcelSheet
    {
        public ExportExcelSheet(DataGrid dataGrid, DateTime dt)
        {
            Application excel = new Application();
            excel.Visible = true;
            Worksheet sheet1 = excel.Workbooks.Add(Missing.Value).Sheets[1];

            sheet1.Range["A1:C1"].Merge();
            sheet1.Range["A2:C2"].Merge();

            string title = "Sample List in ExcelSheet";

            sheet1.Range["A1:C1"].Value = title;
            sheet1.Range["A1:C1"].Font.Bold = true;
            sheet1.Range["A2:C2"].Value =dt.ToShortDateString();

            for (int j = 0; j < dataGrid.Columns.Count; j++)
            {
                Range myRange = (Range)sheet1.Cells[4, j + 1];
                myRange.Font.Bold = true;
                string header = dataGrid.Columns[j].Header.ToString();
                sheet1.Columns[j + 1].ColumnWidth = header.Length + 5;
                myRange.Value2 = header;
            }
            for (int j = 0; j < dataGrid.Items.Count;j++)
            {
                var item = dataGrid.Items[j] as Product;
                sheet1.Cells[j + 5, 1].Value = item.ProductId;
                sheet1.Cells[j + 5, 2].Value = item.ProductName;
                sheet1.Cells[j + 5, 3].Value = item.Price;
                sheet1.Cells[j + 5, 4].value = item.Quantity;
            }
              

        }
    }
}
