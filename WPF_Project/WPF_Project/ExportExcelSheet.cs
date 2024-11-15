using System;
using System.Reflection;
using System.Windows.Controls;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace WPF_Project
{
    public class ExportExcelSheet
    {
        public ExportExcelSheet(DataGrid dataGrid, DateTime dt, string filePath)
        {
            Application excel = new Application();
            excel.Visible = true;
            Workbook workbook = excel.Workbooks.Open(filePath); // Mở Workbook đã tồn tại
            Worksheet sheet1 = workbook.Sheets[0];

            sheet1.Range["A1:C1"].Merge();
            sheet1.Range["A2:C2"].Merge();

            string title = "Sample List in ExcelSheet";

            sheet1.Range["A1:C1"].Value = title;
            sheet1.Range["A1:C1"].Font.Bold = true;
            sheet1.Range["A2:C2"].Value = dt.ToShortDateString();

            for (int j = 0; j < dataGrid.Columns.Count; j++)
            {
                Range myRange = (Range)sheet1.Cells[4, j + 1];
                myRange.Font.Bold = true;
                string header = dataGrid.Columns[j].Header.ToString();
                sheet1.Columns[j + 1].ColumnWidth = header.Length + 5;
                myRange.Value2 = header;
            }
            for (int j = 0; j < dataGrid.Items.Count; j++)
            {
                var item = dataGrid.Items[j] as ControlVoice.Product;
                sheet1.Cells[j + 5, 1].Value = item.ProductId;
                sheet1.Cells[j + 5, 2].Value = item.ProductName;
                sheet1.Cells[j + 5, 3].Value = item.Price;
                sheet1.Cells[j + 5, 4].Value = item.Quantity;
                sheet1.Cells[j + 5, 5].Value = item.Status;
            }

            workbook.Save(); // Lưu lại Workbook đã mở
         
         
        }
    }
}
