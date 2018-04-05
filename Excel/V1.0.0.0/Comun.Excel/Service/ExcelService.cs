using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Comun.Excel.Service
{
    public class ExcelService
    {

        public DataTable GetTablaFormXLSX(Stream st)
        {

            DataTable dt = new DataTable();

            using (var excel = new ExcelPackage(st))
            {

                var ws = excel.Workbook.Worksheets.First();
                var hasHeader = true;  //ajustar cabecera

                // añade DataColumns al DataTable
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    dt.Columns.Add(hasHeader ? firstRowCell.Text
                        : String.Format("Column {0}", firstRowCell.Start.Column));

                // agrega DataRows al DataTable
                int startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = dt.NewRow();
                    foreach (var cell in wsRow)
                        row[cell.Start.Column - 1] = cell.Text;
                    dt.Rows.Add(row);
                }

            }

            return dt;
        }
    }
}
