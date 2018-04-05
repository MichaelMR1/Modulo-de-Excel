using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Drawing;

namespace Comun.Excel
{
    public class ExcelReport : ActionResult
    {
        private Stream excelStream;
        private String fileName;
        private bool saveAsXML;

        /// <summary>
        /// Creates a new ActionResult for saving excel files
        /// </summary>
        /// <param name="excel">byte array from excel workbook</param>
        /// <param name="fileName">string defining file name</param>
        public ExcelReport(DataTable excel, String fileName)
        {
            excelStream = new MemoryStream(DataTableToByte(excel));
            this.fileName = fileName;
            saveAsXML = false;
        }


        /// <summary>
        /// Creates a new ActionResult for saving excel files
        /// </summary>
        /// <param name="excel">byte array from excel workbook</param>
        /// <param name="fileName">string defining file name</param>
        public ExcelReport(byte[] excel, String fileName)
        {
            excelStream = new MemoryStream(excel);
            this.fileName = fileName;
            saveAsXML = false;
        }


        /// <summary>
        /// Creates a new ActionResult for saving excel files
        /// </summary>
        /// <param name="excel">byte array from excel workbook</param>
        /// <param name="fileName">string defining file name</param>
        /// <param name="saveAsXML">defines the content type as XML</param>
        public ExcelReport(byte[] excel, String fileName, bool saveAsXML)
        {
            excelStream = new MemoryStream(excel);
            this.fileName = fileName;
            this.saveAsXML = saveAsXML;
        }

        public override void ExecuteResult(ControllerContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            HttpResponseBase response = context.HttpContext.Response;

            response.ContentType = (saveAsXML) ? "text/xml" : "application/vnd.ms-excel";

            response.AddHeader("content-disposition", "attachment; filename=" + fileName);

            byte[] buffer = new byte[4096];

            while (true)
            {
                int read = this.excelStream.Read(buffer, 0, buffer.Length);
                if (read == 0)
                {
                    break;
                }

                response.OutputStream.Write(buffer, 0, read);
            }

            response.End();
        }


        public Byte[] DataTableToByte(DataTable dt)
        {
            using (ExcelPackage excelPkg = new ExcelPackage())
            {
                excelPkg.Workbook.Properties.Author = Assembly.GetExecutingAssembly()
                    .GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false)
                    .OfType<AssemblyDescriptionAttribute>().FirstOrDefault().Description;

                excelPkg.Workbook.Properties.Title = "Reporte " + System.DateTime.Now.ToLongDateString();

                ExcelWorksheet oSheet = CreateSheet(excelPkg, "Reporte");
                int rowIndex = 1;
                CreateHeader(oSheet, ref rowIndex, dt);
                CreateData(oSheet, ref rowIndex, dt);

                //TODO cambiar campos fechas en formato exportable


                //oSheet.Cells["A1:T" + rowIndex.ToString()].AutoFilter = true;
                oSheet.Cells[oSheet.Dimension.Address].AutoFitColumns();
                return excelPkg.GetAsByteArray();

            }

        }

        static public string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        private ExcelWorksheet CreateSheet(ExcelPackage excelPkg, string sheetName)
        {
            ExcelWorksheet oSheet = excelPkg.Workbook.Worksheets.Add(sheetName);
            // Setting default font for whole sheet
            oSheet.Cells.Style.Font.Name = "Calibri";
            // Setting font size for whole sheet
            oSheet.Cells.Style.Font.Size = 11;
            return oSheet;
        }

        private void CreateHeader(ExcelWorksheet oSheet, ref int rowIndex, DataTable dt)
        {
            int colIndex = 1;
            foreach (DataColumn dc in dt.Columns)
            {
                var cell = oSheet.Cells[rowIndex, colIndex];

                var fill = cell.Style.Fill;
                fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                var border = cell.Style.Border;
                border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                cell.Value = dc.ColumnName;

                colIndex++;
            }
        }

        private void CreateData(ExcelWorksheet oSheet, ref int rowIndex, DataTable dt)
        {
            int colIndex = 0;
            foreach (DataRow dr in dt.Rows)
            {
                colIndex = 1;
                rowIndex++;

                foreach (DataColumn dc in dt.Columns)
                {
                    var cell = oSheet.Cells[rowIndex, colIndex];

                    if (dc.DataType == typeof(DateTime))
                    {
                        if (dr.Field<DateTime?>(dc.ColumnName) != null)
                        {
                            cell.Style.Numberformat.Format = "dd/mm/yyyy";
                            cell.Formula = "=Date(" + dr.Field<DateTime>(dc.ColumnName).Year.ToString() + "," + dr.Field<DateTime>(dc.ColumnName).Month.ToString() + "," + dr.Field<DateTime>(dc.ColumnName).Day.ToString() + ")";
                        }
                        else
                        {
                            cell.Value = "";
                        }
                    }
                    else if (dc.DataType == typeof(int))
                    {
                        if (dr.Field<int?>(dc.ColumnName) != null)
                        {
                            cell.Style.Numberformat.Format = "0";
                            cell.Value = Convert.ToInt32(dr[dc.ColumnName]);
                        }
                        else
                        {
                            cell.Value = "";
                        }
                    }
                    else
                    {
                        cell.Value = dr[dc.ColumnName].ToString();
                    }

                    // Setting border of the cell
                    var border = cell.Style.Border;
                    border.Left.Style = border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    colIndex++;
                }
            }
        }
    }
}