using System;
using System.IO;

namespace Comun.Excel.Service
{
    internal class ExcelPackage : IDisposable
    {
        private Stream st;

        public ExcelPackage(Stream st)
        {
            this.st = st;
        }

        public object Workbook { get; internal set; }
    }
}