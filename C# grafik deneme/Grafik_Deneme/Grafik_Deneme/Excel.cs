using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Grafik_Deneme
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        public Excel()
        {
            ;
        }
        public Excel(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = excel.Worksheets[Sheet];
        }
        public int deneme()
        {
            _Excel.Range _Range = ws.UsedRange;
            int x = _Range.ReadingOrder;
            return x;
        }
        public Tuple<int, int> RowsAndColumns()
        {
            _Excel.Range _Range = ws.UsedRange;
            int nRows = _Range.Rows.Count;
            int nCols = _Range.Columns.Count;
            return new Tuple<int, int>(nRows, nCols);
        }
        public object ReadCell(int i, int j)
        {

            if (ws.Cells[i, j].Value2 != null)
            {
                object sendData = ws.Cells[i, j].Value2;
                return sendData;
            }
            else
            {
                return "";
            }
        }

        public void WritetoCell(int i, int j, string value)
        {
            i++; j++;
            ws.Cells[i, j].Value2 = value;
        }
        public void Save()
        {
            wb.Save();
        }
        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }
        public void Close()
        {
            wb.Close();
        }
        public void CreateNewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        }
    }
}