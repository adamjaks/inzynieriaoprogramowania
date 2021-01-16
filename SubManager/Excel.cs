﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace SubManager
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel(string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }
        public Excel()
        {

        }

        public string ReadCell(int i, int j)
        {
            if (ws != null && ws.Cells[i, j].Value2 != null)
            {
                return ws.Cells[i, j].Value2;
            }
            else
                return "";
        }
        public void WriteCell(int i, int j, string s)
        {
            ws.Cells[i, j].Value2 = s;

        }
        public void SaveFile()
        {
            wb.Save(); 
        }
    }
}
