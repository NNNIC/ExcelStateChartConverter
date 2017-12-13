using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelStateChartConverter
{
    public class ExcelWork : IDisposable
    {
        public void Dispose() //ref http://ufcpp.net/study/csharp/rm_disposable.html
        {
            
            GC.SuppressFinalize(this);

            if (m_sheetCtr!=null)
            {
                m_sheetCtr.Dispose();
                m_sheetCtr = null;
            }
            if (m_book!=null)
            {
                m_book.Dispose();
                m_book = null;
            }
        }


        ExcelUtil.BookCtr  m_book;
        ExcelUtil.SheetCtr m_sheetCtr;
        object[,]          m_values;

        public string m_file;

        public void Load(string file)
        {
            m_book = ExcelUtil.OpenBook(file);
        }

        public void SetSheet(string name)
        {
            m_sheetCtr = m_book.GetSheet(name);
            m_values   = m_sheetCtr.GetValues();
        }

        public int GetRowLenght()
        {
            return m_values.GetLength(0);
        }
        public int GetColLength()
        {
            return m_values.GetLength(1);
        }

        public string GetValue(int row, int col) //base 0
        {
            var r = row+1;
            var c = col+1;
            if (
                (r > 0 && r <= m_values.GetLength(0))
                &&
                (c > 0 && c <= m_values.GetLength(1))
                )
            {
                return m_values[r,c].ToString();
            }
            return null;
        }

        public object[,] GetValues()
        {
            return m_values;
        }
    }
}
