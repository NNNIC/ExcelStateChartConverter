using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelStateChartConverter
{
    class LoadTemplateAndValues
    {
        public string    m_template_statefunc { get; private set; }
        public string    m_template_source    { get; private set; }
        public object[,] m_values             { get; private set; }

        public LoadTemplateAndValues(string file)
        {
            using (var ew = new ExcelWork())
            {
                ew.Load(file);
                //
                ew.SetSheet("template-source");
                m_template_source = ew.GetValue(0,0);
             
                ew.SetSheet("template-statefunc");
                m_template_statefunc = ew.GetValue(0,0);   

                ew.SetSheet("state-chart");
                m_values = (object[,])ew.GetValues().Clone();
            }
        }

        public int GetMaxRow() { return m_values.GetLength(0); }
        public int GetMaxCol() { return m_values.GetLength(1); }
        public string GetValue(int row,int col) // base 0
        {
            if (
                (row >= 0 && row < GetMaxRow()) 
                &&
                (col >=0 && col < GetMaxCol())
                )
            {
                try {
                    var v = m_values[row+1,col+1].ToString(); 
                    if (v!=null && (v.Length>0 && v[0]!='#'))
                    {
                        return v;
                    }
                }
                catch {
                    
                }
            }
            return "";
        }

        //出力ソース名を取得
        public string GetInitalSource(out string filename)
        {
            string mark = ":output=";
            filename = string.Empty;
            var output = string.Empty;
            foreach(var i in EditUtil.Split(m_template_source))
            {
                if (string.IsNullOrEmpty(i) || string.IsNullOrEmpty(i.TrimEnd())) continue;
                var l = i.TrimEnd();
                if (l.StartsWith(mark))
                {
                    filename = l.Substring(mark.Length);
                    continue;
                }
                if (l[0]==':') continue;

                output += l + "\n";
            }
            return output;
        }

        //プログラムランゲージ指定を取得
        public string GetInitalSource(out string filename, out string lang, out string enc, out string tempfunc)
        {
            string mark_output   = ":output=";
            string mark_lang     = ":lang=";
            string mark_enc      = ":enc=";
            string mark_tempfunc = ":templatefunc=";
            filename = string.Empty;
            lang     = string.Empty;
            enc      = string.Empty;
            tempfunc = string.Empty;
            var output = string.Empty;
            foreach(var i in EditUtil.Split(m_template_source))
            {
                if (string.IsNullOrEmpty(i) || string.IsNullOrEmpty(i.TrimEnd())) continue;
                var l = i.TrimEnd();
                if (l.StartsWith(mark_output))
                {
                    filename = l.Substring(mark_output.Length);
                    continue;
                }
                if (l.StartsWith(mark_lang))
                {
                    lang = l.Substring(mark_lang.Length);
                }
                if (l.StartsWith(mark_enc))
                {
                    enc =  l.Substring(mark_enc.Length);
                }
                if (l.StartsWith(mark_tempfunc))
                {
                    tempfunc =  l.Substring(mark_tempfunc.Length);
                }

                if (l[0]==':') continue;

                output += l + "\n";
            }
            return output;
        }


        //ファンクション用ファイル取得
        public string GetInitialFuncSource(string tempfunc=null)
        {
            if (string.IsNullOrEmpty(tempfunc)) tempfunc = m_template_statefunc;

            var output = string.Empty;
            foreach(var i in EditUtil.Split(tempfunc))
            {
                if (string.IsNullOrEmpty(i) || string.IsNullOrEmpty(i.TrimEnd())) continue;
                var l = i.TrimEnd();
                if (l[0]==':') continue;

                output += l + "\n";

            }
            return output;
        }
    }
}
