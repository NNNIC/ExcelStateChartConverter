using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExcelStateChartConverter
{
    class Program
    {
        const int NAME_COL  =1;
        const int START_COL =3;
        const int STATE_ROW = 1;

        static LoadTemplateAndValues m_ld;
        static string                m_enc = "utf-8";
        static string                m_overwride_tempfunc = string.Empty;
        static string                m_curpath;

        static void Main(string[] args)　// 0 --- excel file, 1 --- output dir
        {
            //System.Diagnostics.Debugger.Break();

            m_ld = new LoadTemplateAndValues(args[0]);
            m_curpath = Path.GetDirectoryName( Path.GetFullPath(args[0]) );

            string filename,lang,enc,tempfunc;
            var outpusrc = m_ld.GetInitalSource(out filename, out lang, out enc, out tempfunc);
            if (!string.IsNullOrEmpty(enc)) m_enc = enc;
            if (!string.IsNullOrEmpty(tempfunc))
            {
                m_overwride_tempfunc = File.ReadAllText(Path.Combine(m_curpath,tempfunc),Encoding.GetEncoding(enc));
            }
            

            //stateを収集
            var state_str = string.Empty;
            var state_list = new List<string>();
            {
                for(var c = START_COL; c<m_ld.GetMaxCol(); c++)
                {
                    var s = m_ld.GetValue(STATE_ROW,c);
                    if (string.IsNullOrEmpty(s) || string.IsNullOrEmpty(s.Trim())) continue;
                    state_list.Add(s.Trim());
                }

                //
                state_list.ForEach(s=> {
                    if (!string.IsNullOrEmpty(state_str)) state_str += ",";
                    state_str += s + "\n";
                });
            }

            // func_str作成
            var func_str = string.Empty;
            foreach(var st in state_list)
            {
                var tmpstr = m_ld.GetInitialFuncSource(m_overwride_tempfunc);
                tmpstr = convert_partsifexist_obs(st,tmpstr);
                tmpstr = convert_names(st,tmpstr);

                func_str +=  EditUtil.TrimLines(tmpstr);
            }

            //outputsrc作成
            outpusrc = EditUtil.Insert(outpusrc,"$contents1$",state_str);         
            outpusrc = EditUtil.Insert(outpusrc,"$contents2$",func_str);

            if (lang.ToLower()=="vba")
            {
                var select_str = string.Empty;
                foreach(var st in state_list)
                {
                    select_str += "Case \"" +st + "\"\n\t" + st + "(bFirst)\n";
                } 
                outpusrc = EditUtil.Insert(outpusrc, "$contents3$",select_str);
            }

            //
            outpusrc = "//This source is created by ExcelStateChartConverter.exe. Source : " + args[0] + "\n" + outpusrc;
           

            File.WriteAllText(Path.Combine(args[1],filename),outpusrc,Encoding.UTF8);

        }

        //ソース内の 存在確認後利用する部分を展開する
        private static string convert_partsifexist(string st, string tmpstr)
        {
            var output = string.Empty;
            var lines = tmpstr.Split('\n');
            for (var n = 0; n < lines.Length; n++)
            {
                var l = lines[n].TrimEnd();
                int index = 0;
                var e = EditUtil.Extract(l, out index);
                if (!string.IsNullOrWhiteSpace(e))
                {
                    var namecmt = e + "-cmt";
                    var indentspace = (new string(' ', index));
                    if (!string.IsNullOrWhiteSpace(_getvalue(st,namecmt))) //コメント行あり
                    {
                        output += indentspace + "/*" + "\n";
                        output += indentspace + string.Format("  [[{0}]]", namecmt) + "\n";
                        output += indentspace + "*/" + "\n";
                    }
                }
                output += l + "\n";
            }
            return output;
        }
        private static string convert_partsifexist_obs(string st,string tmpstr)
        {
            var mark1 = "<<<?";
            var mark2 = ">>>";
            var output = string.Empty;
            var lines = tmpstr.Split('\n');
            for(var n = 0; n<lines.Length; n++)
            {
                var l = lines[n].TrimEnd();
                var oindex = l.IndexOf("<<<?");
                if (oindex >=0)
                {
                    var name = l.Substring(oindex + mark1.Length).Trim();
                    var bExist  =  !string.IsNullOrEmpty( _getvalue(st,name));
                    
                    for(n++; n < lines.Length; n++)
                    {
                        var j = lines[n].TrimEnd();
                        if (j.Contains(mark2)) break;
                        if (bExist) {
                            output += j + "\n";
                        }
                    }
                }
                else
                {
                    output += l + "\n";
                }
            }
            return output;
        }

        //ソース内のname部分を値に変換する
        private static string convert_names(string st, string tmpstr)
        {
            var output = string.Empty;
            var lines = tmpstr.Split('\n');
            for(var n = 0; n<lines.Length; n++)
            {
                var l = lines[n].TrimEnd();
                var oindex = l.IndexOf("[[");
                var cindex = l.IndexOf("]]");
                if (oindex >= 0 && cindex >= 0 && oindex < cindex)
                {
                    var name = l.Substring(oindex+2,cindex - oindex -2);
                    var val = _getvalue(st,name).Trim();
                    
                    var result = EditUtil.Insert(l,"[[" + name + "]]",val).TrimEnd();
                    if (!string.IsNullOrWhiteSpace(result))
                    {
                         output += result + "\n";
                    }

                }
                else
                {
                    output += l + "\n";
                }
            }
            return output;
        }

        // ==== tools for this class ===
        static int _getColIndexByState(string st)
        {
            var row = _getRowIndexByName("state");
            for(var c = START_COL; c<m_ld.GetMaxCol();c++)
            {
                var s = m_ld.GetValue(row,c).Trim();
                if (s==st)
                {
                    return c;
                }
            }
            return -1;
        }

        static int _getRowIndexByName(string nm)
        {
            for(var i = 0; i<m_ld.GetMaxRow();i++)
            {
                var s = m_ld.GetValue(i,NAME_COL).Trim();
                if (s==nm)
                {
                    return i;
                }
            }
            return -1;
        }

        static string _getvalue(string state, string name)
        {
            var row = _getRowIndexByName(name);
            var col = _getColIndexByState(state);

            return m_ld.GetValue(row,col);
        }
 
        
    }
}
