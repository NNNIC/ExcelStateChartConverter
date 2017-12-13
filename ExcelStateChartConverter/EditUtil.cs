using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelStateChartConverter
{
    class EditUtil
    {
        internal static string Insert(string src, string id, string state_str)
        {
            var output = string.Empty;
            foreach(var i in  Split(src))
            {
                var s = i.TrimEnd();
                if (s.Contains(id))
                {
                    var numofsp = s.IndexOf(id);

                    var state_lines = Split(state_str);
                    if (state_lines.Length <= 1)
                    {
                        output += s.Replace(id,state_str.Trim()) + "\n";
                    }
                    else
                    {
                        var pretext = s.Substring(0,numofsp);
                        if (!string.IsNullOrWhiteSpace(pretext))
                        { 
                            output += pretext + "\n";
                        }
                        foreach(var j in state_str.Split('\n'))
                        {
                            output += (new string(' ',numofsp)) + j.TrimEnd() + "\n";
                        }
                        var posttext = s.Substring(numofsp + id.Length);
                        if (!string.IsNullOrWhiteSpace(posttext))
                        {
                            output += posttext + "\n";
                        }
                    }
                }
                else
                {
                    output += s.TrimEnd() + "\n";
                }
            }
            return output;
        }


        public static string[] Split(string s)
        {
            var lines = new List<string>(s.Split('\n'));
            var n = lines.Count;
            for(var loop = 0; loop < n ; loop++)
            {
                if (lines.Count==0) break;
                if (string.IsNullOrWhiteSpace((lines[lines.Count-1].Trim())))
                {
                    lines.RemoveAt(lines.Count-1);
                }
                else
                {
                    break;
                }
            }
            return lines.ToArray();
        }

        public static string TrimLines(string s)
        {
            var output = string.Empty;
            Array.ForEach(Split(s),i=>output += i.TrimEnd() + "\n");
            return output;
        }

        public static string Extract(string s, out int index)
        {
            index = -1;
            var si = s.IndexOf("[[");
            if (si<0) return string.Empty;

            var ei = s.IndexOf("]]",si);
            if (ei<0) return string.Empty;

            for(var i = 0; i<s.Length; i++)
            {
                if (s[i] > ' ') { index = i; break;}
            }

            return s.Substring(si+2, ei-(si+2));
        }

    }
}
