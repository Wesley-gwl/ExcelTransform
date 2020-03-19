using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTransform
{
    public class TxtHelper
    {
        public static string Read(string path)
        {
            var txt = string.Empty;
            using (StreamReader sr = new StreamReader(path, Encoding.Default))
            {
                int lineCount = 0;
                while (sr.Peek() > 0)
                {
                    lineCount++;
                    string temp = sr.ReadLine();
                    txt += temp;
                }
            }
            return txt;
        }
    }
}