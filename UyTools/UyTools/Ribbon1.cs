using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Threading;

namespace UyTools
{
    public partial class Ribbon
    {
        public List<bool> checks = new List<bool>();
        public List<List<string>> fixied=new List<List<string>>();
        public string dic;
        public List<string> words=new List<string>();

        public string getdic()
        {
            String line;
            using (StreamReader sr = new StreamReader("ug.txt"))
            {
                line = sr.ReadToEnd();
            }
            return line;
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            dic = this.getdic();
        }
        TimeSpan ts;
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            DateTime dt = DateTime.Now;
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            int counts = doc.Words.Count;
            string txt = doc.Range(doc.Words.First.Start, doc.Words.Last.End).get_XML();
            StringBuilder sb = new StringBuilder();
            sb.Append(txt);
            doc.Range(doc.Words.First.Start, doc.Words.Last.End).InsertXML(Alkatip2Unicode(sb).ToString());
            doc.Content.Font.Name = "UKIJ Tuz";
            this.ts = DateTime.Now - dt;
            double ave = counts / Convert.ToDouble(this.ts.TotalSeconds);
            MessageBox.Show("Total Time Spent:" +  this.ts.TotalSeconds.ToString().Substring(0,4) + " seconds" + "\n" + "Processed Words：" + counts + "\n"  + "Words per Second：" +((int)ave).ToString()+" Words/S");

        }
        public StringBuilder Alkatip2Unicode(StringBuilder alkatipStr)
        {
            alkatipStr.Replace("الله", "ﷲ");
            alkatipStr.Replace("\u0626", "\u06D0");
            alkatipStr.Replace("\u0638", "\u0626");
            alkatipStr.Replace("\u0629", "\u06D5");
            alkatipStr.Replace("\u0635", "\u067E");
            alkatipStr.Replace("\u0622", "\u0698");
            alkatipStr.Replace("\u0636", "\u06AF");
            alkatipStr.Replace("\u062B", "\u06AD");
            alkatipStr.Replace("\u0647", "\u06BE");
            alkatipStr.Replace("\u0630", "\u06C7");
            alkatipStr.Replace("\u0623", "\u06C6");
            alkatipStr.Replace("\u0649", "\u06C8");
            alkatipStr.Replace("\u0624", "\u06CB");
            alkatipStr.Replace("\u0639", "\u0649");
            alkatipStr.Replace("\u062D", "\u0686");
            return alkatipStr;

        }
        
        public String Alkatip2Unicode(String alkatipStr)
        {
            const String AlkatipChars = "ظةصآضثهذأىؤئعح";
            const String UnicodeChars = "ئەپژگڭھۇۆۈۋېىچ";

            String unicodestr = "";

            foreach (char c in alkatipStr)
            {
                int idx = AlkatipChars.IndexOf(c);
                if (idx >= 0)
                {
                    unicodestr += UnicodeChars[idx];
                }
                else
                {
                    unicodestr += c.ToString();
                }
            }

            return unicodestr;
        }

        public List<string> stringtolist(string s)
        {
            List<string> words = new List<string>();
            string word = "";
            for (int i = 1; i < s.Length; i++)
            {
                if (char.IsLetterOrDigit(s[i - 1]))
                {
                    word += s[i - 1];
                }
                if ((s[i - 1]) == '\'') continue;
                if (char.IsWhiteSpace(s[i - 1]))
                {
                    if (word != "") words.Add(word);
                    word = "";
                }
                if (char.IsControl(s[i - 1]))
                {
                    if (word != "") words.Add(word);
                    word = "";
                    word += s[i - 1];
                    if (word != "") words.Add(word);
                    word = "";
                }
                if (char.IsControl(s[i - 1]) == true && s[i] == ' ')
                {
                    word += s[i];
                    if (word != "") words.Add(word);
                    word = "";

                }
                if ((char.IsPunctuation(s[i - 1]) == true || char.IsSymbol(s[i - 1]) == true) && char.IsPunctuation(s[i]) == false && char.IsSymbol(s[i]) == false)
                {
                    if (word != "") words.Add(word);
                    word = "";
                    word += s[i - 1];
                    if (word != "") words.Add(word);
                    word = "";
                }
            }
            return words;

        }


        public bool diccheck(string check)
        {
            if (dic.Contains(check)) return true;
            else return false;
        }

        private void fixeachword()
        {
            int i=0;
            foreach (bool check in checks)
            {
                if (check==false)
                {
                    fixied.Add(Globals.ThisAddIn.suggestions(words[i]));
                    //foreach(string sug in fixied[i])
                    //{
                    //    MessageBox.Show(words[i]+":"+sug);
                    //}
                }
                i++;
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            fixied.Clear();
            words.Clear();
            checks.Clear();
            DateTime dt = DateTime.Now;
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            Microsoft.Office.Interop.Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            int count = doc.Words.Count;
            string all = doc.Range(doc.Words.First.Start, doc.Words.Last.End).Text;
            words = stringtolist(all);
            words.Add("\r");
            int c = 0;
            
            if (words.Count == count)
            {
                foreach (string word in words)
                {
                    checks.Add(true);
                    // c++;
                    string each = word.Trim();
                    if (each.Contains("ا") || each.Contains("ى") || each.Contains("ە") || each.Contains("و") || each.Contains("ې") || each.Contains("ۇ") || each.Contains("ۆ") || each.Contains("ۈ"))
                    {
                        bool check = diccheck(each);
                        checks[c] = check;
                    }
                    c++;
                }
            }
            else
            {
                MessageBox.Show("It has some problem with getting words quickly, It will try wiht slowly way");
                foreach (Microsoft.Office.Interop.Word.Range word in document.Words)
                {
                    checks.Add(true);
                    string each = word.Text.Trim();
                    if (each.Contains("ا") || each.Contains("ى") || each.Contains("ە") || each.Contains("و") || each.Contains("ې") || each.Contains("ۇ") || each.Contains("ۆ") || each.Contains("ۈ"))
                    {
                        bool check = diccheck(each);
                        checks[c] = check;
                    }
                    c++;
                }

            }
            
            int cc = 0;
            foreach (Microsoft.Office.Interop.Word.Range word in document.Words)
            {
               // MessageBox.Show(hi[cc].ToString());
                if (checks[cc] == false)
               {
                   //MessageBox.Show(word.Text);
                   word.Font.Underline = Word.WdUnderline.wdUnderlineWavy;
                   word.Font.UnderlineColor = Word.WdColor.wdColorRed;
                  
               }
               cc++;
            }


            this.ts = DateTime.Now - dt;
            double ave = count / Convert.ToDouble(this.ts.TotalSeconds);
            MessageBox.Show("Total Time Spent:" + this.ts.TotalSeconds.ToString().Substring(0, 4) + " seconds" + "\n" + "Processed Words：" + count + "\n" + "Words per Second：" + ((int)ave).ToString() + " Words/S");
            Thread fixthread = new Thread(fixeachword);
            fixthread.Start();

        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 about = new Form1();
            about.ShowDialog();
        }



    }
}



 
