using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Word;
using Word = Microsoft.Office.Interop.Word;
namespace UyTools
{
    public partial class Fixing : Form
    {
        int c = 0;
        int q = 0;
        private Ribbon ribbon;
        public Fixing()
        {
            InitializeComponent();
        }

        private void Fixing_Load(object sender, EventArgs e)
        {
            c = 0;
            q = 0;
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            List<bool> me = Globals.Ribbons.Ribbon1.checks;
            int counts = me.Count;
            for (int i = c; i <= counts; i++)
            {
                c++;
                

                if (me[i] == false)
                {
                    label1.Text = doc.Words[c].Text;

                    try
                    {
                        listBox1.DataSource = Globals.Ribbons.Ribbon1.fixied[q];
                        
                    }
                    catch (Exception exception)
                    {
                        listBox1.DataSource = Globals.ThisAddIn.suggestions(doc.Words[i].Text.Trim());
                    }
                    q++;
                    break;
                }
                
            }
 

        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            
                          var doc = Globals.ThisAddIn.Application.ActiveDocument;
            List<bool> me = Globals.Ribbons.Ribbon1.checks;
            int counts = me.Count;
            if (c >= counts) this.Close();
            doc.Words[c].Text = listBox1.SelectedItem + " ";
            //MessageBox.Show(listBox1.SelectedItem);
            doc.Words[c].Font.Underline = Word.WdUnderline.wdUnderlineNone;
            for (int i = c; i < counts; i++)
            {
                c++;
                if (c >= counts) this.ParentForm.Close();
                if (me[i] == false)
                {
                    label1.Text = doc.Words[c].Text;
                    //MessageBox.Show(doc.Words[c + 1].Text);

                    try
                    {
                        listBox1.DataSource = Globals.Ribbons.Ribbon1.fixied[q];
                        

                    }
                    catch (Exception exception)
                    {
                        listBox1.DataSource = Globals.ThisAddIn.suggestions(doc.Words[i].Text.Trim());
                    }

                    q++;
                    break;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            List<bool> me = Globals.Ribbons.Ribbon1.checks;
            int counts = me.Count;
            if (c >= counts) this.Close();
            //doc.Words[c].Text = listBox1.SelectedItem + " ";
            //MessageBox.Show(listBox1.SelectedItem);
            doc.Words[c].Font.Underline = Word.WdUnderline.wdUnderlineNone;
            for (int i = c; i < counts; i++)
            {
                c++;
                if (c >= counts) this.ParentForm.Close();
                if (me[i] == false)
                {
                    label1.Text = doc.Words[c].Text;
                    //MessageBox.Show(doc.Words[c + 1].Text);

                    try
                    {
                        listBox1.DataSource = Globals.Ribbons.Ribbon1.fixied[q];


                    }
                    catch (Exception exception)
                    {
                        listBox1.DataSource = Globals.ThisAddIn.suggestions(doc.Words[i].Text.Trim());
                    }

                    q++;
                    break;
                }
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            List<bool> me = Globals.Ribbons.Ribbon1.checks;
            int counts = me.Count;
            if (c >= counts) this.Close();
            using (System.IO.StreamWriter file = new System.IO.StreamWriter("ug.txt", true,Encoding.UTF8))
            {
                file.WriteLine(doc.Words[c].Text.Trim());
            }
            Globals.Ribbons.Ribbon1.dic += doc.Words[c].Text.Trim();

            //doc.Words[c].Text = listBox1.SelectedItem + " ";
            //MessageBox.Show(listBox1.SelectedItem);
            doc.Words[c].Font.Underline = Word.WdUnderline.wdUnderlineNone;
            for (int i = c; i < counts; i++)
            {
                c++;
                if (c >= counts) this.ParentForm.Close();
                if (me[i] == false)
                {
                    label1.Text = doc.Words[c].Text;
                    //MessageBox.Show(doc.Words[c + 1].Text);

                    try
                    {
                        listBox1.DataSource = Globals.Ribbons.Ribbon1.fixied[q];


                    }
                    catch (Exception exception)
                    {
                        listBox1.DataSource = Globals.ThisAddIn.suggestions(doc.Words[i].Text.Trim());
                    }

                    q++;
                    break;
                }
            }

        

        }
    }
}
