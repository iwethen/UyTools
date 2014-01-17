using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using NHunspell;
using System.Reflection;
using Microsoft.Office.Core;
using System.Windows.Forms;

namespace UyTools
{
    public partial class ThisAddIn
    {
        _CommandBarButtonEvents_ClickEventHandler fixeventHandler;
        _CommandBarButtonEvents_ClickEventHandler moreeventHandler;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
             try
            {

                fixeventHandler = new _CommandBarButtonEvents_ClickEventHandler(Fix_Click);
                moreeventHandler = new _CommandBarButtonEvents_ClickEventHandler(More_Click);
                Word.Application applicationObject = Globals.ThisAddIn.Application as Word.Application;
                applicationObject.WindowBeforeRightClick += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(App_WindowBeforeRightClick);
 
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error: " + exception.Message);
            }
        }

        void App_WindowBeforeRightClick(Microsoft.Office.Interop.Word.Selection Sel, ref bool Cancel)
        {
            var doc = Globals .ThisAddIn .Application .ActiveDocument;
            try
            {
                
                
                
                this.RemoveItem();
                this.addmore();
                if (Globals.ThisAddIn.Application.Selection.Words.First.Font.Underline == Word.WdUnderline.wdUnderlineWavy)
                {
                    
                    List<string> seggrsted = this.suggestions(Globals.ThisAddIn.Application.Selection.Words.First.Text);

                    this.AddItem(seggrsted[0]);


                }
                
                
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error: " + exception.Message);
            }

        }
        private void addmore()
        {
            Word.Application applicationObject = Globals.ThisAddIn.Application as Word.Application;
            CommandBarButton addmore = applicationObject.CommandBars.FindControl(MsoControlType.msoControlButton, missing, "MORE_TAG", missing) as CommandBarButton;
            if (addmore != null)
            {
                System.Diagnostics.Debug.WriteLine("Found button, attaching handler");
                addmore.Click += moreeventHandler;
                return;
            }
            CommandBar popupCommandBar = applicationObject.CommandBars["Text"];
            bool isFound = false;
            foreach (object _object in popupCommandBar.Controls)
            {
                CommandBarButton _commandBarButton = _object as CommandBarButton;
                if (_commandBarButton == null) continue;
                if (_commandBarButton.Tag.Equals("MORE_TAG"))
                {
                    isFound = true;
                    System.Diagnostics.Debug.WriteLine("Found existing button. Will attach a handler.");
                    addmore.Click += moreeventHandler;
                    break;
                }
            }
            if (!isFound)
            {
                addmore = (CommandBarButton)popupCommandBar.Controls.Add(MsoControlType.msoControlButton, missing, missing, missing, true);
                System.Diagnostics.Debug.WriteLine("Created new button, adding handler");
                addmore.Click += moreeventHandler;
                addmore.Caption = "More";
                addmore.FaceId = 356;
                addmore.Tag = "MORE_TAG";
                addmore.BeginGroup = false;
            }
        }
        private void AddItem(string sug)
        {
            Word.Application applicationObject = Globals.ThisAddIn.Application as Word.Application;
            CommandBarButton commandBarButton = applicationObject.CommandBars.FindControl(MsoControlType.msoControlButton, missing, "HELLO_TAG", missing) as CommandBarButton;
            if (commandBarButton != null)
            {
                System.Diagnostics.Debug.WriteLine("Found button, attaching handler");
                commandBarButton.Click += fixeventHandler;
                return;
            }
            CommandBar popupCommandBar = applicationObject.CommandBars["Text"];
            bool isFound = false;
            foreach (object _object in popupCommandBar.Controls)
            {
                CommandBarButton _commandBarButton = _object as CommandBarButton;
                if (_commandBarButton == null) continue;
                if (_commandBarButton.Tag.Equals("HELLO_TAG"))
                {
                    isFound = true;
                    System.Diagnostics.Debug.WriteLine("Found existing button. Will attach a handler.");
                    commandBarButton.Click += fixeventHandler;
                    break;
                }
            }
            if (!isFound)
            {
                commandBarButton = (CommandBarButton)popupCommandBar.Controls.Add(MsoControlType.msoControlButton, missing, missing, missing, true);
                System.Diagnostics.Debug.WriteLine("Created new button, adding handler");
                commandBarButton.Click += fixeventHandler;
                commandBarButton.Caption = sug;
                commandBarButton.FaceId = 356;
                commandBarButton.Tag = "HELLO_TAG";
                commandBarButton.BeginGroup = true;
            }
        }

        private void RemoveItem()
        {
            Word.Application applicationObject = Globals.ThisAddIn.Application as Word.Application;
            CommandBar popupCommandBar = applicationObject.CommandBars["Text"];
            foreach (object _object in popupCommandBar.Controls)
            {
                CommandBarButton commandBarButton = _object as CommandBarButton;
                if (commandBarButton == null) continue;

                    popupCommandBar.Reset();
                
            }
        }
        private void Fix_Click(CommandBarButton cmdBarbutton, ref bool cancel)
        {
           // System.Windows.Forms.MessageBox.Show("Hello !!! Happy Programming", "Hello !!!");
            Globals.ThisAddIn.Application.Selection.Words.First.Font.Underline = Word.WdUnderline.wdUnderlineNone;
            Globals.ThisAddIn.Application.Selection.Words.First.Text = cmdBarbutton.Caption+" ";
            this.RemoveItem();

        }
        private void More_Click(CommandBarButton cmdBarbutton, ref bool cancel)
        {
           Fixing fixing=new Fixing();
           fixing.ShowDialog();
           this.RemoveItem();


        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Word.Application App = Globals.ThisAddIn.Application as Word.Application;
            App.WindowBeforeRightClick -= new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(App_WindowBeforeRightClick);

        }

        public bool check(string word)
        {
            bool correct;
            using (Hunspell hunspell = new Hunspell("ug.aff", "ug.dic"))
            {
                correct = hunspell.Spell(word);
            }
            return correct;
        }
        public bool add(string word)
        {
            bool added;
            using (Hunspell hunspell = new Hunspell("ug.aff", "ug.dic"))
            {
                added = hunspell.Add(word);
            }
            return added;
        }
        public List<string> suggestions(string word)
        {
            List<string> suggestions;
            using (Hunspell hunspell = new Hunspell("ug.aff", "ug.dic"))
            {
                suggestions = hunspell.Suggest(word);
            }
            return suggestions;
        }
        public bool hi ()
        {
            return true;
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
