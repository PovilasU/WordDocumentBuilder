using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace WordDocumentBuilder
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            // Create a new Word application.
            // Word.Application wordApp = new Word.Application();
            // Word.Document doc1 = wordApp.Documents.Open(@"C:\Projects\C#\WordDocumentBuilder\word_documents\doc_new.doc");

            // Go to the end of the document.
            object oEndOfDoc = "\\endofdoc";
            object oMissing = System.Reflection.Missing.Value;


            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = true;

            object oTemplate = "C:\\Projects\\C#\\WordDocumentBuilder\\word_documents\\template1.dot";
            oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing,
            ref oMissing, ref oMissing);

            object oBookMark1 = "oBookMark1";
            oDoc.Bookmarks.get_Item(ref oBookMark1).Range.Text = "Some Text Here";

            //Insert a paragraph at the end of the document.
            Word.Paragraph oPara2;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara2.Range.Text = "Heading 2";
            oPara2.Format.SpaceAfter = 6;
            oPara2.Range.InsertParagraphAfter();


            Word.Range rng1 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            // Insert a page break at the end of the document.
            object breakType = Word.WdBreakType.wdPageBreak;
            rng1.InsertBreak(ref breakType);


            ///

            Word.Range rng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            // Insert the second document at the end of the first document.
            rng.InsertFile(@"C:\Projects\C#\WordDocumentBuilder\word_documents\doc2.doc", ref oMissing, ref oMissing, ref oMissing, ref oMissing);



         

            // Save the document.
            try
            {
                object fileName = @"C:\Projects\C#\WordDocumentBuilder\word_documents\newBar5.doc";
                oDoc.SaveAs(ref fileName);
                MessageBox.Show("File created successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }

            //Close this form.
             this.Close();

            // Close the document.
            // doc1.Close();

            // Quit Word application.
            //wordApp.Quit();
        }
    }
}
