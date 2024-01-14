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

          //  object oTemplate = "C:\\Projects\\C#\\WordDocumentBuilder\\word_documents\\template1_2.dot";
            object oTemplate = "C:\\templates\\template1_4.dot";
            // object oTemplate = "C:\\Projects\\C#\\WordDocumentBuilder\\word_documents\\Title.dot";
            oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing,
            ref oMissing, ref oMissing);

            //Create Header
            object oBookMark1 = "oBookMark1"; 
            oDoc.Bookmarks.get_Item(ref oBookMark1).Range.Text = "Some Text Here";
            
            object oBookMarkHeader = "oBookMarkHeader";
            oDoc.Bookmarks.get_Item(ref oBookMarkHeader).Range.Text = "Sky Walker";

            //Insert a paragraph at the end of the document.
            Word.Paragraph oPara2;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara2.Range.Text = "Heading 2Foo bar";
            oPara2.Format.SpaceAfter = 6;

            object oStyleName = "Heading 1";
            oPara2.Range.set_Style(ref oStyleName);
            oPara2.Range.InsertParagraphAfter();


            //Insert another paragraph.
            Word.Paragraph oPara3;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara3.Range.Text = "This is a sentence of normal text. Now here is a table:";
            oPara3.Range.Font.Bold = 0;
            oPara3.Format.SpaceAfter = 24;
            object oStyleName1 = "boo";
            oPara3.Range.set_Style(ref oStyleName1);
            oPara3.Range.InsertParagraphAfter();

            //Insert a 3 x 5 table, fill it with data, and make the first row
            //bold and italic.
            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 3, 5, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            int r, c;
            string strText;
            for (r = 1; r <= 3; r++)
                for (c = 1; c <= 5; c++)
                {
                    strText = "r" + r + "c" + c;
                    oTable.Cell(r, c).Range.Text = strText;
                }
            oTable.Rows[1].Range.Font.Bold = 1;
            oTable.Rows[1].Range.Font.Italic = 1;





            //Inserting last Page
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
                object fileName = @"C:\Projects\C#\WordDocumentBuilder\word_documents\newBar6.doc";
                //  oDoc.SaveAs(ref fileName);
                MessageBox.Show("File created successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }


            // Update the table of contents.
            if (oDoc.TablesOfContents.Count > 0)
            {
                Word.TableOfContents toc = oDoc.TablesOfContents[1];
                toc.Update();
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
