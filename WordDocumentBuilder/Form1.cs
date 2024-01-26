using System;
using System.IO;
using System.Windows.Forms;
using WordDocumentBuilder.Models;
using Word = Microsoft.Office.Interop.Word;

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

            //write a method to return word doc


            /*                DocumentModel model = new DocumentModel("C:\\templates\\template1_4.dot");

                            model.SetBookmarkText("oBookMarkHeader", "Sky Walker");
                            model.AddStyledParagraph( "Heading 2Foo bar234242423");
                            model.AddStyledParagraph("?@@@@@This is a sentence of normal text. Now here is a table:", "boo", 24);
                            model.CreateTable( 13, 5);
                            model.InsertPageBreak();
                            model.InsertFileAtEnd(@"C:\Projects\C#\WordDocumentBuilder\word_documents\doc2.doc");
                            model.UpdateTableOfContents();
                            model.SaveDocument(@"C:\Projects\C#\WordDocumentBuilder\word_documents\newBar6.doc");

                      //  CreateWordDocument();


                        //Close this form.
                        this.Close();*/

            CreateAndSaveDocument();

            // Close the document.
            // doc1.Close();

            // Quit Word application.
            //wordApp.Quit();
        }

        public void CreateAndSaveDocument()
        {
            DocumentModel model = new DocumentModel("C:\\templates\\template1_4.dot");
/*
            model.SetBookmarkText("oBookMarkHeader", "Sky Walker");
            model.AddStyledParagraph("Heading 2Foo bar234242423");
            model.AddStyledParagraph("?@@@@@This is a sentence of normal text. Now here is a table:", "boo", 24);
            model.CreateTable(13, 5);
            model.InsertPageBreak();
            model.InsertFileAtEnd(@"C:\Projects\C#\WordDocumentBuilder\word_documents\doc2.doc");
            model.UpdateTableOfContents();*/

            model.AddStyledListParagraph("aaa");
           // model.SaveDocument(@"C:\Projects\C#\WordDocumentBuilder\word_documents\newBar6.doc");


      



            this.Close();
        }


    }

}
