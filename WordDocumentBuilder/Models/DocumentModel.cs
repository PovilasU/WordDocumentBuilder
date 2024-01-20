using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;



namespace WordDocumentBuilder.Models
{




    public class DocumentModel
    {
        public Word._Document Document { get; private set; }
        private object oMissing = System.Reflection.Missing.Value;
        private object oEndOfDoc = "\\endofdoc";

        public DocumentModel(string templatePath)
        {
            Word._Application oWord = new Word.Application();
            oWord.Visible = true;

            object oTemplate = templatePath;
            Document = oWord.Documents.Add(ref oTemplate, ref oMissing, ref oMissing, ref oMissing);
        }

        public void SetBookmarkText(string bookmark, string text)
        {
            if (!Document.Bookmarks.Exists(bookmark))
            {
                throw new ArgumentException("Bookmark does not exist in the document.", nameof(bookmark));
            }

            object oBookmark = bookmark;
            Document.Bookmarks.get_Item(ref oBookmark).Range.Text = text;
        }

        public Word.Table CreateTable(int numRows, int numColumns, float spaceAfter = 6, int bold = 1, int italic = 1)
        {

            object oEndOfDoc = "\\endofdoc";
            Word.Table oTable;
            object oRng = Document.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = Document.Tables.Add((Range)oRng, numRows, numColumns, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = spaceAfter;

            int r, c;
            string strText;
            for (r = 1; r <= numRows; r++)
                for (c = 1; c <= numColumns; c++)
                {
                    strText = "r" + r + "c" + c;
                    oTable.Cell(r, c).Range.Text = strText;
                }
            oTable.Rows[1].Range.Font.Bold = bold;
            oTable.Rows[1].Range.Font.Italic = italic;

            return oTable;
        }

        public Word.Paragraph AddStyledParagraph(string text, string styleName = "Heading 1", float spaceAfter = 6)
        {
            object oEndOfDoc = "\\endofdoc";
            Word.Paragraph oPara;
            object oRng = Document.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara = Document.Content.Paragraphs.Add(ref oRng);
            oPara.Range.Text = text;
            oPara.Format.SpaceAfter = spaceAfter;

            object oStyle = styleName;
            oPara.Range.set_Style(ref oStyle);
            oPara.Range.InsertParagraphAfter();

            return oPara;
        }

        public void InsertPageBreak()
        {
            Word.Range rng = Document.Bookmarks.get_Item(ref oEndOfDoc).Range;

            // Insert a page break at the end of the document.
            object breakType = Word.WdBreakType.wdPageBreak;
            rng.InsertBreak(ref breakType);
        }

        public void InsertFileAtEnd(string filePath)
        {
            Word.Range rng = Document.Bookmarks.get_Item(ref oEndOfDoc).Range;

            // Insert the file at the end of the document.
            rng.InsertFile(filePath, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
        }

        public void UpdateTableOfContents()
        {
            if (Document.TablesOfContents.Count > 0)
            {
                Word.TableOfContents toc = Document.TablesOfContents[1];
                toc.Update();
            }
        }


        public void SaveDocument(string fileName)
        {
            try
            {
                object oFileName = fileName;
                Document.SaveAs(ref oFileName);
                MessageBox.Show("File created successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
        }


    }

}


