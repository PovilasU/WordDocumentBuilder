using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Collections.Generic;


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
         //   oWord.Visible = true;
            oWord.Visible = false;

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


        public List<string> ProcessString(ref string str, ref string[] strings)
        {
            List<string> stringsList = new List<string>();
            Match match = Regex.Match(str, @"\n(?<charAfterNewLine>.)");

            if (match.Success)
            {
                char charAfterNewLine = match.Groups["charAfterNewLine"].Value[0];
                if (char.IsLower(charAfterNewLine))
                {
                  //  Console.WriteLine("The character after '\\n' is lower case.");
                    str = str.Replace("\n", " ");
                 //   Debug.WriteLine($"Replaced '\\n' with ' ' in s1. New s1: '{str}'");
                    stringsList.Add(str);
                }
                else if (char.IsUpper(charAfterNewLine))
                {
                   // Console.WriteLine("The character after '\\n' is upper case.");
                    MatchCollection matches = Regex.Matches(str, @"\n(?<charAfterNewLine>[A-Z].*?)");

                    foreach (Match match1 in matches)
                    {
                        if (char.IsUpper(charAfterNewLine))
                        {
                            string substringBeforeNewLine = str.Substring(0, match1.Index);
                            Array.Resize(ref strings, strings.Length + 1); // Resize the array
                            strings[strings.Length - 1] = substringBeforeNewLine; // Add the substring to the array
                            stringsList.Add(substringBeforeNewLine);

                            if (match1.Index + 1 < str.Length)
                            {
                                str = str.Substring(match1.Index + 1); // Update str to be the remaining substring
                            }
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("No '\\n' found in the string.");
            }

            return stringsList;
        }


        public Word.Paragraph AddStyledListParagraph(string text, string styleName = "Normal", float spaceAfter = 6)
        {
           // string[] strings;

            string s1 = "Upgrade to a bla bla operating\nsystem or whatever";
         //   s1 = Regex.Replace(s1, @"\n(?=[a-z])", " ");

           //  strings = new string[] { s1 };

            string s2 = "Please1 update to OS 9.3.4 or above\nPlease2 do what I say 4.5.6 or above\nPlease3 update to OS 8.9.9 or above";


            //string s1 = "Upgrade to a bla bla operating\nsystem or whatever";
            string[] strings = new string[0]; // Initialize the array

            string str = s2;
            //List<string> stringsList = new List<string>();
            List<string> stringsList = ProcessString(ref str, ref strings);

            /*
                        Match match = Regex.Match(str, @"\n(?<charAfterNewLine>.)");

                             if (match.Success)
                             {
                                 char charAfterNewLine = match.Groups["charAfterNewLine"].Value[0];
                                 if (char.IsLower(charAfterNewLine))
                                 {
                                     Console.WriteLine("The character after '\\n' is lower case.");
                                     str = str.Replace("\n", " ");
                                     Debug.WriteLine($"Replaced '\\n' with ' ' in s1. New s1: '{str}'");
                                stringsList.Add(str);
                            }
                                 else if (char.IsUpper(charAfterNewLine))
                                 {
                                     Console.WriteLine("The character after '\\n' is upper case.");
                                MatchCollection matches = Regex.Matches(str, @"\n(?<charAfterNewLine>[A-Z].*?)");

                                int count = 0;

                                foreach (Match match1 in matches)
                                {
                                   // char charAfterNewLine = match1.Groups["charAfterNewLine"].Value[0];
                                    if (char.IsUpper(charAfterNewLine))
                                    {
                                        string substringBeforeNewLine = str.Substring(0, match1.Index);
                                        Array.Resize(ref strings, strings.Length + 1); // Resize the array
                                        strings[strings.Length - 1] = substringBeforeNewLine; // Add the substring to the array
                                                                                              //Debug.WriteLine($"Added '{substringBeforeNewLine}' to the array.");
                                                                                              //stringsList
                                        stringsList.Add(substringBeforeNewLine);
                                        count++;
                                        //  str = str.Substring(match.Index + 1); // Update str to be the remaining substring
                                        if (match1.Index + 1 < str.Length)
                                        {
                                            str = str.Substring(match1.Index + 1); // Update str to be the remaining substring
                                        }

                                    }
                                    else
                                    {
                                        //Debug.WriteLine("what is here");
                                        // Debug.WriteLine(str.Substring(0, match.Index));
                                    }
                                }
                            }
                             }
                             else
                             {
                                 Console.WriteLine("No '\\n' found in the string.");
                             }*/



            


            foreach (string item in stringsList)
            {
                 Debug.WriteLine(item); //does the job
            }


            //using System.Text.RegularExpressions;

            //string s1 = "Upgrade to a bla bla operating\nsystem or whatever";
            /*            string[] strings = new string[0]; // Initialize the array

                        string str = s1;
                        List<string> stringsList = new List<string>();*/


            //string str = s2;

            // Match all occurrences of substrings starting with a capital letter after "\n"
            /*        MatchCollection matches = Regex.Matches(str, @"\n(?<charAfterNewLine>[A-Z].*?)");

                    int count = 0;

                    foreach (Match match1 in matches)
                    {
                        char charAfterNewLine = match1.Groups["charAfterNewLine"].Value[0];
                        if (char.IsUpper(charAfterNewLine))
                        {
                            string substringBeforeNewLine = str.Substring(0, match1.Index);
                            Array.Resize(ref strings, strings.Length + 1); // Resize the array
                            strings[strings.Length - 1] = substringBeforeNewLine; // Add the substring to the array
                                                                                  //Debug.WriteLine($"Added '{substringBeforeNewLine}' to the array.");
                                                                                  //stringsList
                            stringsList.Add(substringBeforeNewLine);
                            count++;
                            //  str = str.Substring(match.Index + 1); // Update str to be the remaining substring
                            if (match1.Index + 1 < str.Length)
                            {
                                str = str.Substring(match1.Index + 1); // Update str to be the remaining substring
                            }

                        }
                        else
                        {
                            //Debug.WriteLine("what is here");
                           // Debug.WriteLine(str.Substring(0, match.Index));
                        }
                    }*/

            foreach (string item in stringsList)
            {
               // Debug.WriteLine(item); //does the job
            }


            // Add the remaining substring to the array
            Array.Resize(ref strings, strings.Length + 1);
         //   strings[strings.Length - 1] = str;
           // Debug.WriteLine($"Added '{str}' to the array.");

            //Debug.WriteLine($"Number of substrings starting with a capital letter after '\\n': {count}");
            foreach (string s in strings)
            {
               // Debug.WriteLine($"Array element: '{s}'");
            }




            /*            Match match = Regex.Match(str, @"(?<substringBeforeNewLine>.*?)\n(?<charAfterNewLine>.)");
                        int count = 0;
                        if (match.Success)
                        {
                            char charAfterNewLine = match.Groups["charAfterNewLine"].Value[0];
                            if (char.IsUpper(charAfterNewLine))
                            {
                                string substringBeforeNewLine = match.Groups["substringBeforeNewLine"].Value;
                                Array.Resize(ref strings, strings.Length + 1); // Resize the array
                                strings[strings.Length - 1] = substringBeforeNewLine; // Add the substring to the array
                                Debug.WriteLine($"Added '{substringBeforeNewLine}' to the array.");

                                count++;
                            }
                            else
                            {
                                str = str.Replace("\n", " ");
                                Debug.WriteLine($"Replaced '\\n' with ' ' in s1. New s1: '{str}'");
                            }
                        }
                        else
                        {
                            Debug.WriteLine("No '\\n' found in the string.");
                        }*/

            //  Debug.WriteLine($"new s1@@: '{str}'");
            //Debug.WriteLine($"new strings@@: '{strings}'");
            Debug.WriteLine($"new strings Length@@: '{strings.Length}'");









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


