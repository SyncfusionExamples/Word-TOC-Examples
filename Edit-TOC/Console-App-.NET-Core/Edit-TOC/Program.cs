using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System;
using System.IO;

namespace Edit_TOC
{
    class Program
    {
        static void Main(string[] args)
        {
            using (WordDocument document = new WordDocument())
            {
                //Opens the Word template document.
                Stream docStream = File.OpenRead(Path.GetFullPath(@"../../../TOC.docx"));
                document.Open(docStream, FormatType.Docx);
                docStream.Dispose();
                //Edits the TOC field.
                TableOfContent toc = document.Sections[0].Body.Paragraphs[2].Items[0] as TableOfContent;
                //Hides the page number in TOC.
                toc.IncludePageNumbers = false;
                //Includes the TC fields in TOC heading.
                toc.UseTableEntryFields = true;
                //Updates the table of contents
                document.UpdateTableOfContents();
                //Saves the file in the given path
                docStream = File.Create(Path.GetFullPath(@"../../../Sample.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }
    }
}
