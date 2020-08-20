using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;

namespace Update_TOC
{
    class Program
    {
        static void Main(string[] args)
        {
            using (WordDocument document = new WordDocument())
            {
                //Opens the Word template document
                Stream docStream = File.OpenRead(Path.GetFullPath(@"../../../TOC.docx"));
                document.Open(docStream, FormatType.Docx);
                docStream.Dispose();
                //Modifies the heading text and inserts a page break
                document.Replace("Section 1", "First section", true, true);
                document.Replace("Paragraph 1", "First paragraph", true, true);
                document.Replace("Paragraph 2", "Second paragraph", true, true);
                document.Replace("Section 2", "Second section", true, true);
                var selection = document.Find("heading 3 style", true, true);
                var paragraph = selection.GetAsOneRange().OwnerParagraph.NextSibling as WParagraph;
                paragraph.AppendBreak(BreakType.PageBreak);
                //Updates the table of contents
                document.UpdateTableOfContents();
                //Saves the file in the given path
                docStream = File.Create(Path.GetFullPath(@"../../../TOC-update.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }
    }
}
