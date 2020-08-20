using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;

namespace TOC_for_custom_styles
{
    class Program
    {
        static void Main(string[] args)
        {
            using (WordDocument document = new WordDocument())
            {
                document.EnsureMinimal();
                document.LastSection.PageSetup.Margins.All = 72;
                WParagraph para = document.LastParagraph;
                para.AppendText("Essential DocIO - Table of Contents");
                para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                para.ApplyStyle(BuiltinStyle.Heading4);
                para = document.LastSection.AddParagraph() as WParagraph;
                //Creates the new custom styles.
                WParagraphStyle pStyle1 = (WParagraphStyle)document.AddParagraphStyle("MyStyle1");
                pStyle1.CharacterFormat.FontSize = 18f;
                WParagraphStyle pStyle2 = (WParagraphStyle)document.AddParagraphStyle("MyStyle2");
                pStyle2.CharacterFormat.FontSize = 16f;
                WParagraphStyle pStyle3 = (WParagraphStyle)document.AddParagraphStyle("MyStyle3");
                pStyle3.CharacterFormat.FontSize = 14f;
                para = document.LastSection.AddParagraph() as WParagraph;
                //Insert TOC field in the Word document.
                TableOfContent toc = para.AppendTOC(1, 3);
                //Sets the Heading Styles to false, to define custom levels for TOC.
                toc.UseHeadingStyles = false;
                //Sets the TOC level style which determines; based on which the TOC should be created.
                toc.SetTOCLevelStyle(1, "MyStyle1");
                toc.SetTOCLevelStyle(2, "MyStyle2");
                toc.SetTOCLevelStyle(3, "MyStyle3");
                //Adds content to the Word document with custom styles.
                WSection section = document.LastSection;
                WParagraph newPara = section.AddParagraph() as WParagraph;
                newPara.AppendBreak(BreakType.PageBreak);
                AddHeading(section, "MyStyle1", "Document with custom styles", "This is the 1st custom style. This sample demonstrates the TOC insertion in a word document. Note that DocIO can insert TOC field in a word document. It can refresh or update TOC field by using UpdateTableOfContents method. MS Word refreshes the TOC field after insertion. Please update the field or press F9 key to refresh the TOC.");
                AddHeading(section, "MyStyle2", "Section 1", "This is the 2nd custom style. A document can contain any number of sections. Sections are used to apply same formatting for a group of paragraphs. You can insert sections by inserting section breaks.");
                AddHeading(section, "MyStyle3", "Paragraph 1", "This is the 3rd custom style. Each section contains any number of paragraphs. A paragraph is a set of statements that gives a meaning for the text.");
                AddHeading(section, "MyStyle3", "Paragraph 2", "This is the 3rd custom style. This demonstrates the paragraphs at the same level and style as that of the previous one. A paragraph can have any number formatting. This can be attained by formatting each text range in the paragraph.");
                AddHeading(section, "Normal", "Paragraph with normal", "This is the paragraph with normal style. This demonstrates the paragraph with outline level 4 and normal style. This can be attained by formatting outline level of the paragraph.");
                //Adds a new section to the Word document.
                section = document.AddSection() as WSection;
                section.PageSetup.Margins.All = 72;
                section.BreakCode = SectionBreakCode.NewPage;
                AddHeading(section, "MyStyle2", "Section 2", "This is the 2nd custom style. A document can contain any number of sections. Sections are used to apply same formatting for a group of paragraphs. You can insert sections by inserting section breaks.");
                AddHeading(section, "MyStyle3", "Paragraph 1", "This is the 3rd custom style. Each section contains any number of paragraphs. A paragraph is a set of statements that gives a meaning for the text.");
                AddHeading(section, "MyStyle3", "Paragraph 2", "This is the 3rd custom style. This demonstrates the paragraphs at the same level and style as that of the previous one. A paragraph can have any number formatting. This can be attained by formatting each text range in the paragraph.");
                //Updates the table of contents.
                document.UpdateTableOfContents();
                //Saves the file in the given path
                Stream docStream = File.Create(Path.GetFullPath(@"../../../TOC-custom-style.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }
        private static void AddHeading(WSection section, string styleName, string headingText, string paragraghText)
        {
            WParagraph newPara = section.AddParagraph() as WParagraph;
            WTextRange text = newPara.AppendText(headingText) as WTextRange;
            newPara.ApplyStyle(styleName);
            newPara = section.AddParagraph() as WParagraph;
            newPara.AppendText(paragraghText);
            section.AddParagraph();
        }
    }
}
