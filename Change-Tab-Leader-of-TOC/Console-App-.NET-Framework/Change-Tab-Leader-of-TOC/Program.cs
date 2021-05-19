using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Change_Tab_Leader_of_TOC
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document
            WordDocument document = new WordDocument();
            //Adds the section into the Word document
            IWSection section = document.AddSection();
            string paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
            //Adds the paragraph into the created section
            IWParagraph paragraph = section.AddParagraph();
            //Appends the TOC field with LowerHeadingLevel and UpperHeadingLevel to determines the TOC entries
            paragraph.AppendTOC(1, 3);
            //Adds the section into the Word document
            section = document.AddSection();
            //Adds the paragraph into the created section
            paragraph = section.AddParagraph();
            //Adds the text for the headings
            paragraph.AppendText("First Chapter");
            //Sets a built-in heading style.
            paragraph.ApplyStyle(BuiltinStyle.Heading1);
            //Adds the text into the paragraph
            section.AddParagraph().AppendText(paraText);
            //Adds the section into the Word document
            section = document.AddSection();
            //Adds the paragraph into the created section
            paragraph = section.AddParagraph();
            //Adds the text for the headings
            paragraph.AppendText("Second Chapter");
            //Sets a built-in heading style.
            paragraph.ApplyStyle(BuiltinStyle.Heading2);
            //Adds the text into the paragraph
            section.AddParagraph().AppendText(paraText);
            //Adds the section into the Word document
            section = document.AddSection();
            //Adds the paragraph into the created section
            paragraph = section.AddParagraph();
            //Adds the text into the headings
            paragraph.AppendText("Third Chapter");
            //Sets a built-in heading style
            paragraph.ApplyStyle(BuiltinStyle.Heading3);
            //Adds the text into the paragraph.
            section.AddParagraph().AppendText(paraText);
            //Updates the table of contents
            document.UpdateTableOfContents();
            //Finds the TOC from Word document.
            TableOfContent toc = FindTableOfContent(document);

            //Change tab leader for table of contents in the Word document
            if (toc != null)
                ChangeTabLeaderForTableOfContents(toc, TabLeader.Hyphenated);

            //Saves and closes the Word document
            document.Save("Result.docx");
            document.Close();

            System.Diagnostics.Process.Start("Result.docx");
        }

        #region HelperMethods
        /// <summary>
        /// Change tab leader for table of contents in the Word document.
        /// </summary>
        /// <param name="toc"></param>
        private static void ChangeTabLeaderForTableOfContents(TableOfContent toc, TabLeader tabLeader)
        {
            //Inserts the bookmark start before the TOC instance.
            BookmarkStart bkmkStart = new BookmarkStart(toc.Document, "tableOfContent");
            toc.OwnerParagraph.Items.Insert(toc.OwnerParagraph.Items.IndexOf(toc), bkmkStart);

            Entity lastItem = FindLastTOCItem(toc);

            //Insert the bookmark end to next of TOC last item.
            BookmarkEnd bkmkEnd = new BookmarkEnd(toc.Document, "tableOfContent");
            WParagraph paragraph = lastItem.Owner as WParagraph;
            paragraph.Items.Insert(paragraph.Items.IndexOf(lastItem) + 1, bkmkEnd);

            //Creates the bookmark navigator instance to access the bookmark
            BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(toc.Document);
            //Moves the virtual cursor to the location before the end of the bookmark "tableOfContent"
            bookmarkNavigator.MoveToBookmark("tableOfContent");
            //Gets the bookmark content
            TextBodyPart part = bookmarkNavigator.GetBookmarkContent();

            //Iterates the items from the bookmark to change the spacing in the table of content
            for (int i = 0; i < part.BodyItems.Count; i++)
            {
                paragraph = part.BodyItems[i] as WParagraph;
                //Sets the tab leader
                if (paragraph.ParagraphFormat.Tabs.Count != 0)
                    paragraph.ParagraphFormat.Tabs[0].TabLeader = tabLeader;
            }

            //Remove the bookmark which we add to get the paragraphs in the table of contents
            Bookmark bookmark = toc.Document.Bookmarks.FindByName("tableOfContent");
            toc.Document.Bookmarks.Remove(bookmark);

        }


        /// <summary>
        /// Finds the last TOC item.
        /// </summary>
        /// <param name="toc"></param>
        /// <returns></returns>
        private static Entity FindLastTOCItem(TableOfContent toc)
        {
            int tocIndex = toc.OwnerParagraph.Items.IndexOf(toc);
            //TOC may contains nested fields and each fields has its owner field end mark 
            //so to indentify the TOC Field end mark (WFieldMark instance) used the stack.
            Stack<Entity> fieldStack = new Stack<Entity>();
            fieldStack.Push(toc);

            //Finds whether TOC end item is exist in same paragraph.
            for (int i = tocIndex + 1; i < toc.OwnerParagraph.Items.Count; i++)
            {
                Entity item = toc.OwnerParagraph.Items[i];

                if (item is WField)
                    fieldStack.Push(item);
                else if (item is WFieldMark && (item as WFieldMark).Type == FieldMarkType.FieldEnd)
                {
                    if (fieldStack.Count == 1)
                    {
                        fieldStack.Clear();
                        return item;
                    }
                    else
                        fieldStack.Pop();
                }

            }

            return FindLastItemInTextBody(toc, fieldStack);
        }
        /// <summary>
        /// Finds the last TOC item from consequence text body items.
        /// </summary>
        /// <param name="toc"></param>
        /// <param name="fieldStack"></param>
        /// <returns></returns>
        private static Entity FindLastItemInTextBody(TableOfContent toc, Stack<Entity> fieldStack)
        {
            WTextBody tBody = toc.OwnerParagraph.OwnerTextBody;

            for (int i = tBody.ChildEntities.IndexOf(toc.OwnerParagraph) + 1; i < tBody.ChildEntities.Count; i++)
            {
                WParagraph paragraph = tBody.ChildEntities[i] as WParagraph;

                foreach (Entity item in paragraph.Items)
                {
                    if (item is WField)
                        fieldStack.Push(item);
                    else if (item is WFieldMark && (item as WFieldMark).Type == FieldMarkType.FieldEnd)
                    {
                        if (fieldStack.Count == 1)
                        {
                            fieldStack.Clear();
                            return item;
                        }
                        else
                            fieldStack.Pop();
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Finds the table of content from Word document by iterating its elements.
        /// </summary>
        /// <param name="document">Word document</param>
        /// <returns></returns>
        private static TableOfContent FindTableOfContent(WordDocument document)
        {
            foreach (var item in document.Visit())
            {
                if (item is TableOfContent)
                    return item as TableOfContent;
            }
            return null;
        }
        #endregion

    }
    #region ExtendedClass
    /// <summary>
    /// DocIO extension class.
    /// </summary>
    public static class DocIOExtensions
    {
        public static IEnumerable<IEntity> Visit(this ICompositeEntity entity)
        {
            var entities = new Stack<IEntity>(new IEntity[] { entity });
            while (entities.Count > 0)
            {
                var e = entities.Pop();
                yield return e;
                if (e is ICompositeEntity)
                {
                    foreach (IEntity childEntity in ((ICompositeEntity)e).ChildEntities)
                    {
                        entities.Push(childEntity);
                    }
                }
            }
        }
    }
    #endregion
}
