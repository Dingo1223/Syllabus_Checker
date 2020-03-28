using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SyllabusChecker
{
    public static class DocComments
    {
        public static void AddComments(List<int> indsTitle, List<int> indsBody, int shift, string path)
        {
            foreach (int ind in indsTitle)
            {
                AddCommentOnParagraph(path, ind, "ошибка в титульнике " + ind.ToString());
            }

            foreach (int ind in indsBody)
            {
                AddCommentOnParagraph(path, ind + shift, "ошибка в теле " + ind.ToString());
            }
        }

        private static void AddCommentOnParagraph(string fileName, int ind, string comment)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, true))
            {
                Paragraph paragraph = document.MainDocumentPart.Document.Descendants<Paragraph>().ElementAt(ind);
                Comments comments = null;
                int id = 0;

                if (document.MainDocumentPart.GetPartsOfType<WordprocessingCommentsPart>().Count() > 0)
                {
                    comments = document.MainDocumentPart.WordprocessingCommentsPart.Comments;
                    if (comments.HasChildren)
                    {
                        id = int.Parse(comments.Descendants<Comment>().Select(e => e.Id.Value).Max()) + 1;
                    }
                }
                else
                {
                    WordprocessingCommentsPart commentPart = document.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                    commentPart.Comments = new Comments();
                    comments = commentPart.Comments;
                }

                Paragraph p = new Paragraph(new Run(new Text(comment)));
                Comment cmt = new Comment()
                {
                    Id = id.ToString(),
                    Author = new StringValue("Программа проверки"),
                    Initials = new StringValue("Программа проверки"),
                    Date = DateTime.Now
                };
                cmt.AppendChild(p);
                comments.AppendChild(cmt);
                comments.Save();

                paragraph.InsertBefore(new CommentRangeStart() { Id = id.ToString() }, paragraph.GetFirstChild<Run>());

                CommentRangeEnd cmtEnd = paragraph.InsertAfter(new CommentRangeEnd() { Id = id.ToString() }, paragraph.Elements().Last());

                paragraph.InsertAfter(new Run(new CommentReference() { Id = id.ToString() }), cmtEnd);
            }
        }
    }
}