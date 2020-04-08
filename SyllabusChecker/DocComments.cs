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
        /// <summary>
        /// Добавление комментариев в рабочую программу формата .docx
        /// </summary>
        /// <param name="errorsTitle">Список индексов ошибок в титульном листе, с описаниями</param>
        /// <param name="errorsBody">Список индексов ошибок в остальном документе, с описаниями</param>
        /// <param name="shift">Количество параграфов в титульном листе (вызывает сдвиги в нумерации параграфов)</param>
        /// <param name="path">Путь к сохраняемому файлу с результатами</param>
        public static void AddComments(Dictionary<int, string> errorsTitle,
            Dictionary<int, string> errorsBody, int shift, string path)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(path, true))
            {
                foreach (KeyValuePair<int, string> err in errorsTitle)
                {
                    AddCommentOnParagraph(document, err.Key, err.Value);
                }

                foreach (KeyValuePair<int, string> err in errorsBody)
                {
                    AddCommentOnParagraph(document, err.Key + shift, err.Value);
                }
            }
        }

        /// <summary>
        /// Добавление комментариев в документ формата .docx
        /// </summary>
        /// <param name="errors">Список найденных в документе ошибок</param>
        /// <param name="path">Путь к сохраняемому файлу с результатами</param>
        public static void AddComments(Dictionary<int, string> errors, string path)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(path, true))
            {
                foreach (KeyValuePair<int, string> err in errors)
                {
                    AddCommentOnParagraph(document, err.Key, err.Value);
                }
            }
        }

        /// <summary>
        /// Добавление комментария к параграфу в документе
        /// </summary>
        /// <param name="document">WordprocessingDocument файла результатов</param>
        /// <param name="ind">Индекс параграфа в документе</param>
        /// <param name="comment">Комментарий к ошибке (текст комментария)</param>
        private static void AddCommentOnParagraph(WordprocessingDocument document, int ind, string comment)
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