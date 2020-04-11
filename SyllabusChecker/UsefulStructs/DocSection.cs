using System.Collections.Generic;
using Xceed.Document.NET;

namespace SyllabusChecker
{
    /// <summary>
    /// Для описания секции документа
    /// </summary>
    public struct DocSection
    {
        public int StartedAt;
        public int EndedAt;
        public List<Paragraph> Paragraphs;

        /// <summary>
        /// Секция документа
        /// </summary>
        /// <param name="startedAt">Индекс начала секции в исходном документе (параграфа с заголовком)</param>
        /// <param name="endedAt">Индекс конца секции в исходном документе (параграфа перед следующим заголовком)</param>
        /// <param name="paragraphs">Содержимое секции</param>
        /// <param name="isInSyllabus">Для секций рабочей программы - присутствует ли секция в документе</param>
        public DocSection(int startedAt, int endedAt, List<Paragraph> paragraphs)
        {
            StartedAt = startedAt;
            EndedAt = endedAt;
            Paragraphs = paragraphs;
        }
    }
}
