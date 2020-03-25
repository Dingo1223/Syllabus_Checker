using System.Collections.Generic;
using Xceed.Document.NET;

namespace SyllabusChecker
{
    //Для описания секции документа
    public struct DocSection
    {
        //Индекс начала секции в исходном документе (параграфа с заголовком)
        public int StartedAt;

        //Индекс конца секции в исходном документе (параграфа перед следующим заголовком)
        public int EndedAt;

        //Содержимое секции
        public List<Paragraph> Paragraphs;

        public DocSection(int startedAt, int endedAt, List<Paragraph> paragraphs)
        {
            StartedAt = startedAt;
            EndedAt = endedAt;
            Paragraphs = paragraphs;
        }
    }
}
