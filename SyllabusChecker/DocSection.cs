using System.Collections.Generic;
using Xceed.Document.NET;

namespace SyllabusChecker
{
    public class DocSection
    {
        public int StartedAt;
        public int EndedAt;
        public List<Paragraph> Paragraphs;

        public DocSection(int startedAt, int endedAt, List<Paragraph> paragraphs)
        {
            StartedAt = startedAt;
            EndedAt = endedAt;
            Paragraphs = paragraphs;
        }
    }
}
