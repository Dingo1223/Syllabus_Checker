using Xceed.Document.NET;

namespace SyllabusChecker
{
    public struct ModelParagraph
    {
        public int SourceIndex; //index in original
        public string Paragraph;

        public ModelParagraph(int sourceIndex, string paragraph)
        {
            SourceIndex = sourceIndex;
            Paragraph = paragraph;
        }
        //public string text; //full text
        //public ElementaryFragment[] fragments; //magic text
    };
}
