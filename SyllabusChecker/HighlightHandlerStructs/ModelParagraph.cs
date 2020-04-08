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
    };
}
