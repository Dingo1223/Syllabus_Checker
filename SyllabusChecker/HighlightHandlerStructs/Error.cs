namespace SyllabusChecker
{
    public struct Error
    {
        //public string paragraph;
        public int Index;
        public string Type;

        public Error(int index, string type)
        {
            Index = index;
            Type = type;
        }
    }
}
