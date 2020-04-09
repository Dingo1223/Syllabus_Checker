namespace SyllabusChecker
{
    public struct Error
    {
        public int Index;
        public string Type;

        public Error(int index, string type)
        {
            Index = index;
            Type = type;
        }
    }
}
