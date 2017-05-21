namespace MMA
{
    public static class Tools
    {
        public static string StringTrim(string name)
        {
            int i = name.IndexOf(":");
            if (i > -1)
                return name.Substring(0, i);
            return name;
        }
    }

    public class Matrix
    {
        public object[,] content;
        public bool hasHeader() { return false; }
        public bool hasRowHeader() { return false; }
        public int contentWidth() { return content.GetLength(1); }
        public int contentHeight() { return content.GetLength(0); }
    }
    public class MatrixH : Matrix
    {
        public object[] columnHeaders;
        public new bool hasHeader() { return true; }
    }
    public class MatrixHR : MatrixH
    {
        public object[] rowHeaders;
        public string upperLeft;
        public new bool hasRowHeader() { return true; }
    }
}