namespace ExcelDemo.Unitity
{
    public class CellInfo
    {
        public CellInfo(string name, string header, int colIndex, string cellFormula, int width,
            bool isNumberColumn = false,bool isMerge=false)
        {
            PropertoryName = name;
            Header = header;
            IsNumberColumn = isNumberColumn;
            ColIndex = colIndex;
            Width = width;
            CellFormula = cellFormula;
            IsMerge = isMerge;
        }

        public string PropertoryName { get; set; }
        public string Header { get; set; }
        public bool IsNumberColumn { get; set; }
        public int ColIndex { get; set; }
        public string CellFormula { get; set; }

        public int Width { get; set; }

        public bool IsMerge { get; set; }

    }
}
