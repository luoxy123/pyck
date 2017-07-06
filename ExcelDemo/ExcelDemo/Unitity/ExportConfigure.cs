using System.Collections.Generic;

namespace ExcelDemo.Unitity
{
    public class ExportConfigure
    {
        public ExportConfigure(string title, string sheetName)
        {
            Title = title;
            SheetName = sheetName;
            Cells = new List<CellInfo>();
        }

        /// <summary>
        ///     标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        ///     sheet名称
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        ///     导入模式（treegrid,single,muti)
        /// </summary>
        public string Mode { get; set; }

        /// <summary>
        ///     关系字段，仅当Mode=treegrid时，组合行列时使用
        /// </summary>
        public string Relationship { get; set; }

        /// <summary>
        ///     行高
        /// </summary>
        public short RowHeight { get; set; }

        public IList<CellInfo> Cells { get; set; }

        public bool IsMerge { get; set; }

    }
}
