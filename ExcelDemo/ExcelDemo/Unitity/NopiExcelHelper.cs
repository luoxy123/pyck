using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Npoi.Core.HPSF;
using Npoi.Core.HSSF.UserModel;
using Npoi.Core.SS.UserModel;
using Npoi.Core.SS.Util;

namespace ExcelDemo.Unitity
{
    public class NopiExcelHelper<TSource> where TSource:class,new()
    {
        /// <summary>
        ///     sheet标题行索引
        /// </summary>
        private const int SheetTitleRowIndex = 0;

        /// <summary>
        ///     sheet中数据内容的表头行索引
        /// </summary>
        private const int SheetDataHeaderRowIndex = 1;


        private readonly HSSFWorkbook _workbook = new HSSFWorkbook();
        private ExportConfigure _config;

        private string title;
        private string sheetname;

        public string Title
        {
            get { return title; }
            set { title = value; }

        }
        public string SheetName
        {
            get { return sheetname; }
            set { sheetname = value; }
        }

        public NopiExcelHelper(string excelTitle, string sheetName)
        {
            title = excelTitle;
            sheetname = sheetName;
        }
        private int _numberColumn = 1;
        public void ExportToFile(string configFilePath, IEnumerable<TSource> source, string savePath = "")
        {
            Init(configFilePath);
            ToExport(source);
            var fileSavePath = string.IsNullOrWhiteSpace(savePath) ? @"d:\\test.xls" : savePath;
            WriteToFile(fileSavePath);
        }
        public MemoryStream ExportToMemoryStream(string configFilePath, IEnumerable<TSource> source)
        {
            Init(configFilePath);
            ToExport(source);

            var memoryStreamFile = new MemoryStream();
            _workbook.Write(memoryStreamFile);

            memoryStreamFile.Flush();
            memoryStreamFile.Position = 0;

            return memoryStreamFile;
        }

        private void ToExport(IEnumerable<TSource> source)
        {
            var sheetCount = 0;
            var newsheet = CreateSheet(_config, sheetCount);
            var rowIndex = 2; //内容默认起始行-第三行，第一行是sheet标题,第二行是数据表头
            foreach (var dr in source)
            {
                //超出5000条数据 创建新的工作簿
                if (rowIndex == 5000)
                {
                    sheetCount++;
                    newsheet = CreateSheet(_config, sheetCount);
                    rowIndex = 2;
                }
                CreateNewRow(_workbook, newsheet, rowIndex, dr, _config);
                rowIndex++;
            }
            var cellStyle = _workbook.CreateCellStyle();
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.Alignment = HorizontalAlignment.Center;
            for (int i = 0; i < _config.Cells.Count; i++)
            {
                if (i >= 3)
                {
                    newsheet.AddMergedRegion(new CellRangeAddress(2, newsheet.LastRowNum, i, i));
                    newsheet.SetDefaultColumnStyle(i, cellStyle);
                }
               
            }

        }


        private void CreateNewRow(IWorkbook workbook, ISheet newsheet, int rowIndex, TSource dr, ExportConfigure config)
        {
            var newRow = newsheet.CreateRow(rowIndex);
            if (config.Mode.ToLower().Equals("treegrid"))
            {
                var relationship = config.Relationship;
                var val = GetFieldValue(dr, relationship);
                var relationshipVal = val?.ToString();
                //TODO:还未实现TreeGrid，及更复杂的表现形式
                InsertCell(dr, newRow, newsheet, workbook, config);
            }
            else
            {
                InsertCell(dr, newRow, newsheet, workbook, config);
            }
        }

        private void InsertCell(TSource dr, IRow newRow, ISheet newsheet, IWorkbook workBook, ExportConfigure config)
        {
            var cellIndex = 0;
            var style = CreateHasBroderCell(workBook);
            //循环导出列
            foreach (var cfg in config.Cells)
            {
                var newCell = CreateCell(newRow, cellIndex, cfg);
                if (cfg.IsNumberColumn)
                {
                    newCell.SetCellValue(_numberColumn);
                    _numberColumn++;
                }
                else
                {
                    SetCellValue(dr, cfg, cellIndex, newsheet, newCell);
                }
                newCell.CellStyle = style;
                cellIndex++;
            }
        }
        private ICell CreateCell(IRow newRow, int cellIndex, CellInfo cfg)
        {
            var newCell = newRow.CreateCell(cellIndex);
            //TODO:根据cfg配置构造cell格式
            return newCell;
        }

        /// <summary>
        ///     创建有边框的CellStyle
        /// </summary>
        /// <param name="workbook"></param>
        private static ICellStyle CreateHasBroderCell(IWorkbook workbook)
        {
            var style = workbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            return style;
        }
        private void SetCellValue(TSource dr, CellInfo cellConfig, int cellIndex, ISheet sheet, ICell excelCol)
        {
            if (cellConfig.Width != 0)
            {
                sheet.SetColumnWidth(cellIndex, cellConfig.Width * 256);
            }

            var t = dr.GetType();
            var propertities = t.GetProperties();
            var pro = propertities.FirstOrDefault(it => it.Name.Equals(cellConfig.PropertoryName));

            if (pro == null) return;
            var val = pro.GetValue(dr, null);
            SetCellValue(excelCol, pro.PropertyType, val);
        }

        private void SetCellValue(ICell newCell, Type proType, object val)
        {
            while (true)
            {
                var strVal = val?.ToString() ?? string.Empty;
                switch (proType.Name)
                {
                    case "String": //字符串类型
                        newCell.SetCellValue(strVal);
                        break;
                    case "DateTime": //日期类型
                        DateTime dateV;
                        DateTime.TryParse(strVal, out dateV);
                        newCell.SetCellValue(dateV.ToString("yyyy-MM-dd HH:mm:ss"));
                        
                        break;
                    case "Boolean": //布尔型
                        var boolV = false;
                        bool.TryParse(strVal, out boolV);
                        newCell.SetCellValue(boolV);
                        break;
                    case "Int16": //整型
                    case "Int32":
                    case "Int64":
                    case "Byte":
                        var intV = 0;
                        int.TryParse(strVal, out intV);
                        newCell.SetCellValue(intV);
                        break;
                    case "Decimal": //浮点型
                    case "Double":
                        //double doubV = 0;
                        //double.TryParse(strVal, out doubV);
                        //newCell.SetCellValue(doubV);
                        newCell.SetCellValue(strVal);
                        break;
                    case "Nullable`1":
                        if (val != null)
                        {
                            var underlyingType = Nullable.GetUnderlyingType(proType);
                            proType = underlyingType;
                            continue;
                        }
                        break;
                    default:
                        newCell.SetCellValue("");
                        break;
                }
                break;
            }
        }

        private static object GetFieldValue(object source, string fieldName)
        {
            var t = source.GetType();
            var propertyInfo = t.GetProperty(fieldName);
            if (propertyInfo == null)
            {
                throw new Exception($"找不到名为{fieldName}的字段！");
            }

            return propertyInfo.GetValue(source, null);
        }

        private void WriteToFile(string fileSavePath)
        {
            //Write the stream data of workbook to the root directory
            var file = new FileStream(fileSavePath, FileMode.Create);
            _workbook.Write(file);
            file.Dispose();
        }

        private void Init(string configFilePath)
        {
            _config = ExportConfigureReader.ReadConfig(configFilePath, title, sheetname);

            #region 右击文件 属性信息

            var dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "四川飞牛巴士";
            _workbook.DocumentSummaryInformation = dsi;

            var si = PropertySetFactory.CreateSummaryInformation();
            si.Author = "四川飞牛巴士科技有限责任公司"; //填加xls文件作者信息
            si.ApplicationName = "飞牛巴士 - 商家版程序"; //填加xls文件创建程序信息
            si.LastAuthor = "四川飞牛巴士科技有限责任公司"; //填加xls文件最后保存者信息
            si.Comments = "四川飞牛巴士科技有限责任公司"; //填加xls文件作者信息
            si.Title = _config.Title; //填加xls文件标题信息
            si.Subject = _config.Title; //填加文件主题信息
            si.CreateDateTime = DateTime.Now;
            _workbook.SummaryInformation = si;

            #endregion
        }

        /// <summary>
        ///     写入标题行
        /// </summary>
        /// <returns>返回标题行下一行的RowIndex</returns>
        private void WriteSheetTitle(ISheet newsheet, ExportConfigure config)
        {
            #region 写入Title行

            //m_newsheet = m_workbook.CreateSheet(m_config.Sheetname);
            var headRow = newsheet.CreateRow(SheetTitleRowIndex);
            var headCell = headRow.CreateCell(0);

            headCell.SetCellValue(_config.Title);

            var mergedRegionCount = config.Cells.Count != 0 ? config.Cells.Count - 1 : 0;
            newsheet.AddMergedRegion(new CellRangeAddress(SheetTitleRowIndex, 0, SheetTitleRowIndex, mergedRegionCount));

            var mergededCell = newsheet.GetRow(0).GetCell(0);
            var style = CreateHasBroderCell(_workbook);
            style.Alignment = HorizontalAlignment.Center;
            var font = _workbook.CreateFont();
            font.FontHeight = 20 * 20;
            style.SetFont(font);
            mergededCell.CellStyle = style;

            #endregion
        }

        private void WriteDataHeader(ISheet newsheet, ExportConfigure config)
        {
            var cellIndex = 0;
            var newRow = newsheet.CreateRow(SheetDataHeaderRowIndex);
            newRow.Height = 20 * 20;

            //循环导出列
            foreach (var cfg in config.Cells)
            {
                var newCell = newRow.CreateCell(cellIndex);
                var style = CreateHasBroderCell(_workbook);
                if (cfg.Width != 0)
                {
                    newsheet.SetColumnWidth(cellIndex, cfg.Width * 256);
                }
                newCell.CellStyle = style;
                newCell.SetCellValue(cfg.Header);
                cellIndex++;
            }
        }

        private ISheet CreateSheet(ExportConfigure config, int sheetCount)
        {
            var sheetname = sheetCount == 0 ? config.SheetName : $"{config.SheetName}_{sheetCount}";
            var newsheet = _workbook.CreateSheet(sheetname);
            WriteSheetTitle(newsheet, config);
            WriteDataHeader(newsheet, config);
            return newsheet;
        }


    }
}
