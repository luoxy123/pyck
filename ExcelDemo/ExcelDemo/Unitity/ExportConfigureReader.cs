using System;
using System.IO;
using System.Xml;

namespace ExcelDemo.Unitity
{
    public class ExportConfigureReader
    {
        public static ExportConfigure ReadConfig(string configureFilePath, string title, string sheetName)
        {
            var config = Read(configureFilePath, title, sheetName);
            return config;
        }

        private static XmlNode GetSingleNode(XmlNode doc, string xmlPath)
        {
            var node = doc.SelectSingleNode(xmlPath);
            return node;

        }

        private static string GetAttributeValue(XmlNode node, string attributeName)
        {
            var attr = node.Attributes?[attributeName];
            return attr == null ? string.Empty : attr.Value;
        }

        private static ExportConfigure Read(string configFilePath, string title, string sheetName)
        {
            var doc = new XmlDocument();
            var file = File.Open(configFilePath, FileMode.Open);
            doc.Load(file);
            file.Dispose();
            var config = new ExportConfigure(title, sheetName);

            var xmlNode = GetSingleNode(doc, "configuration/fileInfo");

            config.Title = string.IsNullOrEmpty(config.Title) ? GetAttributeValue(xmlNode, "title") : config.Title;
            config.Mode = GetAttributeValue(xmlNode, "mode");
            config.Relationship = GetAttributeValue(xmlNode, "relationship");
            config.SheetName = string.IsNullOrEmpty(config.SheetName) ? GetAttributeValue(xmlNode, "sheetName") : config.SheetName;
            var strRowHeight = GetAttributeValue(xmlNode, "rowHeight");
            var strMerge = GetAttributeValue(xmlNode, "IsMerge");
            config.RowHeight = string.IsNullOrEmpty(strRowHeight) ? short.MinValue : short.Parse(strRowHeight);
            config.IsMerge = !string.IsNullOrEmpty(strMerge) && bool.Parse(strMerge);
            
            #region 读取列表内容配置信息

            var nodes = GetSingleNode(doc, "configuration/columns");
            foreach (XmlNode n in nodes.ChildNodes)
            {
                if (n.NodeType == XmlNodeType.Comment)
                {
                    continue;
                }

                var propertoryName = GetAttributeValue(n, "name");
                var header = GetAttributeValue(n, "header");
                var index = GetAttributeValue(n, "index");
                var colIndex = string.IsNullOrEmpty(index) ? -1 : int.Parse(index);
                if (colIndex == -1)
                {
                    throw new Exception("请指定Column节的index值，值必须是正整数！");
                }

                var strWidth = GetAttributeValue(n, "width");
                var width = string.IsNullOrEmpty(strWidth) ? 0 : int.Parse(strWidth);

                var formula = GetAttributeValue(n, "formula");

                var strIsNumberColumn = GetAttributeValue(n, "isNumberColumn");
                var isNumberColumn = !string.IsNullOrEmpty(strIsNumberColumn) && bool.Parse(strIsNumberColumn);

                var strIsMerge = GetAttributeValue(n, "IsMerge");
                var isMerge = !string.IsNullOrEmpty(strIsMerge) && bool.Parse(strIsMerge);

                if (isNumberColumn)
                {
                    //propertoryName = ExcelExportTemplate.NumColumnPropertoryName;
                }
                var cell = new CellInfo(propertoryName, header, colIndex, formula, width, isNumberColumn, isMerge);
                config.Cells.Add(cell);
            }

            #endregion

            return config;
        }
    }
}
