using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelXmlTool
{
    class FormatTool
    {
        //这边需要把换行符号都转换一遍，避免直接被解析掉
        private List<string> _signList = new List<string>();
        private List<string> _signReplaceList = new List<string>();

        public FormatTool()
        {
            _signList.Add("&#10;");
            _signReplaceList.Add("###10###");
            _signList.Add("&#13;");
            _signReplaceList.Add("###13###");
        }

        //对外的接口
        //传入需要格式化的xml文件的全路径
        public void formatXmlFile(string xmlFileFullPath)
        {
            //先处理特殊符号
            replaceSignText(xmlFileFullPath);
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlFileFullPath);
            deleteWorksheetOptionsNode(ref doc);
            deleteDocumentPropertiesNode(ref doc);
            deleteOfficeDocumentSettingsNode(ref doc);
            deleteExcelWorkbookNode(ref doc);
            doc.Save(xmlFileFullPath);
            reduceSignText(xmlFileFullPath);
        }

        //还原特殊符号
        private void reduceSignText(string path)
        {
            //按行读取，存入list
            List<string> fileTextList = new List<string>();
            //先读取文件
            StreamReader sr = new StreamReader(path, Encoding.UTF8);
            String line;
            while ((line = sr.ReadLine()) != null)
            {
                for (int i = 0; i < _signList.Count; i++)
                {
                    line = Regex.Replace(line, _signReplaceList[i], _signList[i]);
                }
                fileTextList.Add(line.ToString());
            }
            sr.Close();
            //写入文件
            FileStream fs = new FileStream(path, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            //开始写入
            for (int i = 0; i < fileTextList.Count; i++)
            {
                sw.WriteLine(fileTextList[i]);
            }
            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();
        }

        //替换符号文件
        private void replaceSignText(string path)
        {
            //按行读取，存入list
            List<string> fileTextList = new List<string>();
            //先读取文件
            StreamReader sr = new StreamReader(path, Encoding.UTF8);
            String line;
            while ((line = sr.ReadLine()) != null)
            {
                for (int i = 0; i < _signList.Count; i++)
                {
                    line = Regex.Replace(line, _signList[i], _signReplaceList[i]);
                }
                fileTextList.Add(line.ToString());
            }
            sr.Close();
            //写入文件
            FileStream fs = new FileStream(path, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            //开始写入
            for (int i = 0; i < fileTextList.Count; i++)
            {
                sw.WriteLine(fileTextList[i]);
            }
            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();
        }

        //删除WorksheetOptions结点
        //结点位置位于Workbook->Worksheet->WorksheetOptions
        private void deleteWorksheetOptionsNode(ref XmlDocument doc)
        {
            XmlNode rootNode = doc.DocumentElement;
            for (int i = 0; i < rootNode.ChildNodes.Count; i++)
            {
                XmlNode workSheetNode = rootNode.ChildNodes[i];
                if (workSheetNode.Name == "Worksheet")
                {
                    for (int j = 0; j < workSheetNode.ChildNodes.Count; j++)
                    {
                        XmlNode worksheetOptionsNode = workSheetNode.ChildNodes[j];
                        if (worksheetOptionsNode.Name == "WorksheetOptions")
                        {
                            formatSheetVisable(ref worksheetOptionsNode);
                        }
                        else if (worksheetOptionsNode.Name == "Table")
                        {
                            formatCellIndex(ref worksheetOptionsNode, ref doc);
                        }
                        else if (worksheetOptionsNode.Name == "x:WorksheetOptions")
                        {
                            workSheetNode.RemoveChild(worksheetOptionsNode);
                            j--;
                        }
                    }
                }
            }
        }

        //处理index的问题
        //就是新建一个<Cell/>结点
        private void formatCellIndex(ref XmlNode tableBode, ref XmlDocument doc)
        {
            List<XmlNode> newXmlNodeList = new List<XmlNode>();
            XmlElement nullCellNode;
            XmlElement nullRowNode;
            XmlElement dataNode;
            int i;
            //标记行的数量
            int rowNodeNumber = 0;
            for (int k = 0; k < tableBode.ChildNodes.Count; k++)
            {
                XmlNode rowNode = tableBode.ChildNodes[k];
                //标记合并单元格的数量
                int mergeAcrossNumber = 0;
                //判断是否有空行
                //行结点
                if (rowNode.Name == "Row")
                {
                    rowNodeNumber++;
                    if (rowNode.Attributes["ss:Index"] != null)
                    {
                        int nowRowIndex = Int32.Parse(rowNode.Attributes["ss:Index"].Value);
                        //移除相关的index属性
                        rowNode.Attributes.Remove(rowNode.Attributes["ss:Index"]);
                        for (int l = rowNodeNumber; l < nowRowIndex; l++)
                        {
                            nullRowNode = doc.CreateElement("Row", doc.DocumentElement.NamespaceURI);
                            nullRowNode.RemoveAllAttributes();
                            tableBode.InsertBefore(nullRowNode, rowNode);
                        }
                    }
                    newXmlNodeList.Clear();
                    for (int j = 0; j < rowNode.ChildNodes.Count; j++)
                    {
                        XmlNode cellNode = rowNode.ChildNodes[j];
                        if (cellNode.Name == "Cell")
                        {
                            //判断有没有合并单元格的情况，属性名是ss:MergeAcross，属性值表示合并了几个格子（不包括自己）
                            if (cellNode.Attributes["ss:MergeAcross"] != null)
                            {
                                //mergeAcrossNumber = Int32.Parse(cellNode.Attributes["ss:MergeAcross"].Value);
                            }
                            if (cellNode.Attributes["ss:Index"] != null)
                            {
                                int index = Int32.Parse(cellNode.Attributes["ss:Index"].Value);
                                for (i = newXmlNodeList.Count; i < index - 1 - mergeAcrossNumber; i++)
                                {
                                    nullCellNode = doc.CreateElement("Cell", doc.DocumentElement.NamespaceURI);
                                    //nullDataNode = doc.CreateElement("Data");
                                    //nullDataNode.SetAttribute("ss:Type" , "string");
                                    //nullDataNode.InnerText = "";
                                    //nullCellNode.AppendChild(nullDataNode);
                                    nullCellNode.RemoveAllAttributes();
                                    newXmlNodeList.Add(nullCellNode);
                                }
                                cellNode.Attributes.Remove(cellNode.Attributes["ss:Index"]);
                                mergeAcrossNumber = 0;
                            }
                            if (cellNode.ChildNodes.Count > 0 && cellNode.ChildNodes[0].Name == "ss:Data")
                            {
                                string data = "";
                                for (int l = 0; l < cellNode.ChildNodes[0].ChildNodes.Count; l++)
                                {
                                    data = data + cellNode.ChildNodes[0].ChildNodes[l].InnerText;
                                }
                                cellNode.ChildNodes[0].InnerText = data;
                            }
                            newXmlNodeList.Add(cellNode);
                        }
                    }
                    //移除所有的子节点
                    int childNodeNumber = rowNode.ChildNodes.Count;
                    for (int l = 0; l < childNodeNumber; l++)
                    {
                        rowNode.RemoveChild(rowNode.ChildNodes[0]);
                    }
                    //开始替换节点
                    for (i = 0; i < newXmlNodeList.Count; i++)
                    {
                        rowNode.AppendChild(newXmlNodeList[i]);
                    }
                }
            }
        }

        //处理隐藏table的问题
        private void formatSheetVisable(ref XmlNode worksheetOptionsNode)
        {
            XmlNode childNode = null;
            int len = worksheetOptionsNode.ChildNodes.Count;
            for (int i = 0; i < len; i++)
            {
                if (worksheetOptionsNode.ChildNodes[0].Name == "Visible")
                {
                    childNode = worksheetOptionsNode.ChildNodes[0];
                }
                worksheetOptionsNode.RemoveChild(worksheetOptionsNode.ChildNodes[0]);
            }
            if (childNode != null)
            {
                worksheetOptionsNode.AppendChild(childNode);
            }
        }

        //删除DocumentProperties节点
        private void deleteDocumentPropertiesNode(ref XmlDocument doc)
        {
            XmlNode rootNode = doc.DocumentElement;
            for (int i = 0; i < rootNode.ChildNodes.Count; i++)
            {
                XmlNode documentPropertiesNode = rootNode.ChildNodes[i];
                if (documentPropertiesNode.Name == "DocumentProperties")
                {
                    //标记有格式化
                    Program.haveFormat = true;
                    rootNode.RemoveChild(documentPropertiesNode);
                }
            }
        }

        //删除OfficeDocumentSettings节点
        private void deleteOfficeDocumentSettingsNode(ref XmlDocument doc)
        {
            XmlNode rootNode = doc.DocumentElement;
            for (int i = 0; i < rootNode.ChildNodes.Count; i++)
            {
                XmlNode officeDocumentSettingsNode = rootNode.ChildNodes[i];
                if (officeDocumentSettingsNode.Name == "OfficeDocumentSettings")
                {
                    rootNode.RemoveChild(officeDocumentSettingsNode);
                }
            }
        }

        //删除ExcelWorkbook节点
        private void deleteExcelWorkbookNode(ref XmlDocument doc)
        {
            XmlNode rootNode = doc.DocumentElement;
            for (int i = 0; i < rootNode.ChildNodes.Count; i++)
            {
                XmlNode excelWorkbookNode = rootNode.ChildNodes[i];
                if (excelWorkbookNode.Name == "ExcelWorkbook")
                {
                    rootNode.RemoveChild(excelWorkbookNode);
                }
            }
        }
    }
}
