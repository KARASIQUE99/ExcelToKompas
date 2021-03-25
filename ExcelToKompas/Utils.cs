using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Kompas6API5;
using Kompas6Constants;
using KAPITypes;

namespace ExcelToKompas
{
    class Utils
    {
        public static Excel._Worksheet getWorksheetByFileName(Excel.Application excelApp, string fileName)
        {
            try
            {
                excelApp.Workbooks.Close();
                var workBook = excelApp.Workbooks.Open(fileName);
                return workBook.ActiveSheet;
            } catch (Exception e)
            {
                return null;
            }
        }
        public static List<string> getFirstRowAsList(Excel._Worksheet workSheet)
        {
            try
            {
                List<string> output = new List<string>();
                output.Add("");
                for (int col = 1; col <= workSheet.UsedRange.Columns.Count; col++)
                {
                    string value = Convert.ToString(workSheet.Cells[1, col].Value);
                    if (value != null) output.Add(value);
                }
                return output;
            } catch (Exception e)
            {
                return null;
            }
        }

        public static string[] getEntities(ListBox.ObjectCollection collection)
        {
            string[] output = new string[collection.Count];
            for (int i = 0; i < collection.Count; i++)
                output[i] = collection[i].ToString();
            return output;
        }

        public static string getFileNameWithoutCount(string file)
        {
            try
            {
                string[] arr = file.Split(' ');
                arr[arr.Length - 1] = "";
                string output = string.Join(" ", arr).Trim();
                return output.Substring(0, output.Length - 1);
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public static string changeFileCount(string file, int newCount)
        {
            try
            {
                string[] arr = file.Split(' ');
                arr[arr.Length - 1] = Convert.ToString(newCount);
                return string.Join(" ", arr).Trim();
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public static bool createDirectory(string dirName)
        {
            if (!Directory.Exists(dirName))
            {
                Directory.CreateDirectory(dirName);
                return false;
            }
            else
            {
                var files = Directory.GetFiles(dirName);
                if (files.Length != 0)
                {
                    if (MessageBox.Show("Непустая директория с таким названием уже существует.\nПерезаписать ее?", "Перезаписать директорию?", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        System.IO.DirectoryInfo di = new DirectoryInfo(dirName);
                        foreach (FileInfo file in di.GetFiles()) file.Delete();
                        foreach (DirectoryInfo dir in di.GetDirectories()) dir.Delete(true);
                        return false;
                    } else return true;
                } else
                {
                    return false;
                }
            }
        }

        public static List<Row> group(List<Row> rows)
        {
            try
            {
                return rows.GroupBy(l => l.Name)
                    .Select(cl => new Row
                    {
                        Unknown = cl.First().Unknown,
                        Format = cl.First().Format,
                        Zone = cl.First().Zone,
                        Position = cl.Select(i => i.Position).Distinct().Aggregate((i, j) => i + ", " + j),
                        Mark = cl.First().Mark,
                        Name = cl.First().Name,
                        Count = cl.Sum(c => Convert.ToInt32(c.Count)).ToString(),
                        Note = cl.First().Note,
                        Mass = cl.First().Mass,
                        Material = cl.First().Material,
                        User = cl.Select(i => i.User).Distinct().Aggregate((i, j) => i + "; " + j),
                        Code = cl.First().Code,
                        Factory = cl.First().Factory,
                        DocumentNumber = cl.First().DocumentNumber,
                        DocumentName = cl.First().DocumentName,
                        DocumentCode = cl.First().DocumentCode,
                        CodeOKP = cl.First().CodeOKP,
                        Type = cl.First().Type
                    }).ToList();
            } catch (Exception e)
            {
                return rows;
            }

        }

        public static int getSectionByType(string type, List<List<string>> table)
        {
            for (int i = 0; i < table.Count; i++)
                if (table[i][0] == type) return Convert.ToInt32(table[i][2]);

            return 30;
        }

        public static int getSubsectionByType(string type, List<List<string>> table)
        {
            for (int i = 0; i < table.Count; i++)
                if (table[i][0] == type) return Convert.ToInt32(table[i][1]);

            return 200;
        }

        public static void DisplayInKompas(KompasObject kompas, List<Row> rows, string fileName, string styleName, List<List<string>> table, short styleType)
        {
            ksSpcDocument iDocumentSpc = (ksSpcDocument)kompas.SpcDocument();
            ksDocumentParam iDocumentParam = (ksDocumentParam)kompas.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);

            iDocumentParam.Init();
            iDocumentParam.type = (int)DocType.lt_DocSpc;

            ksSheetPar iSheetParam = (ksSheetPar)iDocumentParam.GetLayoutParam();
            iSheetParam.Init();
            iSheetParam.layoutName = styleName;
            iSheetParam.shtType = styleType;

            iDocumentSpc.ksCreateDocument(iDocumentParam);

            rows = group(rows);
            rows.Sort(new CompRow<Row>());

            for (int i = 1; i <= rows.Count; i++)
            {
                Console.WriteLine(i);

                var type = rows[i - 1].Type;

                ksSpecification iSpc = (ksSpecification)iDocumentSpc.GetSpecification();
                iSpc.ksSpcObjectCreate("", 0, getSectionByType(type, table), getSubsectionByType(type, table), 0, 0);
                int reference = iSpc.ksSpcObjectEnd();
                ksSpcObjParam iSpcObjParam = (ksSpcObjParam)kompas.GetParamStruct((short)StructType2DEnum.ko_SpcObjParam);
                iSpc.ksSpcObjectEdit(reference);
                iDocumentSpc.ksGetObjParam(reference, iSpcObjParam, ldefin2d.ALLPARAM);
                iSpcObjParam.blockNumber = 0;
                iSpcObjParam.draw = 1;
                iSpcObjParam.firstOnSheet = 0;
                iSpcObjParam.ispoln = 0;
                iSpcObjParam.posInc = 1;
                iSpcObjParam.posNotDraw = 0;
                iDocumentSpc.ksSetObjParam(reference, iSpcObjParam, ldefin2d.ALLPARAM);

                for (int j = 0; j < 12; j++)
                {
                    string val = rows[i - 1].getRowAsList()[j];
                    iSpc.ksSetSpcObjectColumnText(j, 1, 0, val);
                }

                reference = iSpc.ksSpcObjectEnd();
            }
            Console.WriteLine(fileName + ".spw");

            workStamp(kompas, iDocumentSpc);

            iDocumentSpc.ksSaveDocument(fileName + ".spw");
        }

        static void workStamp(KompasObject kompas, ksSpcDocument doc)
        {
            ksStamp stamp = (ksStamp)doc.GetStamp();
            stamp.ksOpenStamp();



            ksTextItemParam itemParam = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
            itemParam.Init();
            ksTextItemFont itemFont = (ksTextItemFont)itemParam.GetItemFont();
            itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);

            stamp.ksColumnNumber(1);
            itemParam.s = "Плата";
            stamp.ksTextLine(itemParam);

            stamp.ksColumnNumber(110);
            itemParam.s = "Халтуринский";
            stamp.ksTextLine(itemParam);

            stamp.ksColumnNumber(111);
            itemParam.s = "Краснов";
            stamp.ksTextLine(itemParam);

            stamp.ksColumnNumber(114);
            itemParam.s = "Деревенский";
            stamp.ksTextLine(itemParam);

            stamp.ksColumnNumber(115);
            itemParam.s = "Петров";
            stamp.ksTextLine(itemParam);

            stamp.ksColumnNumber(9);
            itemParam.s = "ООО \"Кванттелеком\"";
            stamp.ksTextLine(itemParam);

            stamp.ksColumnNumber(2);
            itemParam.s = "ГРТВ.000000.002";
            stamp.ksTextLine(itemParam);

            stamp.ksColumnNumber(25);
            itemParam.s = "ГРТВ.000000.001";
            stamp.ksTextLine(itemParam);

            stamp.ksCloseStamp();

        }


        public static string prepare(string columnName, string value)
        {
            if(columnName == "Designator")
            {
                StringBuilder sb = new StringBuilder();
                string[] arr = value.Replace('\n', ' ').Split(' ');
                for (int i = 0; i < arr.Length; i++)
                    if (i % 2 == 0) sb.Append(arr[i] + " ");
                    else sb.Append(arr[i].Replace(',', ' ').Trim() + "\n");


                return sb.ToString().Trim();
            }
            else
            {
                if(value.Length > 40)
                {
                    StringBuilder sb = new StringBuilder();
                    string[] arr = value.Split(' ');
                    int i = 0;
                    foreach (string e in arr)
                    {
                        i += e.Length;
                        Console.WriteLine(sb + "\n\n\n\n");
                        sb.Append(e);
                        sb.Append(" ");
                        if (i >= 20)
                        {
                            sb.Append("\n");
                            i = 0;
                        }
                    }

                    return sb.ToString().Trim();
                } else
                {
                    return value;
                }
            }
        }

        public static List<Row> GetDataFromExcel(Excel._Worksheet workSheet, Dictionary<string, string> names)
        {
            List<Row> result = new List<Row>();

            for (int row = 2; row <= workSheet.UsedRange.Rows.Count; row++)
            {
                Row instance = new Row();
                for (int col = 1; col <= workSheet.UsedRange.Columns.Count; col++)
                {
                    Excel.Range dataRange = workSheet.Cells[row, col];
                    string value = Convert.ToString(dataRange.Value);
                    
                    if (value != null)
                    {
                        string columnName = Convert.ToString(workSheet.Cells[1, col].Value);  
                        if (names.ContainsKey("Формат"))
                            if (columnName == names["Формат"])
                                instance.Format = prepare(columnName, value);
                        if (names.ContainsKey("Зона"))
                            if (columnName == names["Зона"])
                                instance.Zone = prepare(columnName, value);
                        if (names.ContainsKey("Позиция"))
                            if (columnName == names["Позиция"])
                                instance.Position = prepare(columnName, value);
                        if (names.ContainsKey("Обозначение"))
                            if (columnName == names["Обозначение"])
                                instance.Mark = prepare(columnName, value);
                        if (names.ContainsKey("Наименование"))
                            if (columnName == names["Наименование"])
                                instance.Name = prepare(columnName, value);
                        if (names.ContainsKey("Количество"))
                            if (columnName == names["Количество"])
                                instance.Count = prepare(columnName, value);
                        if (names.ContainsKey("Примечание"))
                            if (columnName == names["Примечание"])
                                instance.Note = prepare(columnName, value);
                        if (names.ContainsKey("Масса"))
                            if (columnName == names["Масса"])
                                instance.Mass = prepare(columnName, value);
                        if (names.ContainsKey("Материал"))
                            if (columnName == names["Материал"])
                                instance.Material = prepare(columnName, value);
                        if (names.ContainsKey("Пользователь"))
                            if (columnName == names["Пользователь"])
                                instance.User = prepare(columnName, value);
                        if (names.ContainsKey("Код"))
                            if (columnName == names["Код"])
                                instance.Code = prepare(columnName, value);
                        if (names.ContainsKey("Изготовитель"))
                            if (columnName == names["Изготовитель"])
                                instance.Factory = prepare(columnName, value);
                        if (names.ContainsKey("Тип изделия"))
                            if (columnName == names["Тип изделия"])
                                instance.Type = prepare(columnName, value);
                    }
                }
                if (instance.Name != "") result.Add(instance);
            }
            return result;
        }
    }
}
