using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Transactions;
using Fclp;
using Fclp.Internals.Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using XMind2Xls.XMind;

namespace XMind2Xls
{
    class Program
    {

        public static int WriteXlsTopic(ExcelWorksheet worksheet, MindTopic topic, List<string> pPath, ref int row, List<string> pColumnHeaders)
        {
            pPath.Add(topic.title);
            
            if (topic.children != null)
            {
                foreach (var lSubTopic in topic.children.attached)
                {
                    WriteXlsTopic(worksheet, lSubTopic, pPath, ref row, pColumnHeaders);
                }

                pPath.Remove(pPath.Last());
                return topic.children.attached.Count;
            }


            int lColumnIdx = 1;
            int lColumnValueIdx = -1;
            foreach (var lColumnValue in pPath)
            {
                if (pColumnHeaders.Contains(lColumnValue.ToUpper()))
                {
                    lColumnValueIdx = pColumnHeaders.IndexOf(lColumnValue.ToUpper());
                }
                else
                {
                    if (lColumnValueIdx != -1)
                    {
                        worksheet.Cells[row, lColumnValueIdx + 1].Value = lColumnValue.ToUpper();
                        lColumnIdx++;
                    }
                    else
                    {
                        worksheet.Cells[row, lColumnIdx].Value = lColumnValue.ToUpper();
                        lColumnIdx++;
                    }
                   
                }
                
            }

            pPath.Remove(pPath.Last());
            row++;

            return 1;
        }

        public static void WriteXls(string filename, IList<MindSheet> sheets, List<string> pColumnHeaders)
        {
            File.Delete(filename);
            using (var lPackage = new ExcelPackage(new FileInfo(filename)))
            {
                foreach (var lSheet in sheets)
                {
                    ExcelWorksheet lWorkSheet = lPackage.Workbook.Worksheets.Add(lSheet.title);
                    int lColumnIdx = 1;
                    int lBaseRow = 1;
                    foreach (var lHeader in pColumnHeaders)
                    {
                        lWorkSheet.Cells[lBaseRow, lColumnIdx].Value = lHeader;
                        lColumnIdx++;
                    }

                    if (pColumnHeaders.IsNullOrEmpty())
                    {
                        lBaseRow = 1;
                    }
                    else
                    {
                        lBaseRow = 2;
                    }

                    WriteXlsTopic(lWorkSheet, lSheet.rootTopic, new List<string>(), ref lBaseRow, pColumnHeaders);

                    using (ExcelRange lRange = lWorkSheet.Cells[lWorkSheet.Dimension.Address])
                    {
                        lRange.AutoFitColumns();
                        lRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        lRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        lRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        lRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        lRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        lRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }
                }
                lPackage.Save();
            }
        }

        public static void MergeCells(string pFilename, IList<MindSheet> pSheets, List<string> pColumnHeaders, int pNoHeader)
        {
            using (var lPackage = new ExcelPackage(new FileInfo(pFilename)))
            {
                foreach (var lSheet in pSheets)
                {
                    ExcelWorksheet lWorkSheet = lPackage.Workbook.Worksheets.FirstOrDefault(pWorksheet => lSheet.title.Contains(pWorksheet.Name));
                    if (lWorkSheet != null)
                    {
                        string lPreviousRow = "";
                        int lNbRowToMerge = 0;
                        for (int lRowIndex = 1; lRowIndex <= 1000; lRowIndex++)
                        {
                            string lCurrentRow = "";
                            for (int lColumnIndex = 1; lColumnIndex <= pNoHeader; lColumnIndex++)
                            {
                                lCurrentRow += lWorkSheet.Cells[lRowIndex, lColumnIndex].Value;
                                lCurrentRow += ";";
                            }

                            if (lCurrentRow != lPreviousRow)
                            {
                                if (lNbRowToMerge > 1)
                                {
                                    List<string> lColValues = new List<string>();
                                    for (int lColumnIndex = 1; lColumnIndex <= pNoHeader; lColumnIndex++)
                                    {
                                        lColValues.Add(lWorkSheet.Cells[lRowIndex - lNbRowToMerge, lColumnIndex].Value
                                            ?.ToString());
                                    }

                                    for (int lColumnIndex = pNoHeader + 1;
                                        lColumnIndex <= pColumnHeaders.Count;
                                        lColumnIndex++)
                                    {
                                        bool lHasValue = false;
                                        for (int lSubRow = lRowIndex - lNbRowToMerge;
                                            lSubRow <= lRowIndex - 1;
                                            lSubRow++)
                                        {
                                            if (lWorkSheet.Cells[lSubRow, lColumnIndex].Value != null)
                                            {
                                                if (lColValues.Count < lColumnIndex)
                                                {
                                                    lColValues.Add(lWorkSheet.Cells[lSubRow, lColumnIndex].Value
                                                        .ToString());
                                                    lHasValue = true;
                                                }
                                            }
                                        }

                                        if (lHasValue == false)
                                        {
                                            lColValues.Add("");
                                        }
                                    }

                                    for (int lColumnIndex = 1; lColumnIndex <= pColumnHeaders.Count; lColumnIndex++)
                                    {
                                        lWorkSheet.Cells[lRowIndex - lNbRowToMerge, lColumnIndex].Value =
                                            lColValues[lColumnIndex - 1];
                                    }

                                    int lLowRow = lRowIndex - lNbRowToMerge + 1;
                                    int lNbRows = (lRowIndex - 1) - lLowRow + 1;
                                    lWorkSheet.DeleteRow(lLowRow, lNbRows);
                                    lRowIndex -= lNbRows;
                                }

                                lNbRowToMerge = 1;
                                lPreviousRow = lCurrentRow;
                            }
                            else
                            {
                                lNbRowToMerge++;
                            }
                        }

                        for (int lRowIndex = 1; lRowIndex <= 1000; lRowIndex++)
                        {
                            for (int lColumnIndex = 1; lColumnIndex <= pColumnHeaders.Count; lColumnIndex++)
                            {
                                lWorkSheet.Cells[lRowIndex, lColumnIndex].Style.WrapText = true;
                            }
                        }
                    }
                }
                lPackage.Save();
            }
        }

        public static void CountItemPerLevel(MindTopic topic, int level, ref List<int> itemPerLevel)
        {
            if (itemPerLevel.Count <= level)
            {
                itemPerLevel.Add(1);
            }
            else
            {
                itemPerLevel[level]++;
            }

            if (topic.children != null)
            {
                foreach (var lSubTopic in topic.children.attached)
                {
                    CountItemPerLevel(lSubTopic, level + 1, ref itemPerLevel);
                }
            }
        }

        static void Main(string[] args)
        {
            Console.WriteLine("XMind2Xls v" + Assembly.GetExecutingAssembly().GetName().Version);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var lCommandLineParser = new FluentCommandLineParser<ApplicationArguments>();

            // Specifies all options.
            lCommandLineParser.Setup(arg => arg.InputFile).As('i', "inputFile").Required();
            lCommandLineParser.Setup(arg => arg.OutputFile).As('o', "outputFile").Required();
            lCommandLineParser.Setup(arg => arg.Headers).As('h', "headers").Required();
            lCommandLineParser.Setup(arg => arg.RootAsWorksheet).As('w', "worksheet").SetDefault(true);
            lCommandLineParser.Setup(arg => arg.Silent).As('s', "silent").SetDefault(false); 

            var lResult = lCommandLineParser.Parse(args);

            if (lResult.HasErrors == false)
            {
                Console.WriteLine("Load file " + lCommandLineParser.Object.InputFile);

                MindFile lFile = new MindFile();
                lFile.ReadXMind(lCommandLineParser.Object.InputFile);
                
                Console.WriteLine("Version of the file : " + lFile.Version);
                List<string> pColumnHeaders = new List<string>();
                if (File.Exists(lCommandLineParser.Object.Headers))
                {
                    Console.WriteLine("Load headers file " + lCommandLineParser.Object.Headers);
                    pColumnHeaders = File.ReadAllLines(lCommandLineParser.Object.Headers).ToList();
                    pColumnHeaders = pColumnHeaders.ConvertAll(i => i.ToUpper());
                }

                int lNoHeader = 0;
                foreach (var lColumnHeader in pColumnHeaders)
                {
                    if (string.IsNullOrEmpty(lColumnHeader))
                    {
                        lNoHeader++;
                    }
                }

                if (lFile.Sheets != null && lFile.Sheets.Count != 0)
                {
                    Console.WriteLine("Write output file " + lCommandLineParser.Object.OutputFile);
                    WriteXls(lCommandLineParser.Object.OutputFile, lFile.Sheets, pColumnHeaders);
                    MergeCells(lCommandLineParser.Object.OutputFile, lFile.Sheets, pColumnHeaders, lNoHeader);
                }
               
            }
        }
    }
}
