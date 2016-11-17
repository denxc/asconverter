﻿using NPOI.HSSF.Model;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using NPOI.XSSF.UserModel;

namespace ASConverter {
    public class ASExporter {

        private const int rowHeigthMultiplicator = 568;
        private const int rowWidthMultiplicator = 1309;

        private const string DEFAULT_REPORT = "Консолидированный отчет";
        private const string START_SHIELD_NAME = "Empty";

        private static string[] defaultColumns = { "Клиент", "Книга продаж", "Дата", "Приход", "Расход", "Контрагент",
                                                "ИНН", "Ставка НДС", "Сумма НДС", "Назначение платежа", "П/п", "КПП",
                                                "Р/счет контрагента", "Банк контрагента" };
        
        private const int DEFAULT_ROW_HEIGHT = (int)(0.5 * rowHeigthMultiplicator);
        private const string DOUBLE_FORMAT = "#,##0.00";

        public static void Generate(string aXlsFile) {
            var wb = HSSFWorkbook.Create(InternalWorkbook.CreateWorkbook());
            var sh = (HSSFSheet)wb.CreateSheet(START_SHIELD_NAME);
            using (var fs = new FileStream(aXlsFile, FileMode.Create, FileAccess.Write)) {
                wb.Write(fs);
            }
        }

        public static string[] GetShields(string aXlsFile) {
            var shields = new List<string>();
            using (var fs = new FileStream(aXlsFile, FileMode.Open, FileAccess.Read)) {
                var wb = new HSSFWorkbook(fs);
                for (int i = 0; i < wb.Count; i++) {
                    var shieldName = wb.GetSheetAt(i).SheetName;
                    if (!shieldName.Equals(START_SHIELD_NAME) && !shieldName.Equals(DEFAULT_REPORT)) {
                        shields.Add(shieldName);
                    }                    
                }
            }

            return shields.ToArray();
        }

        public static string GetCompanyInn(string aXlsFile) {            
            using (var fs = new FileStream(aXlsFile, FileMode.Open, FileAccess.Read)) {
                var wb = new HSSFWorkbook(fs);
                for (int i = 0; i < wb.Count; i++) {
                    var shield = wb.GetSheetAt(i);
                    var companyInn = shield?.GetRow(1)?.GetCell(6)?.StringCellValue;
                    if (!string.IsNullOrEmpty(companyInn)) {
                        return companyInn;
                    }                    
                }
            }

            return string.Empty;
        }

        public static void CreateNewShield(string aXlsFile, string shieldName, OrderEntity aOrder, double aStartAmount) {
            HSSFWorkbook wb = null;
            using (var fs = new FileStream(aXlsFile, FileMode.Open, FileAccess.Read)) {
                wb = new HSSFWorkbook(fs);
                fs.Close();                       
            }
            
            var defaultReportShield = wb.GetSheet(DEFAULT_REPORT);
            if (defaultReportShield == null) {
                defaultReportShield = wb.CreateSheet(DEFAULT_REPORT);
                wb.SetSheetOrder(DEFAULT_REPORT, 0);
                PrepareShieldTop(defaultReportShield, aOrder, aStartAmount);
            } else {
                var row = defaultReportShield.GetRow(0);
                var cell = row.GetCell(3);
                var cellValue = cell.NumericCellValue;
                cell.SetCellValue(cellValue + aStartAmount);                
            }

            var defaultReportIndex = wb.GetSheetIndex(defaultReportShield);

            var sheet = wb.CreateSheet(shieldName);
            wb.SetSheetOrder(shieldName, defaultReportIndex);
            PrepareShieldTop(sheet, aOrder, aStartAmount);

            var emptyShield = wb.GetSheet(START_SHIELD_NAME);
            if (emptyShield != null) {
                wb.SetSheetHidden(wb.GetSheetIndex(emptyShield), SheetState.VeryHidden);
            }

            wb.SetActiveSheet(0);

            using (var fs = new FileStream(aXlsFile, FileMode.Create, FileAccess.Write)) {
                wb.Write(fs);
                fs.Close();
            }
        }

        private static void PrepareShieldTop(ISheet sheet, OrderEntity aOrder, double aStartAmount) {
            var boldFont = sheet.Workbook.CreateFont();
            boldFont.FontName = "Calibri";
            boldFont.FontHeightInPoints = 11;
            boldFont.IsBold = true;

            var defaultFont = sheet.Workbook.CreateFont();
            defaultFont.FontName = "Calibri";
            defaultFont.FontHeightInPoints = 11;            

            var doubleCellBoldStyle = sheet.Workbook.CreateCellStyle();
            doubleCellBoldStyle.DataFormat = sheet.Workbook.CreateDataFormat().GetFormat(DOUBLE_FORMAT);
            doubleCellBoldStyle.BorderBottom = BorderStyle.Medium;
            doubleCellBoldStyle.BorderTop = BorderStyle.Medium;
            doubleCellBoldStyle.BorderLeft = BorderStyle.Medium;
            doubleCellBoldStyle.BorderRight = BorderStyle.Medium;
            doubleCellBoldStyle.SetFont(boldFont);

            var doubleCellDefaultStyle = sheet.Workbook.CreateCellStyle();
            doubleCellDefaultStyle.DataFormat = sheet.Workbook.CreateDataFormat().GetFormat(DOUBLE_FORMAT);
            doubleCellDefaultStyle.BorderBottom = BorderStyle.Medium;
            doubleCellDefaultStyle.BorderTop = BorderStyle.Medium;
            doubleCellDefaultStyle.BorderLeft = BorderStyle.Medium;
            doubleCellDefaultStyle.BorderRight = BorderStyle.Medium;
            doubleCellDefaultStyle.SetFont(defaultFont);

            var stringCellBoldStyle = sheet.Workbook.CreateCellStyle();            
            stringCellBoldStyle.BorderBottom = BorderStyle.Medium;
            stringCellBoldStyle.BorderTop = BorderStyle.Medium;
            stringCellBoldStyle.BorderLeft = BorderStyle.Medium;
            stringCellBoldStyle.BorderRight = BorderStyle.Medium;
            stringCellBoldStyle.Alignment = HorizontalAlignment.Center;
            stringCellBoldStyle.SetFont(boldFont);

            var stringCellDefaultStyle = sheet.Workbook.CreateCellStyle();
            stringCellDefaultStyle.BorderBottom = BorderStyle.Medium;
            stringCellDefaultStyle.BorderTop = BorderStyle.Medium;
            stringCellDefaultStyle.BorderLeft = BorderStyle.Medium;
            stringCellDefaultStyle.BorderRight = BorderStyle.Medium;
            stringCellDefaultStyle.Alignment = HorizontalAlignment.Center;
            stringCellDefaultStyle.SetFont(defaultFont);

            var stringCellDefaultLeftStyle = sheet.Workbook.CreateCellStyle();
            stringCellDefaultLeftStyle.BorderBottom = BorderStyle.Medium;
            stringCellDefaultLeftStyle.BorderTop = BorderStyle.Medium;
            stringCellDefaultLeftStyle.BorderLeft = BorderStyle.Medium;
            stringCellDefaultLeftStyle.BorderRight = BorderStyle.Medium;
            stringCellDefaultLeftStyle.Alignment = HorizontalAlignment.Left;
            stringCellDefaultLeftStyle.SetFont(defaultFont);

            var unicCellStyle = sheet.Workbook.CreateCellStyle();
            unicCellStyle.BorderBottom = BorderStyle.Medium;
            unicCellStyle.SetFont(defaultFont);        

            var row = sheet.CreateRow(0);
            row.Height = DEFAULT_ROW_HEIGHT;
            var cell = row.CreateCell(0, CellType.String);
            cell.CellStyle = unicCellStyle;            
            cell.SetCellValue("Остаток на начало периода");
            cell = row.CreateCell(1, CellType.String);
            cell.CellStyle = unicCellStyle;
            cell = row.CreateCell(2, CellType.String);
            cell.CellStyle = unicCellStyle;
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, 2));

            cell = row.CreateCell(3, CellType.Numeric);            
            cell.CellStyle = doubleCellDefaultStyle;            
            cell.SetCellValue(aStartAmount);

            row = sheet.CreateRow(1);
            row.Height = DEFAULT_ROW_HEIGHT;
            cell = row.CreateCell(0, CellType.String);
            cell.CellStyle = unicCellStyle;            
            cell.SetCellValue("Обороты за период");
            cell = row.CreateCell(1, CellType.String);
            cell.CellStyle = unicCellStyle;
            cell = row.CreateCell(2, CellType.String);
            cell.CellStyle = unicCellStyle;
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(1, 1, 0, 2));
            cell = row.CreateCell(3, CellType.Numeric);
            cell.CellStyle = doubleCellDefaultStyle;
            cell.SetCellValue(0.0);
            cell = row.CreateCell(4, CellType.Numeric);
            cell.CellStyle = doubleCellDefaultStyle;
            cell.SetCellValue(0.0);
            cell = row.CreateCell(5, CellType.String);
            cell.CellStyle = stringCellBoldStyle;            
            cell.SetCellValue(aOrder.OwnerName);
            cell = row.CreateCell(6, CellType.String);
            cell.CellStyle = stringCellDefaultStyle;            
            cell.SetCellValue(string.IsNullOrEmpty(aOrder.OwnerInn) ? string.Empty : aOrder.OwnerInn);
            cell = row.CreateCell(11, CellType.String);
            cell.CellStyle = stringCellDefaultLeftStyle;            
            cell.SetCellValue(string.IsNullOrEmpty(aOrder.OwnerKPP) ? string.Empty : aOrder.OwnerKPP);
            cell = row.CreateCell(12, CellType.String);
            cell.CellStyle = stringCellDefaultLeftStyle;            
            cell.SetCellValue(string.IsNullOrEmpty(aOrder.OwnerAccount) ? string.Empty : aOrder.OwnerAccount);            
            cell = row.CreateCell(13, CellType.String);
            cell.CellStyle = stringCellDefaultLeftStyle;            
            cell.SetCellValue(aOrder.OwnerBank);

            row = sheet.CreateRow(2);
            row.Height = DEFAULT_ROW_HEIGHT;
            cell = row.CreateCell(0, CellType.String);            
            cell.SetCellValue("Остаток на конец периода");
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(2, 2, 0, 2));
            cell = row.CreateCell(3, CellType.Numeric);
            cell.CellStyle = doubleCellBoldStyle;
            cell.SetCellFormula("D1+D2-E2");

            row = sheet.CreateRow(3);
            row.Height = DEFAULT_ROW_HEIGHT;
            for (var i = 0; i < defaultColumns.Length; ++i) {
                cell = row.CreateCell(i);
                cell.CellStyle = stringCellBoldStyle;
                cell.SetCellValue(defaultColumns[i]);
            }

            sheet.SetColumnWidth(0, (int)(2.63 * rowWidthMultiplicator));
            sheet.SetColumnWidth(1, (int)(2.63 * rowWidthMultiplicator));
            sheet.SetColumnWidth(2, (int)(1.7 * rowWidthMultiplicator));
            sheet.SetColumnWidth(3, (int)(2.91 * rowWidthMultiplicator));
            sheet.SetColumnWidth(4, (int)(2.91 * rowWidthMultiplicator));
            sheet.SetColumnWidth(5, (int)(9.79 * rowWidthMultiplicator));
            sheet.SetColumnWidth(6, (int)(2.85 * rowWidthMultiplicator));
            sheet.SetColumnWidth(7, (int)(2.3 * rowWidthMultiplicator));
            sheet.SetColumnWidth(8, (int)(2.46 * rowWidthMultiplicator));
            sheet.SetColumnWidth(9, (int)(20.6 * rowWidthMultiplicator));
            sheet.SetColumnWidth(10, (int)(1.74 * rowWidthMultiplicator));
            sheet.SetColumnWidth(11, (int)(2.17 * rowWidthMultiplicator));
            sheet.SetColumnWidth(12, (int)(4.65 * rowWidthMultiplicator));
            sheet.SetColumnWidth(13, (int)(18.0 * rowWidthMultiplicator));

            sheet.CreateFreezePane(0, 4);
        }

        public static int Export(string destFilePath, string aShieldName, OrderEntity[] orders, double endAmount, out string message) {
            HSSFWorkbook wb = null;
            using (var fs = new FileStream(destFilePath, FileMode.Open, FileAccess.Read)) {
                wb = new HSSFWorkbook(fs);
                fs.Close();
            }

            ClearAllColors(wb);

            var countAdded = 0;
            var shield = wb.GetSheet(aShieldName);
            var defaultReport = wb.GetSheet(DEFAULT_REPORT);
            foreach (var order in orders) {
                if (AddOrdeToShield(shield, order)) {
                    AddOrdeToShield(defaultReport, order);
                    countAdded++;                    
                }
            }

            var defaultAmouond = GetEndAmount(defaultReport);
            var curEndAmount = GetEndAmount(shield);
            if (Math.Abs(curEndAmount - endAmount) > 0.01) {
                message = string.Format("Возможно, за некоторый период выписки не импортированы.\nОжидался остаток: {0:0.##}, остаток в отчете: {1:0.##}", endAmount, curEndAmount);
                ColorEndAmound(shield, true);
            } else {
                message = string.Empty;
                ColorEndAmound(shield, false);
            }
            
            using (var fs = new FileStream(destFilePath, FileMode.Create, FileAccess.Write)) {
                wb.Write(fs);
                fs.Close();
            }
                        
            return countAdded;
        }

        private static void ColorEndAmound(ISheet shield, bool isRed) {
            var defaultFont = shield.Workbook.CreateFont();
            defaultFont.FontName = "Calibri";
            defaultFont.Color = isRed ? HSSFColor.Red.Index : HSSFColor.Black.Index;
            defaultFont.IsBold = true;
            defaultFont.FontHeightInPoints = 11;

            var doubleCellStyle = shield.Workbook.CreateCellStyle();
            doubleCellStyle.DataFormat = shield.Workbook.CreateDataFormat().GetFormat(DOUBLE_FORMAT);
            doubleCellStyle.BorderRight = BorderStyle.Medium;
            doubleCellStyle.BorderLeft = BorderStyle.Medium;
            doubleCellStyle.SetFont(defaultFont);

            var row = shield.GetRow(2);
            row.GetCell(3).CellStyle = doubleCellStyle;
        }

        private static double GetEndAmount(ISheet shield) {
            var row = shield.GetRow(1);
            var prihod = row.GetCell(3);
            var rashod = row.GetCell(4);

            prihod.SetCellFormula(string.Format("SUM(D{0}: D{1})", 5, shield.LastRowNum + 1));
            rashod.SetCellFormula(string.Format("SUM(E{0}: E{1})", 5, shield.LastRowNum + 1));

            XSSFFormulaEvaluator.EvaluateAllFormulaCells(shield.Workbook);

            row = shield.GetRow(2);
            var amountCell = row.GetCell(3);
            var amount = amountCell.NumericCellValue;

            return amount;
        }        

        private static bool AddOrdeToShield(ISheet shield, OrderEntity order) {
            var position = FindBestPosition(shield, order);            
            if (position < 0) {
                return false;
            }

            var defaultFont = shield.Workbook.CreateFont();
            defaultFont.FontName = "Calibri";
            defaultFont.Color = HSSFColor.Green.Index;
            defaultFont.FontHeightInPoints = 11;

            var doubleCellStyle = shield.Workbook.CreateCellStyle();
            doubleCellStyle.DataFormat = shield.Workbook.CreateDataFormat().GetFormat(DOUBLE_FORMAT);
            doubleCellStyle.BorderRight = BorderStyle.Medium;
            doubleCellStyle.SetFont(defaultFont);

            var stringCellStyle = shield.Workbook.CreateCellStyle();            
            stringCellStyle.BorderRight = BorderStyle.Medium;
            stringCellStyle.SetFont(defaultFont);

            var ndsCellStyle = shield.Workbook.CreateCellStyle();
            ndsCellStyle.BorderRight = BorderStyle.Medium;
            ndsCellStyle.Alignment = HorizontalAlignment.Right;
            ndsCellStyle.SetFont(defaultFont);

            var dataCellStyle = shield.Workbook.CreateCellStyle();
            dataCellStyle.DataFormat = shield.Workbook.CreateDataFormat().GetFormat("dd.mm.yy");
            dataCellStyle.BorderRight = BorderStyle.Medium;
            dataCellStyle.SetFont(defaultFont);

            var row = InsertRow(shield, position);            
            row.Height = DEFAULT_ROW_HEIGHT;

            // клиент.
            var cell = row.CreateCell(0);
            cell.CellStyle = stringCellStyle;            
            cell.SetCellValue(string.Empty);

            // книга продаж.
            cell = row.CreateCell(1);
            cell.CellStyle = stringCellStyle;            
            cell.SetCellValue(string.Empty);

            // дата.
            cell = row.CreateCell(2);
            cell.CellStyle = dataCellStyle;
            cell.SetCellValue(order.date.Date);            

            // Приход.
            cell = row.CreateCell(3);
            cell.CellStyle = doubleCellStyle;
            if (order.amountPostupilo > 0) {
                cell.SetCellValue(order.amountPostupilo);                
            } else {
                cell.SetCellValue(string.Empty);
            }

            // расход.
            cell = row.CreateCell(4);
            cell.CellStyle = doubleCellStyle;
            if (order.amountSpisano > 0) {
                cell.SetCellValue(order.amountSpisano);
            } else {
                cell.SetCellValue(string.Empty);                
            }

            // контрагент.
            cell = row.CreateCell(5);
            cell.CellStyle = stringCellStyle;            
            cell.SetCellValue(order.ContractorName);

            // инн
            cell = row.CreateCell(6);
            cell.CellStyle = stringCellStyle;            
            cell.SetCellValue(string.IsNullOrEmpty(order.ContractorINN) ? string.Empty : order.ContractorINN);

            // ставка ндс.
            cell = row.CreateCell(7);
            cell.CellStyle = ndsCellStyle;
            if (order.Nds == -1) {
                cell.SetCellValue("Без НДС");
            } else if (order.Nds == -2) {
                cell.SetCellValue(string.Empty);
            } else {
                cell.SetCellValue(order.Nds + "");
            }

            // сумма ндс.            
            cell = row.CreateCell(8);
            cell.CellStyle = doubleCellStyle;
            if (order.Nds < 0) {
                cell.SetCellValue(string.Empty);
            } else {                
                cell.SetCellValue(order.NdsSum);
            }

            // назначение платежа.
            cell = row.CreateCell(9);
            cell.CellStyle = stringCellStyle;            
            cell.SetCellValue(order.PayDestination);

            // П.п
            cell = row.CreateCell(10);
            cell.CellStyle = stringCellStyle;
            cell.SetCellValue(order.Number);

            //КПП
            cell = row.CreateCell(11);
            cell.CellStyle = stringCellStyle;
            if (string.IsNullOrEmpty(order.ContractorKPP)) {
                cell.SetCellType(CellType.Blank);
            } else {
                cell.SetCellValue(order.ContractorKPP);
            }            

            // р счет контрагента.
            cell = row.CreateCell(12);
            cell.CellStyle = stringCellStyle;            
            cell.SetCellValue(order.ContractorAccount);

            // банк контрагента.
            cell = row.CreateCell(13);
            cell.CellStyle = stringCellStyle;
            cell.SetCellType(CellType.String);
            cell.SetCellValue(order.ContractorBank);

            return true;
        }

        private static int FindBestPosition(ISheet shield, OrderEntity order) {
            var currentRow = 0;
            for (currentRow = shield.LastRowNum; currentRow > 3; --currentRow) {
                var row = shield.GetRow(currentRow);
                var curData = row.GetCell(2).DateCellValue;
                if (curData.Date < order.date.Date) {
                    return currentRow + 1;
                }
                if (curData.Date == order.date) {
                    if (CheckEquivalent(row, order)) {
                        return -1;
                    }
                }
            }

            return currentRow + 1;
        }

        private static bool CheckEquivalent(IRow row, OrderEntity order) {
            var contractor = row.GetCell(5).StringCellValue;
            var destination = row.GetCell(9).StringCellValue;
            var number = (int)row.GetCell(10).NumericCellValue;

            return contractor.Equals(order.ContractorName) && destination.Equals(order.PayDestination) && number == order.Number;
        }

        private static IRow InsertRow(ISheet shield, int destinationRowNum) {
            var newRow = shield.GetRow(destinationRowNum);
            if (newRow != null) {
                shield.ShiftRows(destinationRowNum, shield.LastRowNum, 1);
            } else {
                newRow = shield.CreateRow(destinationRowNum);
            }

            return newRow;
        }

        private static void ClearAllColors(HSSFWorkbook wb) {
            var defaultFont = wb.CreateFont();
            defaultFont.FontName = "Calibri";
            defaultFont.Color = HSSFColor.Black.Index;
            defaultFont.FontHeightInPoints = 11;

            var shieldsCount = wb.NumberOfSheets;
            for (var i = 0; i < shieldsCount; ++i) {
                var shield = wb.GetSheetAt(i);
                var lastRowNumber = shield.LastRowNum;
                for (var j = 4; j <= lastRowNumber; ++j) {
                    var row = shield.GetRow(j);
                    if (row.Cells != null) {
                        foreach (var cell in row.Cells) {
                            cell.CellStyle.SetFont(defaultFont);
                        }
                    }
                }
            }
        }        
    }
}
