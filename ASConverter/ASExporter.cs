using NPOI.HSSF.Model;
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

        private static string[] defaultColumns = { "Клиент", "Книга продаж", "Дата", "Приход", "Расход", "Контраген",
                                                "ИНН", "Ставка НДС", "Сумма НДС", "Назначение платежа", "П/п", "КПП",
                                                "Р/счет контрагента", "Банк контрагента" };
        
        private const int DEFAULT_ROW_HEIGHT = (int)(0.5 * rowHeigthMultiplicator);

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

            var sheet = wb.CreateSheet(shieldName);
            PrepareShieldTop(sheet, aOrder, aStartAmount);

            var defaultReportShield = wb.GetSheet(DEFAULT_REPORT);
            if (defaultReportShield == null) {
                defaultReportShield = wb.CreateSheet(DEFAULT_REPORT);
                PrepareShieldTop(defaultReportShield, aOrder, aStartAmount);
            } else {
                var row = defaultReportShield.GetRow(0);
                var cell = row.GetCell(3);
                var cellValue = cell.NumericCellValue;
                cell.SetCellValue(cellValue + aStartAmount);
            }

            var emptyShield = wb.GetSheet(START_SHIELD_NAME);
            if (emptyShield != null) {                
                wb.SetSheetHidden(wb.GetSheetIndex(emptyShield), SheetState.VeryHidden);
            }

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

            var doubleCellStyle = sheet.Workbook.CreateCellStyle();
            doubleCellStyle.DataFormat = sheet.Workbook.CreateDataFormat().GetFormat("#0.00");
            doubleCellStyle.BorderBottom = BorderStyle.Medium;
            doubleCellStyle.BorderTop = BorderStyle.Medium;
            doubleCellStyle.BorderLeft = BorderStyle.Medium;
            doubleCellStyle.BorderRight = BorderStyle.Medium;
            doubleCellStyle.SetFont(boldFont);

            var stringCellStyle = sheet.Workbook.CreateCellStyle();            
            stringCellStyle.BorderBottom = BorderStyle.Medium;
            stringCellStyle.BorderTop = BorderStyle.Medium;
            stringCellStyle.BorderLeft = BorderStyle.Medium;
            stringCellStyle.BorderRight = BorderStyle.Medium;
            stringCellStyle.Alignment = HorizontalAlignment.Center;
            stringCellStyle.SetFont(boldFont);

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
            cell.SetCellValue(string.Empty);
            cell = row.CreateCell(2, CellType.String);
            cell.CellStyle = unicCellStyle;
            cell.SetCellValue(string.Empty);

            cell = row.CreateCell(3, CellType.Numeric);            
            cell.CellStyle = doubleCellStyle;            
            cell.SetCellValue(aStartAmount);

            row = sheet.CreateRow(1);
            row.Height = DEFAULT_ROW_HEIGHT;
            cell = row.CreateCell(0, CellType.String);
            cell.CellStyle = unicCellStyle;            
            cell.SetCellValue("Обороты за период");
            cell = row.CreateCell(1, CellType.String);
            cell.CellStyle = unicCellStyle;
            cell.SetCellValue(string.Empty);
            cell = row.CreateCell(2, CellType.String);
            cell.CellStyle = unicCellStyle;
            cell.SetCellValue(string.Empty);
            cell = row.CreateCell(3, CellType.Numeric);
            cell.CellStyle = doubleCellStyle;
            cell.SetCellValue(0.0);
            cell = row.CreateCell(4, CellType.Numeric);
            cell.CellStyle = doubleCellStyle;
            cell.SetCellValue(0.0);
            cell = row.CreateCell(5, CellType.String);
            cell.CellStyle = stringCellStyle;            
            cell.SetCellValue(aOrder.OwnerName);
            cell = row.CreateCell(6, CellType.String);
            cell.CellStyle = stringCellStyle;
            cell.SetCellValue(string.IsNullOrEmpty(aOrder.OwnerInn) ? string.Empty : aOrder.OwnerInn);
            cell = row.CreateCell(11, CellType.String);
            cell.CellStyle = stringCellStyle;
            cell.SetCellValue(string.IsNullOrEmpty(aOrder.OwnerKPP) ? string.Empty : aOrder.OwnerKPP);
            cell = row.CreateCell(12, CellType.String);
            cell.CellStyle = stringCellStyle;
            cell.SetCellValue(string.IsNullOrEmpty(aOrder.OwnerAccount) ? string.Empty : aOrder.OwnerAccount);            
            cell = row.CreateCell(13, CellType.String);
            cell.CellStyle = stringCellStyle;
            cell.SetCellValue(aOrder.OwnerBank);

            row = sheet.CreateRow(2);
            row.Height = DEFAULT_ROW_HEIGHT;
            cell = row.CreateCell(0, CellType.String);
            cell.SetCellValue("Остаток на конец периода");
            cell = row.CreateCell(3, CellType.Numeric);
            cell.CellStyle = doubleCellStyle;
            cell.SetCellFormula("D1+D2-E2");

            row = sheet.CreateRow(3);
            row.Height = DEFAULT_ROW_HEIGHT;
            for (var i = 0; i < defaultColumns.Length; ++i) {
                cell = row.CreateCell(i);
                cell.CellStyle = stringCellStyle;
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
            sheet.SetColumnWidth(13, (int)(8.75 * rowWidthMultiplicator));

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

        private static void ColorEndAmound(ISheet shield, bool isGreen) {
            var defaultFont = shield.Workbook.CreateFont();
            defaultFont.FontName = "Calibri";
            defaultFont.Color = isGreen ? HSSFColor.Red.Index : HSSFColor.Black.Index;
            defaultFont.IsBold = true;
            defaultFont.FontHeightInPoints = 11;

            var doubleCellStyle = shield.Workbook.CreateCellStyle();
            doubleCellStyle.DataFormat = shield.Workbook.CreateDataFormat().GetFormat("#0.00");
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
            doubleCellStyle.DataFormat = shield.Workbook.CreateDataFormat().GetFormat("#0.00");
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
            cell.SetCellValue(order.date);

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
            cell.SetCellValue(string.IsNullOrEmpty(order.ContractorKPP) ? string.Empty : order.ContractorKPP);

            // р счет контрагента.
            cell = row.CreateCell(12);
            cell.CellStyle = stringCellStyle;
            cell.SetCellValue(order.ContractorAccount);

            // банк контрагента.
            cell = row.CreateCell(13);
            cell.CellStyle = stringCellStyle;
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

        //private static string DEFAULT_REPORT = "Консолидированный отчет";

        //private ExcelWorksheet lastList = null;

        //public int Export(string aSourceFile, string aDestFile, out string aMessage) {
        //    var startAmount = 0.0;
        //    var endAmount = 0.0;
        //    var orders = new OrdersProvider().LoadOrders(aSourceFile, out startAmount, out endAmount);

        //    var xlsFile = new FileInfo(aDestFile);
        //    var pck = new ExcelPackage(xlsFile);

        //    ClearAllColors(pck);

        //    var spisano = 0.0;
        //    var postupilo = 0.0;

        //    var count = 0;
        //    for (var i = orders.Count - 1; i >= 0; --i) {
        //        var order = orders[i];
        //        if (AddOrderToFile(order, pck, startAmount)) {
        //            count++;
        //            spisano += order.amountSpisano;
        //            postupilo += order.amountPostupilo;
        //        } else {
        //            order.WasAdded = true;
        //        }
        //    }

        //    AddOrdersToDefaultList(orders, pck);

        //    aMessage = string.Empty;
        //    if (lastList != null) {
        //        CheckAmountResult(endAmount, out aMessage);
        //    }

        //    pck.Save();
        //    return count;
        //}

        //private void CheckAmountResult(double endAmount, out string aMessage) {
        //    var lastRowIndex = lastList.Dimension.End.Row;
        //    var startAmount = Convert.ToDouble(lastList.Cells[lastRowIndex, 1].Value);
        //    lastList.Cells[lastRowIndex, 3].Calculate();
        //    lastList.Cells[lastRowIndex, 4].Calculate();
        //    var postupilo = Convert.ToDouble(lastList.Cells[lastRowIndex, 3].Value);
        //    var spisano = Convert.ToDouble(lastList.Cells[lastRowIndex, 4].Value);

        //    aMessage = string.Empty;
        //    var expected = startAmount + postupilo - spisano;            
        //    if (Math.Abs(expected - endAmount) > 0.0001) {
        //        aMessage = string.Format("Возможно, за некоторый период выписки не импортированы.\nОжидался остаток: {0:0.##}, остаток в отчете: {1:0.##}", endAmount, expected);
        //    }
        //}

        //private void AddOrdersToDefaultList(List<OrderEntity> orders, ExcelPackage pck) {
        //    var lists = pck.Workbook.Worksheets;
        //    ExcelWorksheet list = null;
        //    for (var i = 1; i <= lists.Count; ++i) {
        //        if (lists[i].Name.Equals(DEFAULT_REPORT)) {
        //            list = lists[i];
        //            break;
        //        }
        //    }

        //    if (list == null) {
        //        list = lists.Add(DEFAULT_REPORT);
        //    }

        //    for (var i = orders.Count - 1; i >= 0; --i) { 
        //        AddOrderToList(list, orders[i], false, 0.0);
        //    }            
        //}

        //private void ClearAllColors(ExcelPackage pck) {
        //    var lists = pck.Workbook.Worksheets;
        //    for (var i = 1; i <= lists.Count; ++i) {
        //        var list = lists[i];
        //        if (list.Dimension != null) {
        //            for (var j = 1; j <= list.Dimension.Rows; ++j) {
        //                list.Row(j).Style.Font.Color.SetColor(Color.Black);
        //            }
        //        }                
        //    }
        //}

        //private bool AddOrderToFile(OrderEntity order, ExcelPackage pck, double aStartAmount) {
        //    var lists = pck.Workbook.Worksheets;
        //    for (var i = 1; i <= lists.Count; ++i) {                
        //        if (order.OwnerBank.ToLower().StartsWith(lists[i].Name.ToLower()) ||
        //            lists[i].Name.ToLower().StartsWith(order.OwnerBank.ToLower())) {
        //            lastList = lists[i];
        //            return AddOrderToList(lists[i], order, true, 0.0);
        //        }
        //    }

        //    lastList = lists.Add(order.OwnerBank);            
        //    return AddOrderToList(lastList, order, true, aStartAmount);
        //}

        //private bool AddOrderToList(ExcelWorksheet exelList, OrderEntity order, bool isCheckDublicats, double aStartAmount) {
        //    if (!CheckPreparation(exelList)) {
        //        PrepareExelList(exelList, aStartAmount);
        //    }

        //    if (isCheckDublicats && CheckExisting(exelList, order)) {
        //        return false;
        //    }

        //    if (order.WasAdded) {
        //        return false;
        //    }

        //    var rowIndex = FindLastRowIndex(exelList);
        //    exelList.InsertRow(rowIndex, 1);
        //    exelList.Row(rowIndex).Style.Font.Color.SetColor(Color.Green);
        //    var colIndex = 1;
        //    exelList.Cells[rowIndex, colIndex++].Value = order.Number;
        //    exelList.Cells[rowIndex, colIndex].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
        //    exelList.Cells[rowIndex, colIndex++].Value = order.date.Date;            
        //    exelList.Cells[rowIndex, colIndex++].Value = order.amountPostupilo > 0 ? order.amountPostupilo : 0;            
        //    exelList.Cells[rowIndex, colIndex++].Value = order.amountSpisano > 0 ? order.amountSpisano :  0;
        //    exelList.Cells[rowIndex, colIndex++].Value = order.ContractorName;
        //    if (order.Nds < 0) {
        //        exelList.Cells[rowIndex, colIndex++].Value = "Уточнить";
        //        exelList.Cells[rowIndex, colIndex++].Value = "";
        //    } else {
        //        exelList.Cells[rowIndex, colIndex++].Value = order.Nds;
        //        exelList.Cells[rowIndex, colIndex++].Value = order.NdsSum;
        //    }             
        //    exelList.Cells[rowIndex, colIndex++].Value = string.IsNullOrEmpty(order.ContractorINN) ? string.Empty : order.ContractorINN;
        //    exelList.Cells[rowIndex, colIndex++].Value = string.IsNullOrEmpty(order.ContractorKPP) ? string.Empty : order.ContractorKPP;
        //    exelList.Cells[rowIndex, colIndex++].Value = order.ContractorAccount;
        //    exelList.Cells[rowIndex, colIndex++].Value = order.ContractorBank;
        //    exelList.Cells[rowIndex, colIndex++].Value = order.PayDestination;
        //    for (var i = 0; i < exelColumns.Length; ++i) {
        //        exelList.Column(i + 1).AutoFit();
        //    }

        //    exelList.Cells[rowIndex + 1, 3].Formula = string.Format("SUM(C2:C{0})", rowIndex);
        //    exelList.Cells[rowIndex + 1, 4].Formula = string.Format("SUM(D2:D{0})", rowIndex);

        //    return true;
        //}

        //private bool CheckExisting(ExcelWorksheet exelList, OrderEntity order) {
        //    for (var i = 1; i <= exelList.Dimension.Rows; ++i) {
        //        var torder = new OrderEntity();                
        //        int.TryParse(exelList.Cells[i, 1].Value + "", out torder.Number);                
        //        torder.PayDestination = exelList.Cells[i, 12].Value + "";                

        //        if (torder.Number == order.Number &&                    
        //            torder.PayDestination == order.PayDestination) {
        //            return true;
        //        }
        //    }

        //    return false;
        //}

        //private int FindLastRowIndex(ExcelWorksheet exelList) {
        //    return exelList.Dimension.End.Row;
        //}

        //private void PrepareExelList(ExcelWorksheet exelList, double aStartAmount) {
        //    for (var i = 0; i < exelColumns.Length; ++i) {                
        //        exelList.Cells[1, i + 1].Value = exelColumns[i];
        //        var column = exelList.Column(i + 1);
        //        column.AutoFit();                
        //    }
        //    exelList.Cells[2, 1].Value = aStartAmount;
        //    exelList.Cells[2, 1].Style.Font.Color.SetColor(Color.White);
        //    exelList.Cells[2, 4].Value = "Итого";
        //    exelList.Cells[2, 3].Value = "Итого";            
        //}

        //private bool CheckPreparation(ExcelWorksheet exelList) {
        //    foreach (var row in exelColumns) {
        //        var isFind = false;
        //        for (var i = 0; i < exelColumns.Length + 5; ++i) {
        //            if (row.Equals(exelList.Cells[1, i + 1].Value)) {
        //                isFind = true;
        //                break;
        //            }
        //        }
        //        if (!isFind) {
        //            return false;
        //        }
        //    }

        //    return true;
        //}        
    }
}
