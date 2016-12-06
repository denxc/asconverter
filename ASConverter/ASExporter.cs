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
        private const string START_SHIELD_NAME = "Empty1412";

        private static string[] defaultColumns = { "Клиент", "Книга продаж", "Дата", "Приход", "Расход", "Контрагент",
                                                "ИНН", "Ставка НДС", "Сумма НДС", "Назначение платежа", "П/п", "КПП",
                                                "Р/счет контрагента", "Банк контрагента" };

        private const string DOUBLE_FORMAT = "#,##0.00";
        private const string DATE_FORMAT = "DD.MM.YY;@";

        private const int DEFAULT_ROW_HEIGHT = (int)(0.5 * rowHeigthMultiplicator);

        private static Dictionary<int, ICellStyle> styles = new Dictionary<int, ICellStyle>();
        private static Dictionary<int, ICellStyle> greenStyles = new Dictionary<int, ICellStyle>();

        public static void Generate(string aXlsFile) {
            var wb = HSSFWorkbook.Create(InternalWorkbook.CreateWorkbook());
            var shield = (HSSFSheet)wb.CreateSheet(START_SHIELD_NAME);
            PrepareEmptyShield(shield);
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
                        System.Windows.Forms.Application.DoEvents();
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

        public static void CreateNewShield(string aXlsFile, string shieldName, OrderEntity aOrder, AccountSection aStartAccountSection, AccountSection aEndAccountSection) {
            HSSFWorkbook wb = null;
            using (var fs = new FileStream(aXlsFile, FileMode.Open, FileAccess.Read)) {
                wb = new HSSFWorkbook(fs);
                fs.Close();                       
            }
            
            var defaultReportShield = wb.GetSheet(DEFAULT_REPORT);
            if (defaultReportShield == null) {
                defaultReportShield = wb.CreateSheet(DEFAULT_REPORT);
                wb.SetSheetOrder(DEFAULT_REPORT, 0);
                PrepareShieldTop(defaultReportShield, aOrder, aStartAccountSection, aEndAccountSection);
            } else {
                var row = defaultReportShield.GetRow(0);
                var cell = row.GetCell(3);
                var cellValue = cell.NumericCellValue;
                cell.SetCellValue(cellValue + aStartAccountSection.StartAmount);                                
            }

            var defaultReportIndex = wb.GetSheetIndex(defaultReportShield);

            var sheet = wb.CreateSheet(shieldName);
            wb.SetSheetOrder(shieldName, defaultReportIndex);
            PrepareShieldTop(sheet, aOrder, aStartAccountSection, aEndAccountSection);            

            wb.SetActiveSheet(defaultReportIndex);

            using (var fs = new FileStream(aXlsFile, FileMode.Create, FileAccess.Write)) {
                wb.Write(fs);
                fs.Close();
            }
        }

        private static void PrepareShieldTop(ISheet sheet, OrderEntity aOrder, AccountSection aStartAccountSection, AccountSection aEndAccountSection) {
            var boldFont = sheet.Workbook.CreateFont();
            boldFont.FontName = "Calibri";
            boldFont.FontHeightInPoints = 11;            
            boldFont.IsBold = true;

            var defaultFont = sheet.Workbook.CreateFont();
            defaultFont.FontName = "Calibri";
            defaultFont.FontHeightInPoints = 11;

            var hiddenFont = sheet.Workbook.CreateFont();
            hiddenFont.FontName = "Calibri";
            hiddenFont.FontHeightInPoints = 11;
            hiddenFont.Color = HSSFColor.White.Index;

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

            var dataCellStyle = sheet.Workbook.CreateCellStyle();
            dataCellStyle.DataFormat = sheet.Workbook.CreateDataFormat().GetFormat(DATE_FORMAT);
            dataCellStyle.BorderBottom = BorderStyle.Medium;
            dataCellStyle.BorderTop = BorderStyle.Medium;
            dataCellStyle.BorderLeft = BorderStyle.Medium;            
            dataCellStyle.BorderRight = BorderStyle.Medium;
            dataCellStyle.SetFont(defaultFont);

            var unicCellStyle = sheet.Workbook.CreateCellStyle();
            unicCellStyle.BorderBottom = BorderStyle.Medium;
            unicCellStyle.SetFont(defaultFont);

            var doubleCellHiddenStyle = sheet.Workbook.CreateCellStyle();
            doubleCellHiddenStyle.DataFormat = sheet.Workbook.CreateDataFormat().GetFormat(DOUBLE_FORMAT);
            doubleCellHiddenStyle.SetFont(hiddenFont);

            var row = sheet.CreateRow(0);
            row.Height = DEFAULT_ROW_HEIGHT;
            var cell = row.CreateCell(0);                        
            cell.SetCellValue("Остаток на начало периода");
            cell.CellStyle = unicCellStyle;
            cell = row.CreateCell(1);
            cell.CellStyle = unicCellStyle;
            cell = row.CreateCell(2);
            cell.CellStyle = unicCellStyle;
            if (!sheet.SheetName.Equals(DEFAULT_REPORT)) {
                cell.CellStyle = dataCellStyle;
                cell.SetCellValue(aStartAccountSection.StartData.Date);
            }
            //sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, 2));

            cell = row.CreateCell(3, CellType.Numeric);                        
            cell.SetCellValue(aStartAccountSection.StartAmount);
            cell.CellStyle = doubleCellDefaultStyle;

            row = sheet.CreateRow(1);
            row.Height = DEFAULT_ROW_HEIGHT;
            cell = row.CreateCell(0, CellType.String);            
            cell.SetCellValue("Обороты за период");
            cell.CellStyle = unicCellStyle;
            cell = row.CreateCell(1);
            cell.CellStyle = unicCellStyle;
            cell = row.CreateCell(2);
            cell.CellStyle = unicCellStyle;
            //sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(1, 1, 0, 2));
            cell = row.CreateCell(3, CellType.Numeric);            
            cell.SetCellValue(0.0);
            cell.CellStyle = doubleCellDefaultStyle;
            cell = row.CreateCell(4, CellType.Numeric);            
            cell.SetCellValue(0.0);
            cell.CellStyle = doubleCellDefaultStyle;
            cell = row.CreateCell(5, CellType.String);                     
            cell.SetCellValue(aOrder.OwnerName);
            cell.CellStyle = stringCellBoldStyle;
            cell = row.CreateCell(6, CellType.String);
            cell.SetCellValue(string.IsNullOrEmpty(aOrder.OwnerInn) ? string.Empty : aOrder.OwnerInn);
            cell.CellStyle = stringCellDefaultStyle;
            cell = row.CreateCell(11, CellType.String);                        
            cell.SetCellValue(string.IsNullOrEmpty(aOrder.OwnerKPP) ? string.Empty : aOrder.OwnerKPP);
            cell.CellStyle = stringCellDefaultLeftStyle;
            cell = row.CreateCell(12, CellType.String);                        
            cell.SetCellValue(string.IsNullOrEmpty(aOrder.OwnerAccount) ? string.Empty : aOrder.OwnerAccount);
            cell.CellStyle = stringCellDefaultLeftStyle;
            cell = row.CreateCell(13, CellType.String);                        
            cell.SetCellValue(aOrder.OwnerBank);
            cell.CellStyle = stringCellDefaultLeftStyle;

            row = sheet.CreateRow(2);
            row.Height = DEFAULT_ROW_HEIGHT;
            cell = row.CreateCell(0);            
            cell.SetCellValue("Остаток на конец периода");
            cell = row.CreateCell(1);
            cell.CellStyle = unicCellStyle;
            cell = row.CreateCell(2);
            cell.CellStyle = unicCellStyle;
            if (!sheet.SheetName.Equals(DEFAULT_REPORT)) {
                cell.CellStyle = dataCellStyle;
                cell.SetCellValue(aEndAccountSection.EndData.Date);
            }            
            //sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(2, 2, 0, 2));
            cell = row.CreateCell(3, CellType.Numeric);            
            cell.SetCellFormula("D1+D2-E2");
            cell.CellStyle = doubleCellBoldStyle;            
            if (!sheet.SheetName.Equals(DEFAULT_REPORT)) {
                cell = row.CreateCell(4, CellType.Numeric);
                cell.CellStyle = doubleCellHiddenStyle;
                cell.SetCellValue(aEndAccountSection.EndAmount);
            }            

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
            sheet.SetColumnWidth(5, (int)(8.3 * rowWidthMultiplicator));
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

        private static void PrepareEmptyShield(ISheet shield) {
            var defaultFont = shield.Workbook.CreateFont();
            defaultFont.FontName = "Calibri";
            defaultFont.Color = HSSFColor.Black.Index;
            defaultFont.FontHeightInPoints = 11;
            var greenFont = shield.Workbook.CreateFont();
            greenFont.FontName = "Calibri";
            greenFont.Color = HSSFColor.Green.Index;
            greenFont.FontHeightInPoints = 11;

            var doubleCellStyle = shield.Workbook.CreateCellStyle();
            doubleCellStyle.DataFormat = shield.Workbook.CreateDataFormat().GetFormat(DOUBLE_FORMAT);
            doubleCellStyle.BorderRight = BorderStyle.Medium;
            doubleCellStyle.SetFont(defaultFont);
            var stringCellStyle = shield.Workbook.CreateCellStyle();
            stringCellStyle.BorderRight = BorderStyle.Medium;
            stringCellStyle.SetFont(defaultFont);
            var ndsCellStyle = shield.Workbook.CreateCellStyle();
            ndsCellStyle.BorderRight = BorderStyle.Medium;
            ndsCellStyle.Alignment = HorizontalAlignment.Center;
            ndsCellStyle.SetFont(defaultFont);
            var dataCellStyle = shield.Workbook.CreateCellStyle();
            dataCellStyle.DataFormat = shield.Workbook.CreateDataFormat().GetFormat(DATE_FORMAT);
            dataCellStyle.BorderRight = BorderStyle.Medium;
            dataCellStyle.SetFont(defaultFont);

            var doubleCellGreenStyle = shield.Workbook.CreateCellStyle();
            doubleCellGreenStyle.DataFormat = shield.Workbook.CreateDataFormat().GetFormat(DOUBLE_FORMAT);
            doubleCellGreenStyle.BorderRight = BorderStyle.Medium;
            doubleCellGreenStyle.SetFont(greenFont);
            var stringCellGreenStyle = shield.Workbook.CreateCellStyle();
            stringCellGreenStyle.BorderRight = BorderStyle.Medium;
            stringCellGreenStyle.SetFont(greenFont);
            var ndsCellGreenStyle = shield.Workbook.CreateCellStyle();
            ndsCellGreenStyle.BorderRight = BorderStyle.Medium;
            ndsCellGreenStyle.Alignment = HorizontalAlignment.Center;
            ndsCellGreenStyle.SetFont(greenFont);
            var dataCellGreenStyle = shield.Workbook.CreateCellStyle();
            dataCellGreenStyle.DataFormat = shield.Workbook.CreateDataFormat().GetFormat(DATE_FORMAT);
            dataCellGreenStyle.BorderRight = BorderStyle.Medium;
            dataCellGreenStyle.SetFont(greenFont);

            var row = shield.CreateRow(0);            
            row.CreateCell(0).CellStyle = stringCellStyle;
            row.CreateCell(1).CellStyle = stringCellStyle;
            row.CreateCell(2).CellStyle = dataCellStyle;
            row.CreateCell(3).CellStyle = doubleCellStyle;
            row.CreateCell(4).CellStyle = doubleCellStyle;
            row.CreateCell(5).CellStyle = stringCellStyle;
            row.CreateCell(6).CellStyle = stringCellStyle;
            row.CreateCell(7).CellStyle = ndsCellStyle;
            row.CreateCell(8).CellStyle = doubleCellStyle;
            row.CreateCell(9).CellStyle = stringCellStyle;
            row.CreateCell(10).CellStyle = ndsCellStyle;
            row.CreateCell(11).CellStyle = stringCellStyle;
            row.CreateCell(12).CellStyle = stringCellStyle;
            row.CreateCell(13).CellStyle = stringCellStyle;

            row = shield.CreateRow(1);
            row.CreateCell(0).CellStyle = stringCellGreenStyle;
            row.CreateCell(1).CellStyle = stringCellGreenStyle;
            row.CreateCell(2).CellStyle = dataCellGreenStyle;
            row.CreateCell(3).CellStyle = doubleCellGreenStyle;
            row.CreateCell(4).CellStyle = doubleCellGreenStyle;
            row.CreateCell(5).CellStyle = stringCellGreenStyle;
            row.CreateCell(6).CellStyle = stringCellGreenStyle;
            row.CreateCell(7).CellStyle = ndsCellGreenStyle;
            row.CreateCell(8).CellStyle = doubleCellGreenStyle;
            row.CreateCell(9).CellStyle = stringCellGreenStyle;
            row.CreateCell(10).CellStyle = ndsCellGreenStyle;
            row.CreateCell(11).CellStyle = stringCellGreenStyle;
            row.CreateCell(12).CellStyle = stringCellGreenStyle;
            row.CreateCell(13).CellStyle = stringCellGreenStyle;

            shield.Workbook.SetSheetHidden(shield.Workbook.GetSheetIndex(shield), SheetState.VeryHidden);
        }

        public static int Export(string destFilePath, string aShieldName, OrderEntity[] orders, AccountSection aStartAccountSection, AccountSection aEndAccountSection, out string message) {
            HSSFWorkbook wb = null;
            using (var fs = new FileStream(destFilePath, FileMode.Open, FileAccess.Read)) {
                wb = new HSSFWorkbook(fs);
                fs.Close();
            }            

            ClearAllColors(wb);

            var countAdded = 0;
            var shield = wb.GetSheet(aShieldName);

            LoadCellStyles(wb);            

            var defaultReport = wb.GetSheet(DEFAULT_REPORT);
            foreach (var order in orders) {
                if (AddOrdeToShield(shield, order)) {
                    AddOrdeToShield(defaultReport, order);
                    countAdded++;                    
                }
            }

            var topEndAmount = CheckEndAmount(shield, aEndAccountSection);
            CheckStartAmount(shield, defaultReport, aStartAccountSection);

            EvaluateEndAmount(defaultReport);
            var curEndAmount = EvaluateEndAmount(shield);
            
            if (Math.Abs(curEndAmount - topEndAmount) > 0.01) {
                message = string.Format("Возможно, за некоторый период выписки не импортированы.\nОжидался остаток: {0:0.##}, остаток в отчете: {1:0.##}", topEndAmount, curEndAmount);
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

        private static double CheckEndAmount(ISheet shield, AccountSection aEndAccountSection) {
            var dateCell = shield.GetRow(2).GetCell(2);
            var hiddenAmountCell = shield.GetRow(2).GetCell(4);

            var currentDate = dateCell.DateCellValue;

            if (currentDate.Date >= aEndAccountSection.EndData.Date) {
                return hiddenAmountCell.NumericCellValue;
            }

            dateCell.SetCellValue(aEndAccountSection.EndData.Date);            
            hiddenAmountCell.SetCellValue(aEndAccountSection.EndAmount);

            return aEndAccountSection.EndAmount;
        }

        private static void CheckStartAmount(ISheet shield, ISheet defaultReport, AccountSection aStartAccountSection) {
            var dateCell = shield.GetRow(0).GetCell(2);
            var currentDate = dateCell.DateCellValue;
            if (currentDate.Date < aStartAccountSection.StartData.Date) {
                return;
            }

            var amountCell = shield.GetRow(0).GetCell(3);
            var currentAmount = amountCell.NumericCellValue;

            amountCell.SetCellValue(aStartAccountSection.StartAmount);
            dateCell.SetCellValue(aStartAccountSection.StartData.Date);

            var defaultAmountCell = defaultReport.GetRow(0).GetCell(3);
            var defaultAmount = defaultAmountCell.NumericCellValue;
            defaultAmountCell.SetCellValue(defaultAmount - currentAmount + aStartAccountSection.StartAmount);
        }

        private static void LoadCellStyles(IWorkbook wb) {
            var shield = wb.GetSheet(START_SHIELD_NAME);
            var row = shield.GetRow(0);
            var greenRow = shield.GetRow(1);
            styles.Clear();
            greenStyles.Clear();
            for (var i = 0; i < 14; ++i) {
                styles[i] = row.GetCell(i).CellStyle;
                greenStyles[i] = greenRow.GetCell(i).CellStyle;
            }
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

        private static double EvaluateEndAmount(ISheet shield) {
            var row = shield.GetRow(1);
            var prihod = row.GetCell(3);
            var rashod = row.GetCell(4);            
            prihod.SetCellFormula(string.Format("SUM(D{0}: D{1})", 5, 1000 + shield.LastRowNum + 1));
            rashod.SetCellFormula(string.Format("SUM(E{0}: E{1})", 5, 1000 + shield.LastRowNum + 1));            

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

            var row = InsertRow(shield, position);            
            row.Height = DEFAULT_ROW_HEIGHT;            

            // клиент.
            var cell = row.CreateCell(0);            
            cell.SetCellValue(string.Empty);
            cell.CellStyle = greenStyles[0];

            // книга продаж.
            cell = row.CreateCell(1);            
            cell.SetCellValue(string.Empty);
            cell.CellStyle = greenStyles[1];

            // дата.
            cell = row.CreateCell(2);
            cell.SetCellValue(order.date.Date);            
            cell.CellStyle = greenStyles[2];            

            // Приход.
            cell = row.CreateCell(3);            
            if (order.amountPostupilo > 0) {
                cell.SetCellValue(order.amountPostupilo);                
            } else {
                //cell.SetCellValue(string.Empty);
            }
            cell.CellStyle = greenStyles[3];

            // расход.
            cell = row.CreateCell(4);            
            if (order.amountSpisano > 0) {
                cell.SetCellValue(order.amountSpisano);
            } else {
                //cell.SetCellValue(string.Empty);                
            }
            cell.CellStyle = greenStyles[4];

            // контрагент.
            cell = row.CreateCell(5);                        
            cell.SetCellValue(order.ContractorName);
            cell.CellStyle = greenStyles[5];

            // инн
            cell = row.CreateCell(6);                        
            cell.SetCellValue(string.IsNullOrEmpty(order.ContractorINN) ? string.Empty : order.ContractorINN);
            cell.CellStyle = greenStyles[6];

            // ставка ндс.
            cell = row.CreateCell(7);            
            if (order.Nds == -1) {
                cell.SetCellValue("Без НДС");
            } else if (order.Nds == -2) {
                cell.SetCellValue(string.Empty);
            } else {
                cell.SetCellValue(order.Nds + "");
            }
            cell.CellStyle = greenStyles[7];

            // сумма ндс.            
            cell = row.CreateCell(8);            
            if (order.Nds < 0) {
                //cell.SetCellValue(string.Empty);
            } else {                
                cell.SetCellValue(order.NdsSum);
            }
            cell.CellStyle = greenStyles[8];

            // назначение платежа.
            cell = row.CreateCell(9);            
            cell.SetCellValue(order.PayDestination);
            cell.CellStyle = greenStyles[9];

            // П.п
            cell = row.CreateCell(10);
            cell.SetCellValue(order.Number);
            cell.CellStyle = greenStyles[10];

            //КПП
            cell = row.CreateCell(11);            
            if (string.IsNullOrEmpty(order.ContractorKPP)) {
                //cell.SetCellType(CellType.Blank);
            } else {
                cell.SetCellValue(order.ContractorKPP);
            }
            cell.CellStyle = greenStyles[11];

            // р счет контрагента.
            cell = row.CreateCell(12);                        
            cell.SetCellValue(order.ContractorAccount);
            cell.CellStyle = greenStyles[12];

            // банк контрагента.
            cell = row.CreateCell(13);            
            cell.SetCellValue(order.ContractorBank);
            cell.CellStyle = greenStyles[13];

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
            var amountSpisano = GetNumericValue(row.GetCell(4));
            var amountPostupilo = GetNumericValue(row.GetCell(5));

            return contractor.Equals(order.ContractorName) && 
                destination.Equals(order.PayDestination) && 
                number == order.Number && 
                (amountSpisano == order.amountSpisano || 
                amountPostupilo == order.amountPostupilo);
        }

        private static double GetNumericValue(ICell cell) {
            double result;
            try {
                result = cell.NumericCellValue;
            } catch {
                result = double.MinValue;
            }

            return result;
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
            LoadCellStyles(wb);          
            var shieldsCount = wb.NumberOfSheets;
            for (var i = 0; i < shieldsCount; ++i) {               
                var shield = wb.GetSheetAt(i);
                if (shield.SheetName.Equals(START_SHIELD_NAME)) {
                    continue;
                }
                var lastRowNumber = shield.LastRowNum;
                for (var j = 4; j <= lastRowNumber; ++j) {
                    var row = shield.GetRow(j);
                    if (row.Cells != null) {
                        for (var k = 0; k < 14; ++k) {
                            row.GetCell(k).CellStyle = styles[k];
                        }
                    }
                }
            }
        }        
    }
}
