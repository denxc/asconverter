using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ASConverter {
    public class ASExporter {
        private static string[] exelColumns = { "Номер", "Дата", "Поступление", "Расход", "Контраген", "Ставка НДС", "Сумма НДС", "ИНН контрагента", "КПП Контрагента", "Р.счет контрагента", "Банк", "Назначение платежа" };
        private static string DEFAULT_REPORT = "Консолидированный отчет";

        private ExcelWorksheet lastList = null;

        public int Export(string aSourceFile, string aDestFile, out string aMessage) {
            var startAmount = 0.0;
            var endAmount = 0.0;
            var orders = new OrdersProvider().LoadOrders(aSourceFile, out startAmount, out endAmount);

            var xlsFile = new FileInfo(aDestFile);
            var pck = new ExcelPackage(xlsFile);

            ClearAllColors(pck);

            var spisano = 0.0;
            var postupilo = 0.0;

            var count = 0;
            for (var i = orders.Count - 1; i >= 0; --i) {
                var order = orders[i];
                if (AddOrderToFile(order, pck, startAmount)) {
                    count++;
                    spisano += order.amountSpisano;
                    postupilo += order.amountPostupilo;
                } else {
                    order.WasAdded = true;
                }
            }

            AddOrdersToDefaultList(orders, pck);

            aMessage = string.Empty;
            if (lastList != null) {
                CheckAmountResult(endAmount, out aMessage);
            }

            pck.Save();
            return count;
        }

        private void CheckAmountResult(double endAmount, out string aMessage) {
            var lastRowIndex = lastList.Dimension.End.Row;
            var startAmount = Convert.ToDouble(lastList.Cells[lastRowIndex, 1].Value);
            lastList.Cells[lastRowIndex, 3].Calculate();
            lastList.Cells[lastRowIndex, 4].Calculate();
            var postupilo = Convert.ToDouble(lastList.Cells[lastRowIndex, 3].Value);
            var spisano = Convert.ToDouble(lastList.Cells[lastRowIndex, 4].Value);

            aMessage = string.Empty;
            var expected = startAmount + postupilo - spisano;            
            if (Math.Abs(expected - endAmount) > 0.0001) {
                aMessage = string.Format("Возможно, за некоторый период выписки не импортированы.\nОжидался остаток: {0:0.##}, остаток в отчете: {1:0.##}", endAmount, expected);
            }
        }

        private void AddOrdersToDefaultList(List<OrderEntity> orders, ExcelPackage pck) {
            var lists = pck.Workbook.Worksheets;
            ExcelWorksheet list = null;
            for (var i = 1; i <= lists.Count; ++i) {
                if (lists[i].Name.Equals(DEFAULT_REPORT)) {
                    list = lists[i];
                    break;
                }
            }

            if (list == null) {
                list = lists.Add(DEFAULT_REPORT);
            }

            for (var i = orders.Count - 1; i >= 0; --i) { 
                AddOrderToList(list, orders[i], false, 0.0);
            }            
        }

        private void ClearAllColors(ExcelPackage pck) {
            var lists = pck.Workbook.Worksheets;
            for (var i = 1; i <= lists.Count; ++i) {
                var list = lists[i];
                if (list.Dimension != null) {
                    for (var j = 1; j <= list.Dimension.Rows; ++j) {
                        list.Row(j).Style.Font.Color.SetColor(Color.Black);
                    }
                }                
            }
        }

        private bool AddOrderToFile(OrderEntity order, ExcelPackage pck, double aStartAmount) {
            var lists = pck.Workbook.Worksheets;
            for (var i = 1; i <= lists.Count; ++i) {                
                if (order.OwnerBank.ToLower().StartsWith(lists[i].Name.ToLower()) ||
                    lists[i].Name.ToLower().StartsWith(order.OwnerBank.ToLower())) {
                    lastList = lists[i];
                    return AddOrderToList(lists[i], order, true, 0.0);
                }
            }

            lastList = lists.Add(order.OwnerBank);            
            return AddOrderToList(lastList, order, true, aStartAmount);
        }

        private bool AddOrderToList(ExcelWorksheet exelList, OrderEntity order, bool isCheckDublicats, double aStartAmount) {
            if (!CheckPreparation(exelList)) {
                PrepareExelList(exelList, aStartAmount);
            }

            if (isCheckDublicats && CheckExisting(exelList, order)) {
                return false;
            }

            if (order.WasAdded) {
                return false;
            }
            
            var rowIndex = FindLastRowIndex(exelList);
            exelList.InsertRow(rowIndex, 1);
            exelList.Row(rowIndex).Style.Font.Color.SetColor(Color.Green);
            var colIndex = 1;
            exelList.Cells[rowIndex, colIndex++].Value = order.Number;
            exelList.Cells[rowIndex, colIndex].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
            exelList.Cells[rowIndex, colIndex++].Value = order.date.Date;            
            exelList.Cells[rowIndex, colIndex++].Value = order.amountPostupilo > 0 ? order.amountPostupilo : 0;            
            exelList.Cells[rowIndex, colIndex++].Value = order.amountSpisano > 0 ? order.amountSpisano :  0;
            exelList.Cells[rowIndex, colIndex++].Value = order.ContractorName;
            if (order.Nds < 0) {
                exelList.Cells[rowIndex, colIndex++].Value = "Уточнить";
                exelList.Cells[rowIndex, colIndex++].Value = "";
            } else {
                exelList.Cells[rowIndex, colIndex++].Value = order.Nds;
                exelList.Cells[rowIndex, colIndex++].Value = order.NdsSum;
            }             
            exelList.Cells[rowIndex, colIndex++].Value = string.IsNullOrEmpty(order.ContractorINN) ? string.Empty : order.ContractorINN;
            exelList.Cells[rowIndex, colIndex++].Value = string.IsNullOrEmpty(order.ContractorKPP) ? string.Empty : order.ContractorKPP;
            exelList.Cells[rowIndex, colIndex++].Value = order.ContractorAccount;
            exelList.Cells[rowIndex, colIndex++].Value = order.ContractorBank;
            exelList.Cells[rowIndex, colIndex++].Value = order.PayDestination;
            for (var i = 0; i < exelColumns.Length; ++i) {
                exelList.Column(i + 1).AutoFit();
            }

            exelList.Cells[rowIndex + 1, 3].Formula = string.Format("SUM(C2:C{0})", rowIndex);
            exelList.Cells[rowIndex + 1, 4].Formula = string.Format("SUM(D2:D{0})", rowIndex);

            return true;
        }

        private bool CheckExisting(ExcelWorksheet exelList, OrderEntity order) {
            for (var i = 1; i <= exelList.Dimension.Rows; ++i) {
                var torder = new OrderEntity();                
                int.TryParse(exelList.Cells[i, 1].Value + "", out torder.Number);                
                torder.PayDestination = exelList.Cells[i, 12].Value + "";                

                if (torder.Number == order.Number &&                    
                    torder.PayDestination == order.PayDestination) {
                    return true;
                }
            }

            return false;
        }

        private int FindLastRowIndex(ExcelWorksheet exelList) {
            return exelList.Dimension.End.Row;
        }

        private void PrepareExelList(ExcelWorksheet exelList, double aStartAmount) {
            for (var i = 0; i < exelColumns.Length; ++i) {                
                exelList.Cells[1, i + 1].Value = exelColumns[i];
                var column = exelList.Column(i + 1);
                column.AutoFit();                
            }
            exelList.Cells[2, 1].Value = aStartAmount;
            exelList.Cells[2, 1].Style.Font.Color.SetColor(Color.White);
            exelList.Cells[2, 4].Value = "Итого";
            exelList.Cells[2, 3].Value = "Итого";            
        }

        private bool CheckPreparation(ExcelWorksheet exelList) {
            foreach (var row in exelColumns) {
                var isFind = false;
                for (var i = 0; i < exelColumns.Length + 5; ++i) {
                    if (row.Equals(exelList.Cells[1, i + 1].Value)) {
                        isFind = true;
                        break;
                    }
                }
                if (!isFind) {
                    return false;
                }
            }

            return true;
        }

        internal void Generate(string fileName)
        {
            var xlsFile = new FileInfo(fileName);
            var pck = new ExcelPackage(xlsFile);
            pck.Workbook.Worksheets.Add("Лист 1");
            pck.Save();
        }
    }
}
