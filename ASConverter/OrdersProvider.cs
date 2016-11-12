using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ASConverter {
    public class OrdersProvider {

        private static string START_ORDER = "СекцияДокумент=";
        private static string END_ORDER = "КонецДокумента";
        private static string AMOUNT = "Сумма=";

        private static string DATA_SPISANO = "ДатаСписано";
        private static string DATA_POSTUPILO = "ДатаПоступило";
        private static string PLATELSHIK = "Плательщик=";
        private static string PLATELSHIK1 = "Плательщик1=";
        private static string PLATEL_SCHET = "ПлательщикРасчСчет";
        private static string PLATEL_INN = "ПлательщикИНН";
        private static string PLATEL_KPP = "ПлательщикКПП";
        private static string PLATEL_BANK = "ПлательщикБанк";
        private static string POLUCHATEL = "Получатель=";
        private static string POLUCHATEL1 = "Получатель1=";
        private static string POLUCHAT_SCHET = "ПолучательРасчСчет";
        private static string POLUCHAT_INN = "ПолучательИНН";
        private static string POLUCHAT_KPP = "ПолучательКПП";
        private static string POLUCHAT_BANK = "ПолучательБанк";
        private static string NUMBER = "Номер";

        private static string DESTINATION = "НазначениеПлатежа";

        private static string NO_NDS = "безндс";
        private static string NO_NDS1 = "ндснеоблагается";
        private static string NO_NDS2 = "безналога";

        private static string START_AMOUNT = "НачальныйОстаток";
        private static string END_AMOUNT = "КонечныйОстаток";

        public static List<OrderEntity> LoadOrders(string aFileName, out double aStartAmount, out double aEndAmount) {
            var orders = new List<OrderEntity>();            
            var lines = File.ReadAllLines(aFileName, Encoding.GetEncoding("Windows-1251"));
            aStartAmount = 0;
            aEndAmount = 0;
            var findStartAmount = false;
            for (var i = 0; i < lines.Length; ++i) {
                if (lines[i].StartsWith(END_AMOUNT)) {
                    aEndAmount = Convert.ToDouble(lines[i].Substring(lines[i].LastIndexOf('=') + 1).Replace('.', ','));
                }
                if (!findStartAmount && lines[i].StartsWith(START_AMOUNT)) {
                    aStartAmount = Convert.ToDouble(lines[i].Substring(lines[i].LastIndexOf('=') + 1).Replace('.', ','));
                    findStartAmount = true;
                }

                if (lines[i].StartsWith(START_ORDER)) {
                    var order = LoadOrder(lines, i + 1);                    
                    orders.Add(order);                    
                    Application.DoEvents();
                }
            }            

            return orders;
        }

        private static OrderEntity LoadOrder(string[] lines, int startIndex) {
            var order = new OrderEntity();
            order.WasAdded = false;
            var amount = FindLine(lines, AMOUNT, startIndex).Replace('.', ',');
            var dataSpisano = FindLine(lines, DATA_SPISANO, startIndex);
            var dataPostupilo = FindLine(lines, DATA_POSTUPILO, startIndex);

            var number = FindLine(lines, NUMBER, startIndex);

            var platelshik = FindLine(lines, PLATELSHIK, startIndex);
            var platelshik1 = FindLine(lines, PLATELSHIK1, startIndex);
            var platelSchet = FindLine(lines, PLATEL_SCHET, startIndex);
            var platelInn = FindLine(lines, PLATEL_INN, startIndex);
            var platelKpp = FindLine(lines, PLATEL_KPP, startIndex);
            var platelBank = FindLine(lines, PLATEL_BANK, startIndex);

            var poluchatel = FindLine(lines, POLUCHATEL, startIndex);
            var poluchatel1 = FindLine(lines, POLUCHATEL1, startIndex);
            var poluchatelSchet = FindLine(lines, POLUCHAT_SCHET, startIndex);
            var poluchatelInn = FindLine(lines, POLUCHAT_INN, startIndex);
            var poluchatelKpp = FindLine(lines, POLUCHAT_KPP, startIndex);
            var poluchatelBank = FindLine(lines, POLUCHAT_BANK, startIndex);

            var destination = FindLine(lines, DESTINATION, startIndex);
            order.PayDestination = destination.Substring(destination.IndexOf('=') + 1);

            order.Number = Convert.ToInt32(number.Substring(number.LastIndexOf('=') + 1).Trim());

            var tdestination = destination.ToLower().Replace(" ", string.Empty);
            var amountValue = Convert.ToDouble(amount.Substring(amount.LastIndexOf('=') + 1));
            if (tdestination.Contains(NO_NDS) || tdestination.Contains(NO_NDS1) || tdestination.Contains(NO_NDS2)) {
                order.Nds = -1;                
            } else {
                if (tdestination.Contains("18%") || tdestination.Contains("18.00%") || tdestination.Contains("18.0%") || destination.ToLower().Contains("ндс 18 ")) {
                    order.Nds = 18;
                } else if (tdestination.Contains("10%") || tdestination.Contains("10.0%") || tdestination.Contains("10.00%") || destination.ToLower().Contains("ндс 10 ")) {
                    order.Nds = 10;
                } else if (tdestination.Contains("0%") || destination.ToLower().Contains("ндс 0 ")) {
                    order.Nds = 0;
                } else {
                    order.Nds = -2;
                }
            }

            if (order.Nds != -1) {
                order.NdsSum = order.Nds * ((1.0 * amountValue) / (100.0 + order.Nds));
                order.NdsSum = Math.Round(order.NdsSum, 2);
            }            

            var dataSpisanoStr = dataSpisano?.Substring(dataSpisano.LastIndexOf('=') + 1).Trim();
            if (string.IsNullOrEmpty(dataSpisanoStr)) {
                order.amountPostupilo = amountValue;
                order.amountSpisano = 0;
                order.date = DateTime.Parse(dataPostupilo.Substring(dataPostupilo.LastIndexOf('=') + 1));
                order.OwnerName = poluchatel?.Substring(poluchatel.LastIndexOf('=') + 1);
                if (string.IsNullOrEmpty(order.OwnerName)) {
                    order.OwnerName = poluchatel1.Substring(poluchatel1.LastIndexOf('=') + 1);
                }
                order.OwnerBank = poluchatelBank.Substring(poluchatelBank.LastIndexOf('=') + 1);
                order.OwnerInn = poluchatelInn?.Substring(poluchatelInn.LastIndexOf('=') + 1);
                order.OwnerKPP = poluchatelKpp?.Substring(poluchatelKpp.LastIndexOf('=') + 1);
                order.OwnerAccount = poluchatelSchet.Substring(poluchatelSchet.LastIndexOf('=') + 1);
                order.ContractorName = platelshik?.Substring(platelshik.LastIndexOf('=') + 1);                
                if (string.IsNullOrEmpty(order.ContractorName)) {
                    order.ContractorName = platelshik1.Substring(platelshik1.LastIndexOf('=') + 1);
                }
                order.ContractorKPP = platelKpp?.Substring(platelKpp.LastIndexOf('=') + 1);
                order.ContractorINN = platelInn?.Substring(platelInn.LastIndexOf('=') + 1);
                order.ContractorBank = platelBank.Substring(platelBank.LastIndexOf('=') + 1);
                order.ContractorAccount = platelSchet.Substring(platelSchet.LastIndexOf('=') + 1);
            } else {
                order.amountPostupilo = 0;
                order.amountSpisano = amountValue;
                order.date = DateTime.Parse(dataSpisano.Substring(dataSpisano.LastIndexOf('=') + 1));
                order.OwnerName = platelshik?.Substring(platelshik.LastIndexOf('=') + 1);
                if (string.IsNullOrEmpty(order.OwnerName)) {
                    order.OwnerName = platelshik1.Substring(platelshik1.LastIndexOf('=') + 1);
                }
                order.OwnerBank = platelBank.Substring(platelBank.LastIndexOf('=') + 1);
                order.OwnerInn = platelInn?.Substring(platelInn.LastIndexOf('=') + 1);
                order.OwnerKPP = platelKpp?.Substring(platelKpp.LastIndexOf('=') + 1);
                order.OwnerAccount = platelSchet.Substring(platelSchet.LastIndexOf('=') + 1);
                order.ContractorName = poluchatel?.Substring(poluchatel.LastIndexOf('=') + 1);
                if (string.IsNullOrEmpty(order.ContractorName)) {
                    order.ContractorName = poluchatel1.Substring(poluchatel1.LastIndexOf('=') + 1);
                }
                order.ContractorKPP = poluchatelKpp?.Substring(poluchatelKpp.LastIndexOf('=') + 1);
                order.ContractorINN = poluchatelInn?.Substring(poluchatelInn.LastIndexOf('=') + 1);
                order.ContractorBank = poluchatelBank.Substring(poluchatelBank.LastIndexOf('=') + 1);
                order.ContractorAccount = poluchatelSchet.Substring(poluchatelSchet.LastIndexOf('=') + 1);
            }

            order.OwnerInn = order.OwnerInn?.Replace("ИНН " + order.OwnerInn, string.Empty);
            order.OwnerInn = order.ContractorINN?.Replace("ИНН " + order.ContractorINN, string.Empty);
            order.OwnerInn = order.OwnerInn?.Replace(order.OwnerInn, string.Empty);
            order.OwnerInn = order.ContractorINN?.Replace(order.ContractorINN, string.Empty);

            return order;                                 
        }

        private static string FindLine(string[] lines, string startStr, int startIndex) {
            for (var i = startIndex; i < lines.Length; ++i) {
                var line = lines[i];
                if (line.StartsWith(END_ORDER)) {
                    break;
                }
                if (line.StartsWith(startStr)) {
                    return line;
                }
            }

            return null;
        }
    }
}
