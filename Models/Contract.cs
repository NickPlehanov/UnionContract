using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;

namespace UnionContractWF.Models {
    public class Contract {
        public string Number { get; set; }
        public string Date { get; set; }
        public string ClientName { get; set; }
        public string ClientSmallName { get; set; }
        public string ClientInfo { get; set; }
        public string ObjectName { get; set; }
        public string ObjectAddress { get; set; }
        public string ObjectCost { get; set; }
        public string ObjectTypeService { get; set; }//типы охранных услуг
        public string ObjectMonthlyPay { get; set; }
        public string OwningUser { get; set; }
        public string BlockTypeDistinct { get; set; }
        public string BlockTypeFull { get; set; }
        public string DeviceName { get; set; }
        public string DeviceCount { get; set; }
        public string DeviceSum { get; set; }
        public string AllCount { get; set; }
        public string AllSum { get; set; }
        public Guid? ExecutorID { get; set; }
        public List<ExecutorsInfo> Coexecutors { get; set; }
        public string SendActs { get; set; }
        public string PoryadokActs { get; set; }
        public string ObjectSignalization { get; set; }

        public string os { get; set; }
        public string ps { get; set; }
        public string trs { get; set; } 
        public string service_security { get; set; }

        private string _DateFull {
            get {
                if (string.IsNullOrEmpty(Date))
                    return null;
                else {
                    DateTime dt = DateTime.Parse(Date);
                    return dt.Day + " " + GetMonthName(dt.Month) + " " + dt.Year+"г."; 
                }
            }
            set => value = string.IsNullOrEmpty(value) ? null : value;
        }
        public string DateFull {
            get => _DateFull;
            set {
                _DateFull = value;
            }
        }

        private string GetMonthName(int number) {
            switch (number) {
                case 1:return "января";
                case 2:return "февраля";
                case 3:return "марта";
                case 4:return "апреля";
                case 5:return "мая";
                case 6:return "июня";
                case 7:return "июля";
                case 8:return "августа";
                case 9:return "сентября";
                case 10:return "октября";
                case 11:return "ноября";
                case 12:return "декабря";
                default:return null;
            }
        }
    }
}
