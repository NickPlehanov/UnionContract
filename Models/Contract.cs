using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;

namespace UnionContractWF.Models {
    public class Contract {
        public string Number { get; set; }
        public string Date { get; set; }
        public string ClientName { get; set; }
        public string ClientInfo { get; set; }
        public string ObjectName { get; set; }
        public string ObjectAddress { get; set; }
        public string ObjectCost { get; set; }
        public string ObjectTypeService { get; set; }//типы охранных услуг
        public string ObjectMonthlyPay { get; set; }
        public string OwningUser { get; set; }
        public string BlockType { get; set; }
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
    }
}
