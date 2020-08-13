using Microsoft.Office.Interop.Word;
using System;
using System.Diagnostics;
using System.Linq;
using UnionContractWF.Models;

namespace UnionContractWF.Helpers {
    public static class WordDocument {
        private static void TryGetItem(Document oDoc, object mark, string value) {
            try {
                oDoc.Bookmarks.get_Item(ref mark).Range.Text = value;
            }
            catch { }
        }
        public static void Exchange(Contract contract, string fromPathAndName, string toPathAndName) {
            try {
                var currentPath = AppDomain.CurrentDomain.BaseDirectory;

                var oWord = new Microsoft.Office.Interop.Word.Application();

                var oDoc = oWord.Documents.Add(fromPathAndName);
                TryGetItem(oDoc, Markers.contract_number, contract.Number);
                TryGetItem(oDoc, Markers.contract_number4, contract.Number);
                TryGetItem(oDoc, Markers.object_name, contract.ObjectName);
                TryGetItem(oDoc, Markers.date_day, contract.Date);
                TryGetItem(oDoc, Markers.date_day1, contract.Date);
                TryGetItem(oDoc, Markers.agent_type_header, contract.ClientName);
                TryGetItem(oDoc, Markers.agent_type_header1, contract.ClientName);
                TryGetItem(oDoc, Markers.object_address, contract.ObjectAddress);
                TryGetItem(oDoc, Markers.abonent_cost, contract.ObjectMonthlyPay);
                TryGetItem(oDoc, Markers.object_cost, contract.ObjectCost);
                TryGetItem(oDoc, Markers.object_name1, contract.ObjectName);
                TryGetItem(oDoc, Markers.object_name6, contract.ObjectName);
                TryGetItem(oDoc, Markers.object_address1, contract.ObjectAddress);
                TryGetItem(oDoc, Markers.contract_date3, contract.Date);
                TryGetItem(oDoc, Markers.object_address4, contract.ObjectAddress);
                TryGetItem(oDoc, Markers.object_address5, contract.ObjectAddress);
                TryGetItem(oDoc, Markers.object_name3, contract.ObjectName);
                TryGetItem(oDoc, Markers.fullData, contract.ClientInfo);
                TryGetItem(oDoc, Markers.fullDataFIO, contract.ClientName);
                TryGetItem(oDoc, Markers.owningUser, contract.OwningUser);
                TryGetItem(oDoc, Markers.services, contract.ObjectTypeService);
                TryGetItem(oDoc, Markers.blockType, contract.BlockType);
                TryGetItem(oDoc, Markers.rent_AllCount, contract.AllCount);
                TryGetItem(oDoc, Markers.rent_AllSum, contract.AllSum);
                TryGetItem(oDoc, Markers.rent_deviceName, contract.DeviceName);
                TryGetItem(oDoc, Markers.rent_deviceQty, contract.DeviceCount);
                TryGetItem(oDoc, Markers.rent_deviceSum, contract.DeviceSum);
                TryGetItem(oDoc, Markers.FIO3, contract.ClientName);
                TryGetItem(oDoc, Markers.ClientName1, contract.ClientName);
                TryGetItem(oDoc, Markers.SendAct, contract.SendActs);
                TryGetItem(oDoc, Markers.PoryadokAct, contract.PoryadokActs);
                TryGetItem(oDoc, Markers.object_signalizations, contract.ObjectSignalization);
                TryGetItem(oDoc, Markers.MainExecutorHeader, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseInfo);
                TryGetItem(oDoc, Markers.ExecutorFullInfo, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).FullInfo);
                TryGetItem(oDoc, Markers.ExecutorBossPosition, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition);
                TryGetItem(oDoc, Markers.MainExecutorBossPosition, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition);
                TryGetItem(oDoc, Markers.MainExecutorBossPosition1, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition);
                TryGetItem(oDoc, Markers.ExecutorBoss, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname);
                TryGetItem(oDoc, Markers.MainExecutorBossName, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname);
                TryGetItem(oDoc, Markers.MainExecutorBossName1, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname);
                TryGetItem(oDoc, Markers.ExecutorName1, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
                TryGetItem(oDoc, Markers.ExecutorName2, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
                TryGetItem(oDoc, Markers.ExecutorName3, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
                TryGetItem(oDoc, Markers.ExecutorName4, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
                TryGetItem(oDoc, Markers.MainExecutorName, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
                TryGetItem(oDoc, Markers.LicInfo, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Lic);
                TryGetItem(oDoc, Markers.LicDates, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicDates);
                TryGetItem(oDoc, Markers.LicCase, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicCase);
                TryGetItem(oDoc, Markers.Address, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Address);
                TryGetItem(oDoc, Markers.BossInfo, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossInfo);
                TryGetItem(oDoc, Markers.BossCompany, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossCompany);

                //oDoc.Bookmarks.get_Item(ref Markers.contract_number).Range.Text = contract.Number;
                //oDoc.Bookmarks.get_Item(ref Markers.contract_number4).Range.Text = contract.Number;
                //oDoc.Bookmarks.get_Item(ref Markers.object_name).Range.Text = contract.ObjectName;
                //oDoc.Bookmarks.get_Item(ref Markers.date_day).Range.Text = contract.Date;
                //oDoc.Bookmarks.get_Item(ref Markers.date_day1).Range.Text = contract.Date;
                //oDoc.Bookmarks.get_Item(ref Markers.agent_type_header).Range.Text = contract.ClientName;
                //oDoc.Bookmarks.get_Item(ref Markers.agent_type_header1).Range.Text = contract.ClientName;
                //oDoc.Bookmarks.get_Item(ref Markers.object_address).Range.Text = contract.ObjectAddress;
                //oDoc.Bookmarks.get_Item(ref Markers.abonent_cost).Range.Text = contract.ObjectMonthlyPay;
                //oDoc.Bookmarks.get_Item(ref Markers.object_cost).Range.Text = contract.ObjectCost;
                //oDoc.Bookmarks.get_Item(ref Markers.object_name1).Range.Text = contract.ObjectName;
                //oDoc.Bookmarks.get_Item(ref Markers.object_name6).Range.Text = contract.ObjectName;
                //oDoc.Bookmarks.get_Item(ref Markers.object_address1).Range.Text = contract.ObjectAddress;
                //oDoc.Bookmarks.get_Item(ref Markers.contract_date3).Range.Text = contract.Date;
                //oDoc.Bookmarks.get_Item(ref Markers.object_address4).Range.Text = contract.ObjectAddress;
                //oDoc.Bookmarks.get_Item(ref Markers.object_address5).Range.Text = contract.ObjectAddress;
                //oDoc.Bookmarks.get_Item(ref Markers.object_name3).Range.Text = contract.ObjectName;
                //oDoc.Bookmarks.get_Item(ref Markers.fullData).Range.Text = contract.ClientInfo;
                //oDoc.Bookmarks.get_Item(ref Markers.fullDataFIO).Range.Text = contract.ClientName;
                //oDoc.Bookmarks.get_Item(ref Markers.owningUser).Range.Text = contract.OwningUser;
                //oDoc.Bookmarks.get_Item(ref Markers.services).Range.Text = contract.ObjectTypeService;
                //oDoc.Bookmarks.get_Item(ref Markers.blockType).Range.Text = contract.BlockType;
                //oDoc.Bookmarks.get_Item(ref Markers.rent_AllCount).Range.Text = contract.AllCount;
                //oDoc.Bookmarks.get_Item(ref Markers.rent_AllSum).Range.Text = contract.AllSum;
                //oDoc.Bookmarks.get_Item(ref Markers.rent_deviceName).Range.Text = contract.DeviceName;
                //oDoc.Bookmarks.get_Item(ref Markers.rent_deviceQty).Range.Text = contract.DeviceCount;
                //oDoc.Bookmarks.get_Item(ref Markers.rent_deviceSum).Range.Text = contract.DeviceSum;
                //oDoc.Bookmarks.get_Item(ref Markers.FIO3).Range.Text = contract.ClientName;
                //oDoc.Bookmarks.get_Item(ref Markers.ClientName1).Range.Text = contract.ClientName;
                //oDoc.Bookmarks.get_Item(ref Markers.MainExecutorHeader).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseInfo;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorFullInfo).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).FullInfo;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorBossPosition).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition;
                //oDoc.Bookmarks.get_Item(ref Markers.MainExecutorBossPosition).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorBoss).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname;
                //oDoc.Bookmarks.get_Item(ref Markers.MainExecutorBossName).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorName1).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorName2).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorName3).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorName4).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                //oDoc.Bookmarks.get_Item(ref Markers.MainExecutorName).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                //oDoc.Bookmarks.get_Item(ref Markers.LicInfo).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Lic;
                //oDoc.Bookmarks.get_Item(ref Markers.LicDates).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicDates;
                //oDoc.Bookmarks.get_Item(ref Markers.LicCase).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicCase;
                //oDoc.Bookmarks.get_Item(ref Markers.Address).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Address;
                //oDoc.Bookmarks.get_Item(ref Markers.BossInfo).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossInfo;
                //oDoc.Bookmarks.get_Item(ref Markers.BossCompany).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossCompany;
                //foreach (var item in collection) {

                //}
                string tm = string.Empty;
                string tm1 = string.Empty;
                string tm2 = string.Empty;
                int i = 0;
                foreach (ExecutorsInfo item in contract.Coexecutors.Where(x => x.Id != contract.ExecutorID)) {
                    tm += item.LicenseInfo + Environment.NewLine;
                    tm1 += item.Name + ',';
                    object FullInfoExecutor = "FullInfoExecutor" + (i + 1);
                    object PositionBossExecutor = "PositionBossExecutor" + (i + 1);
                    object NameBossExecutor = "NameBossExecutor" + (i + 1);
                    oDoc.Bookmarks.get_Item(ref FullInfoExecutor).Range.Text = item.FullInfo;
                    oDoc.Bookmarks.get_Item(ref PositionBossExecutor).Range.Text = item.BossPosition;
                    oDoc.Bookmarks.get_Item(ref NameBossExecutor).Range.Text = item.Bossname;
                    i++;
                }
                foreach (ExecutorsInfo item in contract.Coexecutors) {
                    tm2 += item.Name + ',';
                }
                tm1 = tm1.Substring(0, tm1.LastIndexOf(',') - 1);
                tm2 = tm2.Substring(0, tm2.LastIndexOf(',') - 1);
                oDoc.Bookmarks.get_Item(ref Markers.CoExecutorHeader).Range.Text = tm;
                oDoc.Bookmarks.get_Item(ref Markers.coExecutors).Range.Text = tm1;
                oDoc.Bookmarks.get_Item(ref Markers.coExecutors1).Range.Text = tm1;
                oDoc.Bookmarks.get_Item(ref Markers.Executors).Range.Text = tm2;


                try {
                    var resultPath = currentPath.Remove(currentPath.Length - 4, 4) + "Results";
                    oDoc.SaveAs(toPathAndName);
                }
                catch (Exception e) {
                    throw new Exception("FATAL ERROR: ДОКУМЕНТ ОТКРЫТ - " + e.InnerException);
                }
                oDoc.Close();
                oWord.Quit();

                ProcessStartInfo processStartInfo = new ProcessStartInfo();
                processStartInfo.FileName = toPathAndName;
                Process.Start(processStartInfo);
            }
            catch (Exception e) {
                throw new Exception(e.Message + "\n" + e.ToString());
            }
        }
    }
}
