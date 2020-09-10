//using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System;
using System.Diagnostics;
using System.Linq;
using UnionContractWF.Models;

namespace UnionContractWF.Helpers {
    public static class WordDocument {
        private static void TryGetItem(Document oDoc, object mark, string value, int isBold) {
            try {
                oDoc.Bookmarks.get_Item(ref mark).Range.Text = value;
                oDoc.Bookmarks.get_Item(ref mark).Range.Font.BoldBi = isBold;
                Bookmark b = oDoc.Bookmarks.get_Item(ref mark);
                //object endDocument = oDoc.Bookmarks.get_Item(ref mark).Range;
                //Paragraph para = oDoc.Content.Paragraphs.Add(ref endDocument);
                //para.Range.Text = value;
                //para.Range.Font.Bold = isBold;
            }
            catch { }
        }

        public static void Exchange(Contract contract, string fromPathAndName, string toPathAndName) {
            try {
                //Spire.Doc.Document document = new Spire.Doc.Document();
                //document.LoadFromFile(fromPathAndName);

                //Spire.Doc.Documents.BookmarksNavigator bookmarkNavigator = new Spire.Doc.Documents.BookmarksNavigator(document);
                //bookmarkNavigator.MoveToBookmark("contract_number", true, true);
                //bookmarkNavigator.ReplaceBookmarkContent(contract.Number, true);
                //bookmarkNavigator.MoveToBookmark("object_address4", true, true);
                //bookmarkNavigator.ReplaceBookmarkContent(contract.ObjectAddress, true);

                //document.SaveToFile(toPathAndName);
                //document.Close();
                //ProcessStartInfo processStartInfo = new ProcessStartInfo();
                //processStartInfo.FileName = toPathAndName;
                //Process.Start(processStartInfo);

                var currentPath = AppDomain.CurrentDomain.BaseDirectory;

                var oWord = new Microsoft.Office.Interop.Word.Application();

                var oDoc = oWord.Documents.Add(fromPathAndName);
                //foreach (Bookmark item in oDoc.Bookmarks) {
                //    item.Range.Font.Bold = 1;
                //    item.Range.Text = contract.Number;
                //}
                oDoc.Bookmarks.get_Item(ref Markers.ContractNumber).Range.Text = contract.Number;
                oDoc.Bookmarks.get_Item(ref Markers.ContractDate).Range.Text = contract.DateFull;
                oDoc.Bookmarks.get_Item(ref Markers.ClientFullName).Range.Text = contract.ClientName;
                oDoc.Bookmarks.get_Item(ref Markers.ExecutorFullName).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).FullName;
                oDoc.Bookmarks.get_Item(ref Markers.ExecutorLicenseInfo).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseInfo;
                string CoExecutorsLicenseInfo="";
                foreach (ExecutorsInfo item in contract.Coexecutors.Where(x => x.Id != contract.ExecutorID)) 
                    CoExecutorsLicenseInfo += item.Name + ", " + item.LicenseInfo + ";" + Environment.NewLine;
                oDoc.Bookmarks.get_Item(ref Markers.CoExecutorsLicenseInfo).Range.Text = CoExecutorsLicenseInfo.Substring(0, CoExecutorsLicenseInfo.LastIndexOf(';'));
                oDoc.Bookmarks.get_Item(ref Markers.BlockTypeDistinct).Range.Text = contract.BlockTypeDistinct;
                oDoc.Bookmarks.get_Item(ref Markers.BlockTypeFull).Range.Text = contract.BlockTypeFull;
                oDoc.Bookmarks.get_Item(ref Markers.ObjectName).Range.Text = contract.ObjectName;
                oDoc.Bookmarks.get_Item(ref Markers.ObjectAddress).Range.Text = contract.ObjectAddress;
                oDoc.Bookmarks.get_Item(ref Markers.ObjectCost).Range.Text = contract.ObjectCost;
                oDoc.Bookmarks.get_Item(ref Markers.ObjectMonthlyPay).Range.Text = contract.ObjectMonthlyPay;
                oDoc.Bookmarks.get_Item(ref Markers.ExecutorName).Range.Text = contract.Coexecutors.FirstOrDefault(x=>x.Id==contract.ExecutorID).Name;
                oDoc.Bookmarks.get_Item(ref Markers.ExecutorInfo).Range.Text = contract.Coexecutors.FirstOrDefault(x=>x.Id==contract.ExecutorID).FullInfo;
                oDoc.Bookmarks.get_Item(ref Markers.ExecutorPosition).Range.Text = contract.Coexecutors.FirstOrDefault(x=>x.Id==contract.ExecutorID).BossPosition;
                oDoc.Bookmarks.get_Item(ref Markers.ExecutorPositionName).Range.Text = contract.Coexecutors.FirstOrDefault(x=>x.Id==contract.ExecutorID).Bossname;
                oDoc.Bookmarks.get_Item(ref Markers.ClientFullName1).Range.Text = contract.ClientName;
                oDoc.Bookmarks.get_Item(ref Markers.ClientFullInfo).Range.Text = contract.ClientInfo;
                oDoc.Bookmarks.get_Item(ref Markers.ClientName).Range.Text = contract.ClientSmallName;
                string tm = string.Empty;
                string tm1 = string.Empty;
                string tm2 = string.Empty;
                int i = 0;
                var y = contract.Coexecutors.Where(x => x.Id != contract.ExecutorID);
                foreach (ExecutorsInfo item in contract.Coexecutors.Where(x => x.Id != contract.ExecutorID)) {
                    tm += item.LicenseInfo + Environment.NewLine;
                    tm1 += item.Name + ", ";
                    object CoExecutorName = "CoExecutorName" + (i + 1);
                    object CoExecutorInfo = "CoExecutorInfo" + (i + 1);
                    object CoExecutorPosition = "CoExecutorPosition" + (i + 1);
                    object CoExecutorPositionName = "CoExecutorPositionName" + (i + 1);
                    oDoc.Bookmarks.get_Item(ref CoExecutorName).Range.Text = item.Name;
                    oDoc.Bookmarks.get_Item(ref CoExecutorInfo).Range.Text = item.FullInfo;
                    oDoc.Bookmarks.get_Item(ref CoExecutorPosition).Range.Text = item.BossPosition;
                    oDoc.Bookmarks.get_Item(ref CoExecutorPositionName).Range.Text = item.Bossname;
                    i++;
                }
                oDoc.Bookmarks.get_Item(ref Markers.Manager).Range.Text = contract.OwningUser;
                oDoc.Bookmarks.get_Item(ref Markers.ContractDate1).Range.Text = contract.DateFull;
                oDoc.Bookmarks.get_Item(ref Markers.ExecutorName1).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                oDoc.Bookmarks.get_Item(ref Markers.ClientFullName2).Range.Text = contract.ClientName;
                oDoc.Bookmarks.get_Item(ref Markers.ObjectName1).Range.Text = contract.ObjectName;
                oDoc.Bookmarks.get_Item(ref Markers.ObjectAddress1).Range.Text = contract.ObjectAddress;
                oDoc.Bookmarks.get_Item(ref Markers.DevicesName).Range.Text = contract.DeviceName;
                oDoc.Bookmarks.get_Item(ref Markers.DevicesCount).Range.Text = contract.DeviceCount;
                oDoc.Bookmarks.get_Item(ref Markers.DevicesSum).Range.Text = contract.DeviceSum;
                oDoc.Bookmarks.get_Item(ref Markers.DevicesFullCount).Range.Text = contract.AllCount;
                oDoc.Bookmarks.get_Item(ref Markers.DevicesFullSum).Range.Text = contract.AllSum;
                oDoc.Bookmarks.get_Item(ref Markers.ExecutorName2).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                oDoc.Bookmarks.get_Item(ref Markers.ExecutorPosition1).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition;
                oDoc.Bookmarks.get_Item(ref Markers.ExecutorPositionName1).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname;
                oDoc.Bookmarks.get_Item(ref Markers.ClientName1).Range.Text = contract.ClientSmallName;
                oDoc.Bookmarks.get_Item(ref Markers.ExecutorName3).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                oDoc.Bookmarks.get_Item(ref Markers.CoExecutors).Range.Text = tm1.Substring(0, tm1.LastIndexOf(','));
                oDoc.Bookmarks.get_Item(ref Markers.ContractNumber1).Range.Text = contract.Number;
                oDoc.Bookmarks.get_Item(ref Markers.ContractDate2).Range.Text = contract.Date;
                oDoc.Bookmarks.get_Item(ref Markers.ClientFullNameAndObject).Range.Text = contract.ClientName+", "+contract.ObjectName;
                oDoc.Bookmarks.get_Item(ref Markers.ObjectAddress2).Range.Text = contract.ObjectAddress;
                foreach (ExecutorsInfo item in contract.Coexecutors) 
                    tm2 += item.Name + ", ";
                oDoc.Bookmarks.get_Item(ref Markers.Executors).Range.Text = tm2.Substring(0, tm2.LastIndexOf(',') - 1);
                oDoc.Bookmarks.get_Item(ref Markers.ExecutorName4).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                oDoc.Bookmarks.get_Item(ref Markers.ExecutorName5).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                oDoc.Bookmarks.get_Item(ref Markers.LicenseNumber).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicNum;
                oDoc.Bookmarks.get_Item(ref Markers.LicenseIssued).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicIssued;
                oDoc.Bookmarks.get_Item(ref Markers.LicenseIssuedDate).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicIssuedDate;
                oDoc.Bookmarks.get_Item(ref Markers.LicenseExpired).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseExpired;
                oDoc.Bookmarks.get_Item(ref Markers.LicenseDelo).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseDelo;
                oDoc.Bookmarks.get_Item(ref Markers.LicenseAddress).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseAddress;
                oDoc.Bookmarks.get_Item(ref Markers.LicensePhone).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicensePhone;
                oDoc.Bookmarks.get_Item(ref Markers.LicensePositionAndNameAndPhone).Range.Text = 
                    contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition + " " + 
                    contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname + " "+ 
                    contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPhone;
                oDoc.Bookmarks.get_Item(ref Markers.ExecutorPosition2).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition;
                oDoc.Bookmarks.get_Item(ref Markers.ExecutorPositionName2).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname;
                oDoc.Bookmarks.get_Item(ref Markers.CoExecutors1).Range.Text = tm1.Substring(0, tm1.LastIndexOf(',') - 1);
                oDoc.Bookmarks.get_Item(ref Markers.OS).Range.Text = contract.os;
                oDoc.Bookmarks.get_Item(ref Markers.PS).Range.Text = contract.ps;
                oDoc.Bookmarks.get_Item(ref Markers.TRS).Range.Text = contract.trs;
                oDoc.Bookmarks.get_Item(ref Markers.SecurityService).Range.Text = contract.service_security;






                //TryGetItem(oDoc, Markers.contract_number, contract.Number, 1);
                //TryGetItem(oDoc, Markers.ClientFullName, contract.Number, 1);





                //TryGetItem(oDoc, Markers.agent_type_header, contract.ClientName, 1);
                //TryGetItem(oDoc, Markers.MainExecutorHeader, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseInfo, 1);
                //TryGetItem(oDoc, Markers.object_address4, contract.ObjectAddress, 1);
                //TryGetItem(oDoc, Markers.object_name3, contract.ObjectName, 1);
                //TryGetItem(oDoc, Markers.FIO3, contract.ClientName, 1);
                //TryGetItem(oDoc, Markers.ExecutorName4, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name, 1);
                //TryGetItem(oDoc, Markers.MainExecutorName, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name, 1);
                //TryGetItem(oDoc, Markers.LicCase, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicCase, 1);
                //TryGetItem(oDoc, Markers.Address, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Address, 1);
                //TryGetItem(oDoc, Markers.BossInfo, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossInfo, 1);

                //TryGetItem(oDoc, Markers.contract_number4, contract.Number, 0);
                //TryGetItem(oDoc, Markers.object_name, contract.ObjectName, 0);
                //TryGetItem(oDoc, Markers.date_day, contract.Date, 0);
                //TryGetItem(oDoc, Markers.date_day1, contract.Date, 0);
                //TryGetItem(oDoc, Markers.agent_type_header1, contract.ClientSmallName, 0);
                //TryGetItem(oDoc, Markers.object_address, contract.ObjectAddress, 0);
                //TryGetItem(oDoc, Markers.abonent_cost, contract.ObjectMonthlyPay, 0);
                //TryGetItem(oDoc, Markers.object_cost, contract.ObjectCost, 0);
                //TryGetItem(oDoc, Markers.object_name1, contract.ObjectName, 0);
                //TryGetItem(oDoc, Markers.object_address1, contract.ObjectAddress, 0);
                //TryGetItem(oDoc, Markers.contract_date3, contract.Date, 0);
                //TryGetItem(oDoc, Markers.fullData, contract.ClientInfo, 0);
                //TryGetItem(oDoc, Markers.fullDataFIO, contract.ClientSmallName, 0);
                //TryGetItem(oDoc, Markers.owningUser, contract.OwningUser, 0);
                //TryGetItem(oDoc, Markers.services, contract.ObjectTypeService, 0);
                //TryGetItem(oDoc, Markers.blockType, contract.BlockType, 0);
                //TryGetItem(oDoc, Markers.rent_AllCount, contract.AllCount, 0);
                //TryGetItem(oDoc, Markers.rent_AllSum, contract.AllSum, 0);
                //TryGetItem(oDoc, Markers.rent_deviceName, contract.DeviceName, 0);
                //TryGetItem(oDoc, Markers.rent_deviceQty, contract.DeviceCount, 0);
                //TryGetItem(oDoc, Markers.rent_deviceSum, contract.DeviceSum, 0);
                //TryGetItem(oDoc, Markers.ClientName1, contract.ClientName, 0);
                //TryGetItem(oDoc, Markers.ExecutorFullInfo, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).FullInfo, 0);
                //TryGetItem(oDoc, Markers.ExecutorBossPosition, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition, 0);
                //TryGetItem(oDoc, Markers.MainExecutorBossPosition, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition, 0);
                //TryGetItem(oDoc, Markers.MainExecutorBossPosition1, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition, 0);
                //TryGetItem(oDoc, Markers.ExecutorBoss, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname, 0);
                //TryGetItem(oDoc, Markers.MainExecutorBossName, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname, 0);
                //TryGetItem(oDoc, Markers.MainExecutorBossName1, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname, 0);
                //TryGetItem(oDoc, Markers.ExecutorName1, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name, 0);
                //TryGetItem(oDoc, Markers.ExecutorName2, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name, 0);
                //TryGetItem(oDoc, Markers.ExecutorName3, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name, 0);
                //TryGetItem(oDoc, Markers.LicInfo, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Lic, 0);
                //TryGetItem(oDoc, Markers.LicDates, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicDates, 0);
                //TryGetItem(oDoc, Markers.BossCompany, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossCompany, 0);


                //TryGetItem(oDoc, Markers.object_name6, contract.ObjectName, 0);
                //TryGetItem(oDoc, Markers.object_address5, contract.ObjectAddress, 0);
                //TryGetItem(oDoc, Markers.SendAct, contract.SendActs, 0);
                //TryGetItem(oDoc, Markers.PoryadokAct, contract.PoryadokActs, 0);
                //TryGetItem(oDoc, Markers.object_signalizations, contract.ObjectSignalization, 0);

                ////oDoc.Bookmarks.get_Item(ref Markers.contract_number).Range.Text = contract.Number;
                ////oDoc.Bookmarks.get_Item(ref Markers.contract_number4).Range.Text = contract.Number;
                ////oDoc.Bookmarks.get_Item(ref Markers.object_name).Range.Text = contract.ObjectName;
                ////oDoc.Bookmarks.get_Item(ref Markers.date_day).Range.Text = contract.Date;
                ////oDoc.Bookmarks.get_Item(ref Markers.date_day1).Range.Text = contract.Date;
                ////oDoc.Bookmarks.get_Item(ref Markers.agent_type_header).Range.Text = contract.ClientName;
                ////oDoc.Bookmarks.get_Item(ref Markers.agent_type_header1).Range.Text = contract.ClientName;
                ////oDoc.Bookmarks.get_Item(ref Markers.object_address).Range.Text = contract.ObjectAddress;
                ////oDoc.Bookmarks.get_Item(ref Markers.abonent_cost).Range.Text = contract.ObjectMonthlyPay;
                ////oDoc.Bookmarks.get_Item(ref Markers.object_cost).Range.Text = contract.ObjectCost;
                ////oDoc.Bookmarks.get_Item(ref Markers.object_name1).Range.Text = contract.ObjectName;
                ////oDoc.Bookmarks.get_Item(ref Markers.object_name6).Range.Text = contract.ObjectName;
                ////oDoc.Bookmarks.get_Item(ref Markers.object_address1).Range.Text = contract.ObjectAddress;
                ////oDoc.Bookmarks.get_Item(ref Markers.contract_date3).Range.Text = contract.Date;
                ////oDoc.Bookmarks.get_Item(ref Markers.object_address4).Range.Text = contract.ObjectAddress;
                ////oDoc.Bookmarks.get_Item(ref Markers.object_address5).Range.Text = contract.ObjectAddress;
                ////oDoc.Bookmarks.get_Item(ref Markers.object_name3).Range.Text = contract.ObjectName;
                ////oDoc.Bookmarks.get_Item(ref Markers.fullData).Range.Text = contract.ClientInfo;
                ////oDoc.Bookmarks.get_Item(ref Markers.fullDataFIO).Range.Text = contract.ClientName;
                ////oDoc.Bookmarks.get_Item(ref Markers.owningUser).Range.Text = contract.OwningUser;
                ////oDoc.Bookmarks.get_Item(ref Markers.services).Range.Text = contract.ObjectTypeService;
                ////oDoc.Bookmarks.get_Item(ref Markers.blockType).Range.Text = contract.BlockType;
                ////oDoc.Bookmarks.get_Item(ref Markers.rent_AllCount).Range.Text = contract.AllCount;
                ////oDoc.Bookmarks.get_Item(ref Markers.rent_AllSum).Range.Text = contract.AllSum;
                ////oDoc.Bookmarks.get_Item(ref Markers.rent_deviceName).Range.Text = contract.DeviceName;
                ////oDoc.Bookmarks.get_Item(ref Markers.rent_deviceQty).Range.Text = contract.DeviceCount;
                ////oDoc.Bookmarks.get_Item(ref Markers.rent_deviceSum).Range.Text = contract.DeviceSum;
                ////oDoc.Bookmarks.get_Item(ref Markers.FIO3).Range.Text = contract.ClientName;
                ////oDoc.Bookmarks.get_Item(ref Markers.ClientName1).Range.Text = contract.ClientName;
                ////oDoc.Bookmarks.get_Item(ref Markers.MainExecutorHeader).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseInfo;
                ////oDoc.Bookmarks.get_Item(ref Markers.ExecutorFullInfo).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).FullInfo;
                ////oDoc.Bookmarks.get_Item(ref Markers.ExecutorBossPosition).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition;
                ////oDoc.Bookmarks.get_Item(ref Markers.MainExecutorBossPosition).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition;
                ////oDoc.Bookmarks.get_Item(ref Markers.ExecutorBoss).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname;
                ////oDoc.Bookmarks.get_Item(ref Markers.MainExecutorBossName).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname;
                ////oDoc.Bookmarks.get_Item(ref Markers.ExecutorName1).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                ////oDoc.Bookmarks.get_Item(ref Markers.ExecutorName2).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                ////oDoc.Bookmarks.get_Item(ref Markers.ExecutorName3).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                ////oDoc.Bookmarks.get_Item(ref Markers.ExecutorName4).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                ////oDoc.Bookmarks.get_Item(ref Markers.MainExecutorName).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                ////oDoc.Bookmarks.get_Item(ref Markers.LicInfo).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Lic;
                ////oDoc.Bookmarks.get_Item(ref Markers.LicDates).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicDates;
                ////oDoc.Bookmarks.get_Item(ref Markers.LicCase).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicCase;
                ////oDoc.Bookmarks.get_Item(ref Markers.Address).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Address;
                ////oDoc.Bookmarks.get_Item(ref Markers.BossInfo).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossInfo;
                ////oDoc.Bookmarks.get_Item(ref Markers.BossCompany).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossCompany;
                ////foreach (var item in collection) {

                ////}
                //string tm = string.Empty;
                //string tm1 = string.Empty;
                //string tm2 = string.Empty;
                //int i = 0;
                //foreach (ExecutorsInfo item in contract.Coexecutors.Where(x => x.Id != contract.ExecutorID)) {
                //    tm += item.LicenseInfo + Environment.NewLine;
                //    tm1 += item.Name + ',';
                //    object FullInfoExecutor = "FullInfoExecutor" + (i + 1);
                //    object PositionBossExecutor = "PositionBossExecutor" + (i + 1);
                //    object NameBossExecutor = "NameBossExecutor" + (i + 1);
                //    TryGetItem(oDoc, FullInfoExecutor, item.FullInfo, 0);
                //    TryGetItem(oDoc, PositionBossExecutor, item.BossPosition, 0);
                //    //TryGetItem(oDoc, NameBossExecutor, item.Bossname, 0);

                //    //oDoc.Bookmarks.get_Item(ref FullInfoExecutor).Range.Text = item.FullInfo;
                //    //oDoc.Bookmarks.get_Item(ref PositionBossExecutor).Range.Text = item.BossPosition;
                //    //oDoc.Bookmarks.get_Item(ref NameBossExecutor).Range.Text = item.Bossname;
                //    i++;
                //}
                //foreach (ExecutorsInfo item in contract.Coexecutors) {
                //    tm2 += item.Name + ',';
                //}
                //tm1 = tm1.Substring(0, tm1.LastIndexOf(',') - 1);
                //tm2 = tm2.Substring(0, tm2.LastIndexOf(',') - 1);

                //TryGetItem(oDoc, Markers.CoExecutorHeader, tm, 0);
                //TryGetItem(oDoc, Markers.coExecutors, tm, 0);
                //TryGetItem(oDoc, Markers.coExecutors1, tm, 0);
                //TryGetItem(oDoc, Markers.Executors, tm, 0);

                //oDoc.Bookmarks.get_Item(ref Markers.CoExecutorHeader).Range.Text = tm;
                //oDoc.Bookmarks.get_Item(ref Markers.coExecutors).Range.Text = tm1;
                //oDoc.Bookmarks.get_Item(ref Markers.coExecutors1).Range.Text = tm1;
                //oDoc.Bookmarks.get_Item(ref Markers.Executors).Range.Text = tm2;


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

        //private static void Document_BookmarkLayout(object sender, Spire.Doc.Documents.Rendering.BookmarkLevelEventArgs args) {
        //    throw new NotImplementedException();
        //}
    }
}
