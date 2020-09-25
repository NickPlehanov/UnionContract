//using Microsoft.Office.Interop.Word;
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
                TryGetItem(oDoc, Markers.ContractNumber, contract.Number);

                TryGetItem(oDoc, Markers.ContractDate,contract.DateFull);
                TryGetItem(oDoc, Markers.ClientFullName,contract.ClientName);
                TryGetItem(oDoc, Markers.ExecutorFullName,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).FullName);
                TryGetItem(oDoc, Markers.ExecutorLicenseInfo,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseInfo);
                string CoExecutorsLicenseInfo = "";
                foreach (ExecutorsInfo item in contract.Coexecutors.Where(x => x.Id != contract.ExecutorID))
                    CoExecutorsLicenseInfo += item.Name + ", " + item.LicenseInfo + ");" + Environment.NewLine;
                TryGetItem(oDoc, Markers.CoExecutorsLicenseInfo,CoExecutorsLicenseInfo.Substring(0, CoExecutorsLicenseInfo.LastIndexOf(';')));
                TryGetItem(oDoc, Markers.BlockTypeDistinct,contract.BlockTypeDistinct);
                TryGetItem(oDoc, Markers.BlockTypeFull,contract.BlockTypeFull);
                TryGetItem(oDoc, Markers.ObjectName,contract.ObjectName);
                TryGetItem(oDoc, Markers.ObjectAddress,contract.ObjectAddress);
                TryGetItem(oDoc, Markers.ObjectCost,contract.ObjectCost);
                TryGetItem(oDoc, Markers.ObjectMonthlyPay,contract.ObjectMonthlyPay);
                TryGetItem(oDoc, Markers.ExecutorName,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
                TryGetItem(oDoc, Markers.ExecutorInfo,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).FullInfo);
                TryGetItem(oDoc, Markers.ExecutorPosition,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition);
                TryGetItem(oDoc, Markers.ExecutorPositionName,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname);
                TryGetItem(oDoc, Markers.ClientFullName1,contract.ClientName);
                TryGetItem(oDoc, Markers.ClientFullInfo,contract.ClientInfo);
                TryGetItem(oDoc, Markers.ClientName,contract.ClientSmallName);
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
                    TryGetItem(oDoc, CoExecutorName,item.Name);
                    TryGetItem(oDoc, CoExecutorInfo,item.FullInfo);
                    TryGetItem(oDoc, CoExecutorPosition,item.BossPosition);
                    TryGetItem(oDoc, CoExecutorPositionName,item.Bossname);
                    i++;
                }
                TryGetItem(oDoc, Markers.Manager,contract.OwningUser);
                TryGetItem(oDoc, Markers.ContractDate1,contract.DateFull);
                TryGetItem(oDoc, Markers.ExecutorName1,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
                TryGetItem(oDoc, Markers.ClientFullName2,contract.ClientName);
                TryGetItem(oDoc, Markers.ObjectName1,contract.ObjectName);
                TryGetItem(oDoc, Markers.ObjectAddress1,contract.ObjectAddress);
                TryGetItem(oDoc, Markers.DevicesName,contract.DeviceName);
                TryGetItem(oDoc, Markers.DevicesCount,contract.DeviceCount);
                TryGetItem(oDoc, Markers.DevicesSum,contract.DeviceSum);
                TryGetItem(oDoc, Markers.DevicesFullCount,contract.AllCount);
                TryGetItem(oDoc, Markers.DevicesFullSum,contract.AllSum);
                TryGetItem(oDoc, Markers.ExecutorName2,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
                TryGetItem(oDoc, Markers.ExecutorPosition1,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition);
                TryGetItem(oDoc, Markers.ExecutorPositionName1,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname);
                TryGetItem(oDoc, Markers.ClientName1,contract.ClientSmallName);
                TryGetItem(oDoc, Markers.ExecutorName3,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
                TryGetItem(oDoc, Markers.CoExecutors,tm1.Substring(0, tm1.LastIndexOf(',')));
                TryGetItem(oDoc, Markers.ContractNumber1,contract.Number);
                TryGetItem(oDoc, Markers.ContractDate2,contract.Date);
                TryGetItem(oDoc, Markers.ClientFullNameAndObject,contract.ClientName + ", " + contract.ObjectName);
                TryGetItem(oDoc, Markers.ObjectAddress2,contract.ObjectAddress);
                foreach (ExecutorsInfo item in contract.Coexecutors)
                    tm2 += item.Name + ", ";
                TryGetItem(oDoc, Markers.Executors,tm2.Substring(0, tm2.LastIndexOf(',') - 1));
                TryGetItem(oDoc, Markers.ExecutorName4,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
                TryGetItem(oDoc, Markers.ExecutorName5,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
                TryGetItem(oDoc, Markers.LicenseNumber,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicNum);
                TryGetItem(oDoc, Markers.LicenseIssued,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicIssued);
                TryGetItem(oDoc, Markers.LicenseIssuedDate,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicIssuedDate);
                TryGetItem(oDoc, Markers.LicenseExpired,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseExpired);
                TryGetItem(oDoc, Markers.LicenseDelo,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseDelo);
                TryGetItem(oDoc, Markers.LicenseAddress,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseAddress);
                TryGetItem(oDoc, Markers.LicensePhone,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicensePhone);
                TryGetItem(oDoc, Markers.LicensePositionAndNameAndPhone,
                    contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition + " " +
                    contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname + " " +
                    contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPhone);
                TryGetItem(oDoc, Markers.ExecutorPosition2,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition);
                TryGetItem(oDoc, Markers.ExecutorPositionName2,contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname);
                TryGetItem(oDoc, Markers.CoExecutors1,tm1.Substring(0, tm1.LastIndexOf(',') - 1));
                TryGetItem(oDoc, Markers.OS,contract.os);
                TryGetItem(oDoc, Markers.PS,contract.ps);
                TryGetItem(oDoc, Markers.TRS,contract.trs);
                TryGetItem(oDoc, Markers.SecurityService,contract.service_security);

                TryGetItem(oDoc, Markers.SendAct,contract.SendActs);
                TryGetItem(oDoc, Markers.OrderAct,contract.PoryadokActs);
                TryGetItem(oDoc, Markers.SignalingOS,contract.SignalingOS);
                TryGetItem(oDoc, Markers.SignalingPS,contract.SignalingPS);
                TryGetItem(oDoc, Markers.SignalingTRS,contract.SignalingTRS);

                //oDoc.Bookmarks.get_Item(ref Markers.ContractNumber).Range.Text = contract.Number;
                //oDoc.Bookmarks.get_Item(ref Markers.ContractDate).Range.Text = contract.DateFull;
                //oDoc.Bookmarks.get_Item(ref Markers.ClientFullName).Range.Text = contract.ClientName;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorFullName).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).FullName;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorLicenseInfo).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseInfo;
                //string CoExecutorsLicenseInfo="";
                //foreach (ExecutorsInfo item in contract.Coexecutors.Where(x => x.Id != contract.ExecutorID)) 
                //    CoExecutorsLicenseInfo += item.Name + ", " + item.LicenseInfo + ";" + Environment.NewLine;
                //oDoc.Bookmarks.get_Item(ref Markers.CoExecutorsLicenseInfo).Range.Text = CoExecutorsLicenseInfo.Substring(0, CoExecutorsLicenseInfo.LastIndexOf(';'));
                //oDoc.Bookmarks.get_Item(ref Markers.BlockTypeDistinct).Range.Text = contract.BlockTypeDistinct;
                //oDoc.Bookmarks.get_Item(ref Markers.BlockTypeFull).Range.Text = contract.BlockTypeFull;
                //oDoc.Bookmarks.get_Item(ref Markers.ObjectName).Range.Text = contract.ObjectName;
                //oDoc.Bookmarks.get_Item(ref Markers.ObjectAddress).Range.Text = contract.ObjectAddress;
                //oDoc.Bookmarks.get_Item(ref Markers.ObjectCost).Range.Text = contract.ObjectCost;
                //oDoc.Bookmarks.get_Item(ref Markers.ObjectMonthlyPay).Range.Text = contract.ObjectMonthlyPay;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorName).Range.Text = contract.Coexecutors.FirstOrDefault(x=>x.Id==contract.ExecutorID).Name;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorInfo).Range.Text = contract.Coexecutors.FirstOrDefault(x=>x.Id==contract.ExecutorID).FullInfo;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorPosition).Range.Text = contract.Coexecutors.FirstOrDefault(x=>x.Id==contract.ExecutorID).BossPosition;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorPositionName).Range.Text = contract.Coexecutors.FirstOrDefault(x=>x.Id==contract.ExecutorID).Bossname;
                //oDoc.Bookmarks.get_Item(ref Markers.ClientFullName1).Range.Text = contract.ClientName;
                //oDoc.Bookmarks.get_Item(ref Markers.ClientFullInfo).Range.Text = contract.ClientInfo;
                //oDoc.Bookmarks.get_Item(ref Markers.ClientName).Range.Text = contract.ClientSmallName;
                //string tm = string.Empty;
                //string tm1 = string.Empty;
                //string tm2 = string.Empty;
                //int i = 0;
                //var y = contract.Coexecutors.Where(x => x.Id != contract.ExecutorID);
                //foreach (ExecutorsInfo item in contract.Coexecutors.Where(x => x.Id != contract.ExecutorID)) {
                //    tm += item.LicenseInfo + Environment.NewLine;
                //    tm1 += item.Name + ", ";
                //    object CoExecutorName = "CoExecutorName" + (i + 1);
                //    object CoExecutorInfo = "CoExecutorInfo" + (i + 1);
                //    object CoExecutorPosition = "CoExecutorPosition" + (i + 1);
                //    object CoExecutorPositionName = "CoExecutorPositionName" + (i + 1);
                //    oDoc.Bookmarks.get_Item(ref CoExecutorName).Range.Text = item.Name;
                //    oDoc.Bookmarks.get_Item(ref CoExecutorInfo).Range.Text = item.FullInfo;
                //    oDoc.Bookmarks.get_Item(ref CoExecutorPosition).Range.Text = item.BossPosition;
                //    oDoc.Bookmarks.get_Item(ref CoExecutorPositionName).Range.Text = item.Bossname;
                //    i++;
                //}
                //oDoc.Bookmarks.get_Item(ref Markers.Manager).Range.Text = contract.OwningUser;
                //oDoc.Bookmarks.get_Item(ref Markers.ContractDate1).Range.Text = contract.DateFull;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorName1).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                //oDoc.Bookmarks.get_Item(ref Markers.ClientFullName2).Range.Text = contract.ClientName;
                //oDoc.Bookmarks.get_Item(ref Markers.ObjectName1).Range.Text = contract.ObjectName;
                //oDoc.Bookmarks.get_Item(ref Markers.ObjectAddress1).Range.Text = contract.ObjectAddress;
                //oDoc.Bookmarks.get_Item(ref Markers.DevicesName).Range.Text = contract.DeviceName;
                //oDoc.Bookmarks.get_Item(ref Markers.DevicesCount).Range.Text = contract.DeviceCount;
                //oDoc.Bookmarks.get_Item(ref Markers.DevicesSum).Range.Text = contract.DeviceSum;
                //oDoc.Bookmarks.get_Item(ref Markers.DevicesFullCount).Range.Text = contract.AllCount;
                //oDoc.Bookmarks.get_Item(ref Markers.DevicesFullSum).Range.Text = contract.AllSum;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorName2).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorPosition1).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorPositionName1).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname;
                //oDoc.Bookmarks.get_Item(ref Markers.ClientName1).Range.Text = contract.ClientSmallName;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorName3).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                //oDoc.Bookmarks.get_Item(ref Markers.CoExecutors).Range.Text = tm1.Substring(0, tm1.LastIndexOf(','));
                //oDoc.Bookmarks.get_Item(ref Markers.ContractNumber1).Range.Text = contract.Number;
                //oDoc.Bookmarks.get_Item(ref Markers.ContractDate2).Range.Text = contract.Date;
                //oDoc.Bookmarks.get_Item(ref Markers.ClientFullNameAndObject).Range.Text = contract.ClientName+", "+contract.ObjectName;
                //oDoc.Bookmarks.get_Item(ref Markers.ObjectAddress2).Range.Text = contract.ObjectAddress;
                //foreach (ExecutorsInfo item in contract.Coexecutors) 
                //    tm2 += item.Name + ", ";
                //oDoc.Bookmarks.get_Item(ref Markers.Executors).Range.Text = tm2.Substring(0, tm2.LastIndexOf(',') - 1);
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorName4).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorName5).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name;
                //oDoc.Bookmarks.get_Item(ref Markers.LicenseNumber).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicNum;
                //oDoc.Bookmarks.get_Item(ref Markers.LicenseIssued).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicIssued;
                //oDoc.Bookmarks.get_Item(ref Markers.LicenseIssuedDate).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicIssuedDate;
                //oDoc.Bookmarks.get_Item(ref Markers.LicenseExpired).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseExpired;
                //oDoc.Bookmarks.get_Item(ref Markers.LicenseDelo).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseDelo;
                //oDoc.Bookmarks.get_Item(ref Markers.LicenseAddress).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseAddress;
                //oDoc.Bookmarks.get_Item(ref Markers.LicensePhone).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicensePhone;
                //oDoc.Bookmarks.get_Item(ref Markers.LicensePositionAndNameAndPhone).Range.Text = 
                //    contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition + " " + 
                //    contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname + " "+ 
                //    contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPhone;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorPosition2).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition;
                //oDoc.Bookmarks.get_Item(ref Markers.ExecutorPositionName2).Range.Text = contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname;
                //oDoc.Bookmarks.get_Item(ref Markers.CoExecutors1).Range.Text = tm1.Substring(0, tm1.LastIndexOf(',') - 1);
                //oDoc.Bookmarks.get_Item(ref Markers.OS).Range.Text = contract.os;
                //oDoc.Bookmarks.get_Item(ref Markers.PS).Range.Text = contract.ps;
                //oDoc.Bookmarks.get_Item(ref Markers.TRS).Range.Text = contract.trs;
                //oDoc.Bookmarks.get_Item(ref Markers.SecurityService).Range.Text = contract.service_security;

                //oDoc.Bookmarks.get_Item(ref Markers.SendAct).Range.Text = contract.SendActs;
                //oDoc.Bookmarks.get_Item(ref Markers.OrderAct).Range.Text = contract.PoryadokActs;
                //oDoc.Bookmarks.get_Item(ref Markers.SignalingOS).Range.Text = contract.SignalingOS;
                //oDoc.Bookmarks.get_Item(ref Markers.SignalingPS).Range.Text = contract.SignalingPS;
                //oDoc.Bookmarks.get_Item(ref Markers.SignalingTRS).Range.Text = contract.SignalingTRS;





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
