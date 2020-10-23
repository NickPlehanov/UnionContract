//using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
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

				//File.Copy(fromPathAndName, toPathAndName, true);

				var oWord = new Microsoft.Office.Interop.Word.Application();

				var oDoc = oWord.Documents.Add(fromPathAndName);
				//var oDoc = oWord.Documents.Add(toPathAndName);
				TryGetItem(oDoc, Markers.ContractNumber, contract.Number);

				TryGetItem(oDoc, Markers.ContractDate, contract.DateFull);
				TryGetItem(oDoc, Markers.ClientFullName, contract.ClientName);
				TryGetItem(oDoc, Markers.ExecutorFullName, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).FullName);
				TryGetItem(oDoc, Markers.ExecutorLicenseInfo, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseInfo);
				string CoExecutorsLicenseInfo = "";
				foreach(ExecutorsInfo item in contract.Coexecutors.Where(x => x.Id != contract.ExecutorID))
					CoExecutorsLicenseInfo += item.Name + ", " + item.LicenseInfo + ");" + Environment.NewLine;
				TryGetItem(oDoc, Markers.CoExecutorsLicenseInfo, CoExecutorsLicenseInfo.Length > 0 ? CoExecutorsLicenseInfo.Substring(0, CoExecutorsLicenseInfo.LastIndexOf(';')) : null);
				TryGetItem(oDoc, Markers.BlockTypeDistinct, contract.BlockTypeDistinct);
				TryGetItem(oDoc, Markers.BlockTypeFull, contract.BlockTypeFull);
				TryGetItem(oDoc, Markers.ObjectName, contract.ObjectName);
				TryGetItem(oDoc, Markers.ObjectAddress, contract.ObjectAddress);
				TryGetItem(oDoc, Markers.ObjectCost, contract.ObjectCost);
				TryGetItem(oDoc, Markers.ObjectMonthlyPay, contract.ObjectMonthlyPay);
				TryGetItem(oDoc, Markers.ExecutorName, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
				TryGetItem(oDoc, Markers.ExecutorInfo, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).FullInfo);
				TryGetItem(oDoc, Markers.ExecutorPosition, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition);
				TryGetItem(oDoc, Markers.ExecutorPositionName, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname);
				TryGetItem(oDoc, Markers.ClientFullName1, contract.ClientName);
				TryGetItem(oDoc, Markers.ClientFullInfo, contract.ClientInfo);
				TryGetItem(oDoc, Markers.ClientName, contract.ClientSmallName);
				string tm = string.Empty;
				string tm1 = string.Empty;
				string tm2 = string.Empty;
				int i = 0;
				var y = contract.Coexecutors.Where(x => x.Id != contract.ExecutorID);
				foreach(ExecutorsInfo item in contract.Coexecutors.Where(x => x.Id != contract.ExecutorID)) {
					tm += item.LicenseInfo + Environment.NewLine;
					tm1 += item.Name + ", ";
					object CoExecutorName = "CoExecutorName" + (i + 1);
					object CoExecutorInfo = "CoExecutorInfo" + (i + 1);
					object CoExecutorPosition = "CoExecutorPosition" + (i + 1);
					object CoExecutorPositionName = "CoExecutorPositionName" + (i + 1);
					TryGetItem(oDoc, CoExecutorName, item.Name);
					TryGetItem(oDoc, CoExecutorInfo, item.FullInfo);
					TryGetItem(oDoc, CoExecutorPosition, item.BossPosition);
					TryGetItem(oDoc, CoExecutorPositionName, item.Bossname);
					i++;
				}
				TryGetItem(oDoc, Markers.Manager, contract.OwningUser);
				TryGetItem(oDoc, Markers.ContractDate1, contract.DateFull);
				TryGetItem(oDoc, Markers.ExecutorName1, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
				TryGetItem(oDoc, Markers.ClientFullName2, contract.ClientName);
				TryGetItem(oDoc, Markers.ObjectName1, contract.ObjectName);
				TryGetItem(oDoc, Markers.ObjectAddress1, contract.ObjectAddress);
				TryGetItem(oDoc, Markers.DevicesName, contract.DeviceName);
				TryGetItem(oDoc, Markers.DevicesCount, contract.DeviceCount);
				TryGetItem(oDoc, Markers.DevicesSum, contract.DeviceSum);
				TryGetItem(oDoc, Markers.DevicesFullCount, contract.AllCount);
				TryGetItem(oDoc, Markers.DevicesFullSum, contract.AllSum);
				TryGetItem(oDoc, Markers.ExecutorName2, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
				TryGetItem(oDoc, Markers.ExecutorPosition1, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition);
				TryGetItem(oDoc, Markers.ExecutorPositionName1, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname);
				TryGetItem(oDoc, Markers.ClientName1, contract.ClientSmallName);
				TryGetItem(oDoc, Markers.ExecutorName3, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
				TryGetItem(oDoc, Markers.CoExecutors, tm1.Length > 0 ? tm1.Substring(0, tm1.LastIndexOf(',')) : null);
				TryGetItem(oDoc, Markers.ContractNumber1, contract.Number);
				TryGetItem(oDoc, Markers.ContractDate2, contract.Date);
				TryGetItem(oDoc, Markers.ClientFullNameAndObject, contract.ClientName + ", " + contract.ObjectName);
				TryGetItem(oDoc, Markers.ObjectAddress2, contract.ObjectAddress);
				foreach(ExecutorsInfo item in contract.Coexecutors)
					tm2 += item.Name + ", ";
				TryGetItem(oDoc, Markers.Executors, tm2.Length > 0 ? tm2.Substring(0, tm2.LastIndexOf(',') - 1) : null);
				TryGetItem(oDoc, Markers.ExecutorName4, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
				TryGetItem(oDoc, Markers.ExecutorName5, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Name);
				TryGetItem(oDoc, Markers.LicenseNumber, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicNum);
				TryGetItem(oDoc, Markers.LicenseIssued, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicIssued);
				TryGetItem(oDoc, Markers.LicenseIssuedDate, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicIssuedDate);
				TryGetItem(oDoc, Markers.LicenseExpired, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseExpired);
				TryGetItem(oDoc, Markers.LicenseDelo, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseDelo);
				TryGetItem(oDoc, Markers.LicenseAddress, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseAddress);
				TryGetItem(oDoc, Markers.LicensePhone, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicensePhone);
				TryGetItem(oDoc, Markers.LicensePositionAndNameAndPhone,
					contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition + " " +
					contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname + " " +
					contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPhone);
				TryGetItem(oDoc, Markers.ExecutorPosition2, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).BossPosition);
				TryGetItem(oDoc, Markers.ExecutorPositionName2, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).Bossname);
				TryGetItem(oDoc, Markers.CoExecutors1, tm1.Length > 0 ? tm1.Substring(0, tm1.LastIndexOf(',') - 1) : null);
				TryGetItem(oDoc, Markers.OS, contract.os);
				TryGetItem(oDoc, Markers.PS, contract.ps);
				TryGetItem(oDoc, Markers.TRS, contract.trs);
				TryGetItem(oDoc, Markers.SecurityService, contract.service_security);

				TryGetItem(oDoc, Markers.SendAct, contract.SendActs);
				TryGetItem(oDoc, Markers.OrderAct, contract.PoryadokActs);
				TryGetItem(oDoc, Markers.SignalingOS, contract.SignalingOS);
				TryGetItem(oDoc, Markers.SignalingPS, contract.SignalingPS);
				TryGetItem(oDoc, Markers.SignalingTRS, contract.SignalingTRS);
				TryGetItem(oDoc, Markers.ExecutorReqInfo, contract.Coexecutors.FirstOrDefault(x => x.Id == contract.ExecutorID).LicenseInfo);
				TryGetItem(oDoc, Markers.ClientReqInfo, contract.ClientReqInfo);
				TryGetItem(oDoc, Markers.ArrivalDay, contract.arrivar_day);
				TryGetItem(oDoc, Markers.ArrivalNight, contract.arrivar_night);

				oWord.Visible = true;

				//try {
				//	//var resultPath = currentPath.Remove(currentPath.Length - 4, 4) + "Results";
				//	//oDoc.SaveAsQuickStyleSet(toPathAndName);
				//	oDoc.SaveAs(toPathAndName);
				//	//oDoc.Save();
				//}
				//catch(Exception e) {
				//	//MessageBox.Show(e.Message);
				//	throw new Exception("FATAL ERROR: ДОКУМЕНТ ОТКРЫТ - " + e.InnerException);
				//}

				//ProcessStartInfo processStartInfo = new ProcessStartInfo();
				//processStartInfo.FileName = toPathAndName;
				//Process.Start(processStartInfo);

				//try {
				//	oDoc.Close(false, System.Reflection.Missing.Value, toPathAndName);
				//}
				//catch(Exception CloseException) { }
				//oWord.Quit();
				System.Windows.Forms.Application.Exit();
			}
			catch(Exception e) {
				throw new Exception(e.Message + "\n" + e.ToString());
			}
		}

		//private static void Document_BookmarkLayout(object sender, Spire.Doc.Documents.Rendering.BookmarkLevelEventArgs args) {
		//    throw new NotImplementedException();
		//}
	}
}
