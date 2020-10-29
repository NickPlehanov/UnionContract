using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;

namespace UnionContractWF.Models.Vityaz {
	[Table("AccountExtensionBase")]
	public class AccountExBase {
		[Key]
		public Guid? AccountId { get; set; }
		public string New_req_inn { get; set; }
		public string New_fact_addr_kladr { get; set; }
		public string New_pass_serial { get; set; }
		public string New_pass_number { get; set; }
		public string New_pass_issued { get; set; }
		public DateTime? New_pass_date { get; set; }
		public int? New_agent_type { get; set; }
		public int? New_send_acts { get; set; }
		public int? new_poryadok_acts { get; set; }
		public string new_smallname { get; set; }
		/// <summary>
		/// должность директора
		/// </summary>
		public string New_req_status { get; set; }
		//     [NotMapped]
		//     public string New_req_status_rodit {
		//         get => New_req_status.ToLower() + "а";
		//set {
		//             if(value != null)
		//                 New_req_status = New_req_status.ToLower() + "а";
		//}
		//     }
		/// <summary>
		/// на основании чего работает директор
		/// </summary>
		public string New_found { get; set; }
		/// <summary>
		/// ФИО директора
		/// </summary>
		public string New_req_boss_fio { get; set; }

		public string New_req_ogrn { get; set; }
		public string New_req_kpp { get; set; }
		public string New_bank_name { get; set; }
		public string New_bank_bik { get; set; }
		public string New_bank_rs { get; set; }
		public string New_bank_ks { get; set; }


		[NotMapped]
		public string Send_acts {
			get {
				if(New_send_acts.HasValue)
					return New_send_acts.Value == 1 ? "Все акты" : "Только акты о доп. оказываемых услугах";
				else
					return null;
			}
		}
		//private string _Poryadok_acts { get; set; }
		[NotMapped]
		public string Poryadok_acts {
			get {
				if(new_poryadok_acts.HasValue)
					switch(new_poryadok_acts.Value) {
						case 1:
							return "На адрес электронной почты";
						case 2:
							return "На бумажном носителе";
						case 3:
							return "В виде электронного документа с цифровой подписью";
						default:
							return null;
					}
				else
					return null;
			}
		}

	}
	public class AccountExBaseContext : DbContext {
		public AccountExBaseContext() : base("VityazMSCRMEntity") { }
		public DbSet<AccountExBase> AccountExBase { get; set; }
	}
}
