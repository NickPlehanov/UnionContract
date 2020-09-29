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


        private string _Send_acts { get; set; }
        [NotMapped]
        public string Send_acts {
            get => _Send_acts;
            set {
                _Send_acts = int.Parse(value) == 1 ? "Все акты" : "Только акты о доп. оказываемых услугах";
            }
        }
        private string _Poryadok_acts { get; set; }
        [NotMapped]
        public string Poryadok_acts {
            get => _Poryadok_acts;
            set {
                switch (int.Parse(value)) {
                    case 1:
                    _Poryadok_acts = "На адрес электронной почты"; break;
                    case 2:
                    _Poryadok_acts = "На бумажном носителе"; break;
                    case 3:
                    _Poryadok_acts = "В виде электронного документа с цифровой подписью"; break;
                    default:
                    _Poryadok_acts = null; break;
                }
            }
        }

    }
    public class AccountExBaseContext : DbContext {
        public AccountExBaseContext():base("VityazMSCRMEntity") { }
        public DbSet<AccountExBase> AccountExBase { get; set; }
    }
}
