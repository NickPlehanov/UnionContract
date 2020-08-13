using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;

namespace UnionContractWF.Models.Vityaz {
    [Table("New_executorExtensionBase")]
    public class ExecutorExBase {
        [Key]
        public Guid New_executorId { get; set; }
        public string New_name { get; set; }
        public string New_address { get; set; }
        public string New_inn { get; set; }
        public string New_kpp { get; set; }
        public string New_ogrn { get; set; }
        public string New_bank_name { get; set; }
        public string New_bank_rs { get; set; }
        public string New_bank_ks { get; set; }
        public string New_bank_bik { get; set; }
        public string New_boss_name { get; set; }
        public string New_boss_fiio { get; set; }
        public string New_boss_namer { get; set; }
        public string New_boss_fior { get; set; }
        public string New_phone { get; set; }
        public string New_Email { get; set; }
        public string New_Web { get; set; }
        public string New_info1 { get; set; }
        public string New_info2 { get; set; }
        public string New_license_no { get; set; }
        public DateTime? New_license_issued_when { get; set; }
        public string New_license_issued_who { get; set; }
        public DateTime? New_license_issued_till { get; set; }
        public string New_license_case { get; set; }
        public bool? New_isCoExecutor { get; set; }
        public string New_FullName { get; set; }
        public string New_BossPhone { get; set; }
    }
    public class ExecutorExBaseContext : DbContext {
        public ExecutorExBaseContext() : base("VityazMSCRMEntity") { }
        public DbSet<ExecutorExBase> ExecutorExBase { get; set; }
    }
}
