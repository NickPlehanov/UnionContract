using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Windows.Forms;

namespace UnionContractWF.Models.Vityaz {
    [Table("New_guard_objectExtensionBase")]
    public class GuardObjectExBase {
        [Key]
        public Guid? New_guard_objectId { get; set; }
        public string New_name { get; set; }
        public string New_addr_kladr { get; set; }
        public decimal? New_cost { get; set; }
        public bool? New_protection_os { get; set; }
        public bool? New_protection_ps { get; set; }
        public bool? New_protection_trs { get; set; }
        public string New_monthlypay { get; set; }
        public Guid? New_account { get; set; }
        public int? New_dogovor_type { get; set; }
    }
    public class GuardObjectExBaseContext : DbContext {
        public GuardObjectExBaseContext() : base("VityazMSCRMEntity") { }
        public DbSet<GuardObjectExBase> GuardObjectExBase { get; set; }
    }
}
