using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;

namespace UnionContractWF.Models.Vityaz {
    [Table("New_agreementExtensionBase")]
    public class AgreementExBase {
        [Key]
        public Guid? New_agreementId { get; set; }
        public int? New_number { get; set; }
        public DateTime? New_date { get; set; }
        public Guid? New_bp_agreement { get; set; }
        public Guid? New_executor_agreement { get; set; }
        public Guid? New_dogovor_type_agreement { get; set; }
    }
    public class AgreementExBaseContext : DbContext {
        public AgreementExBaseContext() : base("VityazMSCRMEntity") { }
        public DbSet<AgreementExBase> AgreementExBase { get; set; }
    }
}
