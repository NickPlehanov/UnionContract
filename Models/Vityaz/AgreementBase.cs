using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;

namespace UnionContractWF.Models.Vityaz {
    [Table("New_agreementBase")]
    public class AgreementBase {
        [Key]
        public Guid New_agreementId { get; set; }
        public int statecode { get; set; }
        public int statuscode { get; set; }
        public Guid OwningUser { get; set; }
    }
    public class AgreementBaseContext : DbContext {
        public AgreementBaseContext() : base("VityazMSCRMEntity") { }
        public DbSet<AgreementBase> AgreementBase { get; set; }
    }
}
