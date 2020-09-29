using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;

namespace UnionContractWF.Models.Vityaz
{
    [Table("new_new_agreement_new_guard_objectBase")]
    public class AgreementGuardObject
    {
        [Key]
        public Guid new_new_agreement_new_guard_objectId { get; set; }
        public Guid new_agreementid { get; set; }
        public Guid new_guard_objectid { get; set; }
    }
    public class AgreementGuardObjectContext : DbContext
    {
        public AgreementGuardObjectContext() : base("VityazMSCRMEntity") { }
        public DbSet<AgreementGuardObject> AgreementGuardObject { get; set; }
    }
}
