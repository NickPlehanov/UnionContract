using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;

namespace UnionContractWF.Models {
    [Table("ContractTypes")]
    public class ContractTypes {
        [Key]
        public Guid ctp_ID { get; set; }
        public string ctp_Name { get; set; }
        public bool ctp_IsActive { get; set; }
    }
    public class ContractTypesContext : DbContext {
        public ContractTypesContext() : base("UnionEntity") { }
        public DbSet<ContractTypes> ContractTypes { get; set; }
    }
}
