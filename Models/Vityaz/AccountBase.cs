using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;

namespace UnionContractWF.Models.Vityaz {
    [Table("AccountBase")]
    public class AccountBase {
        [Key]
        public Guid AccountId { get; set; }
        public string Name { get; set; }
        public string EMailAddress1 { get; set; }
    }
    public class AccountBaseContext : DbContext {
        public AccountBaseContext() : base("VityazMSCRMEntity") { }
        public DbSet<AccountBase> AccountBase { get; set; }
    }
}
