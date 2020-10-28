using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;

namespace UnionContractWF.Models {
    [Table("SystemUserBase")]
    public class SystemUserBase {
        [Key]
        public Guid SystemUserId { get; set; }
        public string FullName { get; set; }
        public string MobilePhone { get; set; }
    }
    public class SystemUserBaseContext : DbContext {
        public SystemUserBaseContext() : base("VityazMSCRMEntity") { }
        public DbSet<SystemUserBase> SystemUserBase { get; set; }
    }
}
