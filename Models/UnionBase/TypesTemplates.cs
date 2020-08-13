using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;

namespace UnionContractWF.Models {
    [Table("TypesTemplates")]
    public class TypesTemplates {
        [Key]
        public Guid ttmp_ID { get; set; }
        public Guid ttmp_ctp_ID { get; set; }
        public string ttmp_tmp { get; set; }
    }
    public class TypesTemplatesContext : DbContext {
        public TypesTemplatesContext() : base("UnionEntity") { }
        public DbSet<TypesTemplates> TypesTemplates { get; set; }
    }
}
