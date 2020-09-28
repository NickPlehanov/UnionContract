using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;

namespace UnionContractWF.Models.Vityaz {
    [Table("New_dogovor_typeExtensionBase")]
    public class DogovorTypeExBase {
        [Key]
        public Guid New_dogovor_typeId { get; set; }
        public string New_name { get; set; }
        public bool? New_techService { get; set; }
    }
    public class DogovorTypeExBaseContext : DbContext {
        public DogovorTypeExBaseContext() : base("VityazMSCRMEntity") { }
        public DbSet<DogovorTypeExBase> DogovorTypeExBase { get; set; }
    }
}
