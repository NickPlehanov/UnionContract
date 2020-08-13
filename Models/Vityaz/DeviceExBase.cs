using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;

namespace UnionContractWF.Models.Vityaz {
    [Table("New_deviceExtensionBase")]
    public class DeviceExBase {
        [Key]
        public Guid? New_deviceId { get; set; }
        public string New_name { get; set; }
    }
    public class DeviceExBaseContext : DbContext {
        public DeviceExBaseContext() : base("VityazMSCRMEntity") { }
        public DbSet<DeviceExBase> DeviceExBase { get; set; }
    }
}
