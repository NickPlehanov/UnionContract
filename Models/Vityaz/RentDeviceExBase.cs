using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;

namespace UnionContractWF.Models.Vityaz {
    [Table("New_rent_deviceExtensionBase")]
    public class RentDeviceExBase {
        [Key]
        public Guid? New_rent_deviceId { get; set; }
        public Guid? New_device_rent_device { get; set; }
        public decimal? New_price { get; set; }
        public int? New_qty { get; set; }
        public Guid? New_guard_object_rent_device { get; set; }
        //public DeviceExBase DeviceExBase { get; set; }
    }
    public class RentDeviceExBaseContext : DbContext {
        public RentDeviceExBaseContext() : base("VityazMSCRMEntity") { }
        public DbSet<RentDeviceExBase> RentDeviceExBase { get; set; }
    }
}
