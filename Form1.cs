using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using UnionContractWF.Helpers;
using UnionContractWF.Models;
using UnionContractWF.Models.Vityaz;

namespace UnionContractWF {
    public partial class Form1 : Form {
        public Form1() {
            Args = Environment.GetCommandLineArgs();
            InitializeComponent();
        }

        Database database = new Database();

        private string[] _Args { get; set; }
        public string[] Args {
            get => _Args;
            set {
                _Args = value;
            }
        }

        //private ObservableCollection<ContractTypes> _ContractTypes { get; set; }
        public List<ContractTypes> ContractTypes;

        //private ObservableCollection<TypesTemplates> _TypesTemplates { get; set; }
        public List<TypesTemplates> TypesTemplates;

        private ContractTypes _SelectedContractTypes { get; set; }
        public ContractTypes SelectedContractTypes {
            get => _SelectedContractTypes;
            set {
                _SelectedContractTypes = value;
            }
        }

        private void btn_Union_Click(object sender, EventArgs e) {
            Guid agreementId;
            NumByWords numByWords = new NumByWords();
            if (Guid.TryParse(Args[1], out agreementId)) {
                using (AgreementBaseContext agreementBaseContext = new AgreementBaseContext()) {
                    using (AgreementExBaseContext AgreementExBaseContext = new AgreementExBaseContext()) {
                        using (AccountBaseContext AccountBaseContext = new AccountBaseContext()) {
                            using (AccountExBaseContext AccountExBaseContext = new AccountExBaseContext()) {
                                using (GuardObjectExBaseContext GuardObjectExBaseContext = new GuardObjectExBaseContext()) {
                                    using (RentDeviceExBaseContext RentDeviceExBaseContext = new RentDeviceExBaseContext()) {
                                        var AgreementBase = agreementBaseContext.AgreementBase.FirstOrDefault<AgreementBase>(x => x.New_agreementId == agreementId);
                                        var AgreementExBase = AgreementExBaseContext.AgreementExBase.FirstOrDefault<AgreementExBase>(x => x.New_agreementId == agreementId);
                                        var AccountBase = AccountBaseContext.AccountBase.FirstOrDefault<AccountBase>(x => x.AccountId == AgreementExBase.New_bp_agreement);
                                        var AccountExBase = AccountExBaseContext.AccountExBase.FirstOrDefault(x => x.AccountId == AgreementExBase.New_bp_agreement);
                                        var GuardObjectExBase = GuardObjectExBaseContext.GuardObjectExBase.FirstOrDefault<GuardObjectExBase>(x => x.New_account == AccountExBase.AccountId);
                                        List<RentDeviceExBase> Rent_DeviceExBase = RentDeviceExBaseContext.RentDeviceExBase.Where<RentDeviceExBase>(x => x.New_guard_object_rent_device == GuardObjectExBase.New_guard_objectId).ToList();


                                        //AgreementBase AgreementBase = database.GetDbTable<AgreementBase>(true).FirstOrDefault(x => x.New_agreementId == agreementId);
                                        //AgreementExBase AgreementExBase = database.GetDbTable<AgreementExBase>(true).FirstOrDefault(x => x.New_agreementId == agreementId);
                                        //AccountBase AccountBase = database.GetDbTable<AccountBase>(true).FirstOrDefault(x => x.AccountId == AgreementExBase.New_bp_agreement);
                                        //AccountExBase AccountExBase = database.GetDbTable<AccountExBase>(true).FirstOrDefault(x => x.AccountId == AgreementExBase.New_bp_agreement);
                                        //GuardObjectExBase GuardObjectExBase = database.GetDbTable<GuardObjectExBase>(true).FirstOrDefault(x => x.New_account == AccountExBase.AccountId);
                                        //List<RentDeviceExBase> Rent_DeviceExBase = database.GetDbTable<RentDeviceExBase>(true).Where(x => x.New_guard_object_rent_device == GuardObjectExBase.New_guard_objectId).ToList();

                                        //foreach (RentDeviceExBase item in Rent_DeviceExBase) {
                                        //    if (database.GetDbTable<DeviceExBase>(true).Any(x => x.New_deviceId == item.New_device_rent_device))
                                        //        item.New_Name = database.GetDbTable<DeviceExBase>(true).FirstOrDefault(x => x.New_deviceId == item.New_device_rent_device).New_name;
                                        //}
                                        //соисполнители
                                        List<ExecutorsInfo> executorsInfos = new List<ExecutorsInfo>();
                                        using (ExecutorExBaseContext ExecutorExBaseContext = new ExecutorExBaseContext()) {
                                            List<ExecutorExBase> executors = ExecutorExBaseContext.ExecutorExBase.Where(x => x.New_isCoExecutor == true).ToList<ExecutorExBase>();
                                            //List<ExecutorExBase> executors = database.GetDbTable<ExecutorExBase>(true).Where(x => x.New_isCoExecutor == true).ToList();
                                            foreach (ExecutorExBase item in executors) {
                                                if (item.New_executorId != AgreementExBase.New_executor_agreement)
                                                    executorsInfos.Add(new ExecutorsInfo() {
                                                        Id = item.New_executorId,
                                                        FullInfo = /*item.New_name + Environment.NewLine
                                                            +*/ "Местонахождение исполнительного органа:" + Environment.NewLine
                                                            + item.New_address + Environment.NewLine
                                                            + "Многоканальный телефон " + item.New_phone + Environment.NewLine
                                                            + item.New_info1 + Environment.NewLine
                                                            + "ИНН " + item.New_inn + " КПП " + item.New_kpp + Environment.NewLine
                                                            + "ОГРН " + item.New_ogrn + Environment.NewLine
                                                            + "Р/с " + item.New_bank_rs + Environment.NewLine
                                                            + item.New_bank_name + Environment.NewLine
                                                            + "К/с " + item.New_bank_ks + Environment.NewLine
                                                            + "БИК " + item.New_bank_bik + Environment.NewLine
                                                            + "e-mail " + item.New_Email + Environment.NewLine
                                                            + item.New_Web,
                                                        BossPosition = item.New_boss_name + " " + item.New_name,
                                                        Bossname = item.New_boss_fiio,
                                                        //LicenseInfo = item.New_name + ", лицензия № " + item.New_license_no + ", выдана " + DateTime.Parse(item.New_license_issued_when.ToString()).AddHours(5).ToShortDateString() + " " + item.New_license_issued_who + ", в лице " + item.New_boss_namer + " " + item.New_boss_fior + ", действующего на основании Устава",
                                                        LicenseInfo = "лицензия № " + item.New_license_no + ", выдана " + DateTime.Parse(item.New_license_issued_when.ToString()).AddHours(5).ToShortDateString() + " " + item.New_license_issued_who + ", в лице " + item.New_boss_namer + " " + item.New_boss_fior + ", действующего на основании Устава",
                                                        FullName = item.New_FullName,
                                                        Name = item.New_name,
                                                        LicNum = item.New_license_no,
                                                        LicIssued = item.New_license_issued_who,
                                                        LicIssuedDate = DateTime.Parse(item.New_license_issued_when.ToString()).AddHours(5).ToShortDateString(),
                                                        LicenseExpired = DateTime.Parse(item.New_license_issued_till.ToString()).AddHours(5).ToShortDateString(),
                                                        LicenseAddress = item.New_address,
                                                        LicensePhone = item.New_phone,
                                                        BossPhone = item.New_BossPhone,
                                                        LicenseDelo = item.New_license_case
                                                    });
                                                else
                                                    executorsInfos.Add(new ExecutorsInfo() {
                                                        Id = item.New_executorId,
                                                        FullInfo = /*item.New_name + Environment.NewLine
                                                            +*/ item.New_address + Environment.NewLine
                                                            + "Многоканальный телефон " + item.New_phone + Environment.NewLine
                                                            + item.New_info1 + Environment.NewLine
                                                            + "ИНН " + item.New_inn + " КПП " + item.New_kpp + Environment.NewLine
                                                            + "ОГРН " + item.New_ogrn + Environment.NewLine
                                                            + "Р/с " + item.New_bank_rs + Environment.NewLine
                                                            + item.New_bank_name + Environment.NewLine
                                                            + "К/с " + item.New_bank_ks + Environment.NewLine
                                                            + "БИК " + item.New_bank_bik + Environment.NewLine
                                                            + "e-mail " + item.New_Email + Environment.NewLine
                                                            + item.New_Web,
                                                        BossPosition = item.New_boss_name + " " + item.New_name,
                                                        Bossname = item.New_boss_fiio,
                                                        //LicenseInfo = item.New_FullName + "именуемое в дальнейшем «Исполнитель» , лицензия № " + item.New_license_no + ", выдана " + DateTime.Parse(item.New_license_issued_when.ToString()).AddHours(5).ToShortDateString() + " " + item.New_license_issued_who + ", в лице " + item.New_boss_namer + " " + item.New_boss_fior + ", действующего на основании Устава",                                                       
                                                        LicenseInfo = "лицензия № " + item.New_license_no + ", выдана " + DateTime.Parse(item.New_license_issued_when.ToString()).AddHours(5).ToShortDateString() + " " + item.New_license_issued_who + ", в лице " + item.New_boss_namer + " " + item.New_boss_fior + ", действующего на основании Устава",
                                                        FullName = item.New_FullName,
                                                        Name = item.New_name,
                                                        Lic = "Лицензия № " + item.New_license_no + ", выдана " + item.New_license_issued_who,
                                                        LicDates = "Дата выдачи: " + DateTime.Parse(item.New_license_issued_when.ToString()).AddHours(5).ToShortDateString() + ", сроком по " + DateTime.Parse(item.New_license_issued_till.ToString()).AddHours(5).ToShortDateString(),
                                                        LicCase = item.New_license_case,
                                                        Address = item.New_address + ", тел. " + item.New_phone,
                                                        BossInfo = item.New_boss_name + " " + item.New_boss_fiio + ", тел.:" + item.New_BossPhone,
                                                        BossCompany = item.New_boss_name + " " + item.New_name + " " + "                                                   " + item.New_boss_fiio,
                                                        LicNum = item.New_license_no,
                                                        LicIssued = item.New_license_issued_who,
                                                        LicIssuedDate = DateTime.Parse(item.New_license_issued_when.ToString()).AddHours(5).ToShortDateString(),
                                                        LicenseExpired = DateTime.Parse(item.New_license_issued_till.ToString()).AddHours(5).ToShortDateString(),
                                                        LicenseAddress = item.New_address,
                                                        LicensePhone = item.New_phone,
                                                        BossPhone = item.New_BossPhone,
                                                        LicenseDelo = item.New_license_case
                                                    });
                                            }
                                        }
                                        //основные поля договора
                                        Contract contract = new Contract();
                                        contract.ExecutorID = AgreementExBase.New_executor_agreement;
                                        contract.Number = AgreementExBase.New_number.ToString();
                                        contract.Date = DateTime.Parse(AgreementExBase.New_date.ToString()).AddHours(5).ToShortDateString();
                                        contract.ClientName = AccountBase.Name;
                                        contract.ClientSmallName = AccountExBase.new_smallname;
                                        contract.ObjectName = GuardObjectExBase.New_name;
                                        contract.ObjectAddress = GuardObjectExBase.New_addr_kladr;
                                        if (GuardObjectExBase.New_cost.HasValue) {
                                            contract.ObjectCost = GuardObjectExBase.New_cost.Value.ToString("F0", CultureInfo.CurrentCulture) + Environment.NewLine;
                                            contract.ObjectCost += "(" + numByWords.NumPhrase(ulong.Parse(GuardObjectExBase.New_cost.Value.ToString("F0", CultureInfo.CurrentCulture)), true) + ")";
                                        }
                                        else
                                            contract.ObjectCost = string.Empty;
                                        contract.ObjectMonthlyPay = String.Format("{0,-10:F}", GuardObjectExBase.New_monthlypay.ToString()) + Environment.NewLine;
                                        contract.ObjectMonthlyPay += "(" + numByWords.NumPhrase(ulong.Parse(String.Format("{0,-10:F}", GuardObjectExBase.New_monthlypay.ToString())), true) + ")";
                                        contract.SendActs = AccountExBase.Send_acts;
                                        contract.PoryadokActs = AccountExBase.Poryadok_acts;
                                        if (AccountExBase.New_agent_type == 1) {
                                            string pass_date = DateTime.TryParse(AccountExBase.New_pass_date.ToString(), out _) ? DateTime.Parse(AccountExBase.New_pass_date.ToString()).AddHours(5).ToShortDateString() + Environment.NewLine : "" + Environment.NewLine;
                                            contract.ClientInfo = AccountBase.Name + Environment.NewLine
                                                + "ИНН " + AccountExBase.New_req_inn + Environment.NewLine
                                                + "Адрес регистрации: " + AccountExBase.New_fact_addr_kladr + Environment.NewLine
                                                + "Почтовый адрес: " + AccountExBase.New_fact_addr_kladr + Environment.NewLine
                                                + "Серия: " + AccountExBase.New_pass_serial + " Номер: " + AccountExBase.New_pass_number + Environment.NewLine
                                                + "Выдан: " + AccountExBase.New_pass_issued + Environment.NewLine
                                                + "Дата выдачи: " + pass_date
                                                + "e-mail: " + AccountBase.EMailAddress1;
                                        }
                                        else
                                            contract.ClientInfo = AccountBase.Name + Environment.NewLine
                                                + "ИНН " + AccountExBase.New_req_inn + Environment.NewLine
                                                + "Адрес регистрации: " + AccountExBase.New_fact_addr_kladr + Environment.NewLine
                                                + "Почтовый адрес: " + AccountExBase.New_fact_addr_kladr + Environment.NewLine
                                                //+ "Серия: " + AccountExBase.New_pass_serial + " Номер: " + AccountExBase.New_pass_number + Environment.NewLine
                                                //+ "Выдан: " + AccountExBase.New_pass_issued + Environment.NewLine
                                                //+ "Дата выдачи: " + DateTime.Parse(AccountExBase.New_pass_date.ToString()).AddHours(5).ToShortDateString() + Environment.NewLine
                                                + "e-mail: " + AccountBase.EMailAddress1;
                                        using (SystemUserBaseContext systemUserBaseContext = new SystemUserBaseContext()) {
                                            contract.OwningUser = systemUserBaseContext.SystemUserBase.FirstOrDefault(x => x.SystemUserId == AgreementBase.OwningUser).FullName;
                                            //contract.OwningUser = database.GetDbTable<SystemUserBase>(true).FirstOrDefault(x => x.SystemUserId == AgreementBase.OwningUser).FullName;
                                        }
                                        if (bool.Parse(GuardObjectExBase.New_protection_os.ToString())) {
                                            contract.ObjectTypeService = "охрана имущества" + Environment.NewLine;
                                            contract.ObjectSignalization = "Охранная сигнализация" + Environment.NewLine;
                                            contract.os = "■ охрана имущества" + Environment.NewLine;
                                        }
                                        else
                                            //contract.os = "□ охрана имущества" + Environment.NewLine;                                            
                                            contract.os = "";
                                        if (bool.Parse(GuardObjectExBase.New_protection_ps.ToString())) {
                                            contract.ObjectTypeService += "реагирование на сигналы пожар" + Environment.NewLine;
                                            contract.ObjectSignalization = "Пожарная сигнализация" + Environment.NewLine;
                                            contract.ps = "■ реагирование на сигналы пожар" + Environment.NewLine;
                                        }
                                        else
                                            //contract.ps = "□ реагирование на сигналы пожар" + Environment.NewLine;
                                            contract.ps = "";
                                        if (bool.Parse(GuardObjectExBase.New_protection_trs.ToString())) {
                                            contract.ObjectTypeService += "экстренный выезд по тревожной сигнализации" + Environment.NewLine;
                                            contract.ObjectSignalization = "Тревожная сигнализация" + Environment.NewLine;
                                            contract.trs = "■ экстренный выезд по тревожной сигнализации" + Environment.NewLine;
                                        }
                                        else
                                            //contract.trs = "□ экстренный выезд по тревожной сигнализации" + Environment.NewLine;
                                            contract.trs = "";
                                        contract.ObjectTypeService += "эксплуатационное обслуживание средств сигнализации" + Environment.NewLine;
                                        contract.service_security = "■ эксплуатационное обслуживание средств сигнализации" + Environment.NewLine;
                                        //if (bool.Parse(GuardObjectExBase.New_protection_os.ToString()) && bool.Parse(GuardObjectExBase.New_protection_ps.ToString()) && bool.Parse(GuardObjectExBase.New_protection_trs.ToString()))
                                        //    contract.BlockType = "полной";
                                        //else
                                        //    contract.BlockType = "усеченной";
                                        if (GuardObjectExBase.New_dogovor_type == 1) {//стандартный тип договора
                                            contract.BlockTypeFull = "■ c полной блокировкой";
                                            contract.BlockTypeDistinct = "□ c усеченной блокировкой";
                                        }
                                        else if (GuardObjectExBase.New_dogovor_type == 2) {
                                            contract.BlockTypeDistinct = "■ c усеченной блокировкой";
                                            contract.BlockTypeFull = "□ c полной блокировкой";
                                        }
                                        else {
                                            contract.BlockTypeFull = "□ c полной блокировкой";
                                            contract.BlockTypeDistinct = "□ c усеченной блокировкой";
                                        }
                                        decimal _sum = 0;
                                        int c = 1;
                                        using (DeviceExBaseContext deviceExBaseContext = new DeviceExBaseContext()) {
                                            foreach (RentDeviceExBase item in Rent_DeviceExBase) {
                                                //contract.DeviceName += item.DeviceExBase.New_name + Environment.NewLine;
                                                contract.DeviceName += c.ToString()+". " +deviceExBaseContext.DeviceExBase.FirstOrDefault(x => x.New_deviceId == item.New_device_rent_device).New_name + Environment.NewLine;
                                                contract.DeviceCount += item.New_qty + Environment.NewLine;
                                                if (item.New_price.HasValue)
                                                    contract.DeviceSum += (item.New_qty * Decimal.Round(decimal.Parse(item.New_price.ToString()), 2)) + Environment.NewLine;
                                                else
                                                    contract.DeviceSum += Environment.NewLine;
                                                if (!string.IsNullOrEmpty(item.New_price.ToString()) && !string.IsNullOrEmpty(item.New_qty.ToString()))
                                                    _sum += Decimal.Round((decimal)(item.New_qty * item.New_price), 2);
                                                c++;
                                            }
                                        }
                                        contract.AllCount = Rent_DeviceExBase.Count.ToString();
                                        contract.AllSum = _sum.ToString();
                                        contract.Coexecutors = executorsInfos;
                                        string tmp = TypesTemplates.FirstOrDefault(x => x.ttmp_ctp_ID == SelectedContractTypes.ctp_ID).ttmp_tmp;
                                        //WordDocument.Exchange(contract, tmp, PathToSave);                    
                                        WordDocument.Exchange(contract, tmp, @"\\server-nass\Install\ИСХОДНИКИ\Шаблоны договоров\tmp\" + contract.Number + ".docx");
                                        //WordDocument.Exchange(contract, tmp, contract.Number + ".docx");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e) {
            using (ContractTypesContext contractTypesContext = new ContractTypesContext()) {
                ContractTypes = contractTypesContext.ContractTypes.Where(x => x.ctp_IsActive == true).ToList<ContractTypes>();
                cmb_ContractType.DataSource = ContractTypes;
                cmb_ContractType.DisplayMember = "ctp_Name";
            }
            using (TypesTemplatesContext typesTemplatesContext = new TypesTemplatesContext()) {
                TypesTemplates = typesTemplatesContext.TypesTemplates.ToList<TypesTemplates>();
            }
        }

        private void cmb_ContractType_SelectedIndexChanged(object sender, EventArgs e) {
            SelectedContractTypes = (ContractTypes)cmb_ContractType.SelectedItem;
        }

        private void Form1_Activated(object sender, EventArgs e) {
            //ContractTypes = new ObservableCollection<ContractTypes>(database.GetDbTable<ContractTypes>(false));
            //TypesTemplates = new ObservableCollection<TypesTemplates>(database.GetDbTable<TypesTemplates>(false));
        }

        private void button1_Click(object sender, EventArgs e) {
            //ContractTypes = new ObservableCollection<ContractTypes>(database.GetDbTable<ContractTypes>(false));
            //TypesTemplates = new ObservableCollection<TypesTemplates>(database.GetDbTable<TypesTemplates>(false));
            //cmb_ContractType.DataSource = ContractTypes;
            //cmb_ContractType.DisplayMember = "ctp_Name";
        }
    }
}
