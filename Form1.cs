using NPetrovich;
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
			bool? ts = false;
			NumByWords numByWords = new NumByWords();
			if(Guid.TryParse(Args[1], out agreementId)) {
				using(AgreementBaseContext agreementBaseContext = new AgreementBaseContext()) {
					using(AgreementExBaseContext AgreementExBaseContext = new AgreementExBaseContext()) {
						using(AccountBaseContext AccountBaseContext = new AccountBaseContext()) {
							using(AccountExBaseContext AccountExBaseContext = new AccountExBaseContext()) {
								using(AgreementGuardObjectContext agreementGuardObjectContext = new AgreementGuardObjectContext()) {
									using(GuardObjectExBaseContext GuardObjectExBaseContext = new GuardObjectExBaseContext()) {
										using(RentDeviceExBaseContext RentDeviceExBaseContext = new RentDeviceExBaseContext()) {
											var AgreementBase = agreementBaseContext.AgreementBase.FirstOrDefault<AgreementBase>(x => x.New_agreementId == agreementId);
											var AgreementExBase = AgreementExBaseContext.AgreementExBase.FirstOrDefault<AgreementExBase>(x => x.New_agreementId == agreementId);
											var AgreementGuardObject = agreementGuardObjectContext.AgreementGuardObject.FirstOrDefault(x => x.new_agreementid == agreementId);
											var AccountBase = AccountBaseContext.AccountBase.FirstOrDefault<AccountBase>(x => x.AccountId == AgreementExBase.New_bp_agreement);
											var AccountExBase = AccountExBaseContext.AccountExBase.FirstOrDefault(x => x.AccountId == AgreementExBase.New_bp_agreement);
											//var GuardObjectExBase = GuardObjectExBaseContext.GuardObjectExBase.FirstOrDefault<GuardObjectExBase>(x => x.New_account == AccountExBase.AccountId);
											var GuardObjectExBase = GuardObjectExBaseContext.GuardObjectExBase.FirstOrDefault<GuardObjectExBase>(x => x.New_guard_objectId == AgreementGuardObject.new_guard_objectid);
											List<RentDeviceExBase> Rent_DeviceExBase = RentDeviceExBaseContext.RentDeviceExBase.Where<RentDeviceExBase>(x => x.New_guard_object_rent_device == GuardObjectExBase.New_guard_objectId).ToList();
											using(DogovorTypeExBaseContext dogovorTypeExBaseContext = new DogovorTypeExBaseContext()) {
												ts = dogovorTypeExBaseContext.DogovorTypeExBase.FirstOrDefault(x => x.New_dogovor_typeId == AgreementExBase.New_dogovor_type_agreement).New_techService;
												ts = ts.HasValue ? bool.TryParse(ts.ToString(), out _) ? bool.Parse(ts.ToString()) : false : false;
												//ts = dogovorTypeExBaseContext.DogovorTypeExBase.FirstOrDefault(x => x.New_dogovor_typeId == AgreementExBase.New_dogovor_type_agreement).New_techService.HasValue ? bool.TryParse(ts.ToString(), out _) ? bool.Parse(ts.ToString()) : false : false;
											}
											//соисполнители
											List<ExecutorsInfo> executorsInfos = new List<ExecutorsInfo>();
											using(ExecutorExBaseContext ExecutorExBaseContext = new ExecutorExBaseContext()) {
												if((bool)ts) {//если у нас договор на ТО
													List<ExecutorExBase> executors = ExecutorExBaseContext.ExecutorExBase.Where(x => x.New_executorId == AgreementExBase.New_executor_agreement).ToList<ExecutorExBase>();
													foreach(ExecutorExBase item in executors) {
														executorsInfos.Add(new ExecutorsInfo() {
															Id = item.New_executorId,
															FullInfo = /*item.New_name + Environment.NewLine
                                                            +*/ "Местонахождение исполнительного органа:" + Environment.NewLine
																	+ item.New_address + Environment.NewLine
																	+ "Многоканальный телефон " + item.New_phone + Environment.NewLine
																	//+ item.New_info1 + Environment.NewLine
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
															//LicenseInfo = "лицензия № " + item.New_license_no + ", выдана " + DateTime.Parse(item.New_license_issued_when.ToString()).AddHours(5).ToShortDateString() + " " + item.New_license_issued_who + ", в лице " + item.New_boss_namer + " " + item.New_boss_fior + ", действующего на основании Устава",
															LicenseInfo = "в лице " + item.New_boss_namer + " " + item.New_boss_fior + ", действующего на основании Устава, лицензии № " + item.New_license_no + ", выданной " + DateTime.Parse(item.New_license_issued_when.ToString()).AddHours(5).ToShortDateString() + " " + item.New_license_issued_who,
															FullName = item.New_FullName,
															Name = item.New_name,
															LicNum = item.New_license_no,
															LicIssued = item.New_license_issued_who,
															LicIssuedDate = DateTime.Parse(item.New_license_issued_when.ToString()).AddHours(5).ToShortDateString(),
															//LicenseExpired = DateTime.Parse(item.New_license_issued_till.ToString()).AddHours(5).ToShortDateString(),
															LicenseAddress = item.New_address,
															LicensePhone = item.New_phone,
															BossPhone = item.New_BossPhone,
															LicenseDelo = item.New_license_case
														});
													}
												}
												else {
													List<ExecutorExBase> executors = ExecutorExBaseContext.ExecutorExBase.Where(x => x.New_isCoExecutor == true).ToList<ExecutorExBase>();
													foreach(ExecutorExBase item in executors) {
														if(item.New_executorId != AgreementExBase.New_executor_agreement)
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
											}
											//основные поля договора
											Contract contract = new Contract();
											if(AccountExBase.New_agent_type == 2) //юр.лицо
												contract.ClientReqInfo = "в лице "+ GetGenitive(AccountExBase.New_req_status + " " + AccountExBase.New_req_boss_fio) + ", действующего на основании " + AccountExBase.New_found;
											else
												contract.ClientReqInfo = "";
											//contract.ClientReqInfo = "в лице " + AccountExBase.New_req_status + " " + AccountExBase.New_req_boss_fio + ", действующего на основании " + AccountExBase.New_found;
											contract.ExecutorID = AgreementExBase.New_executor_agreement;
											contract.Number = AgreementExBase.New_number.ToString();
											//contract.Date = DateTime.Parse(AgreementExBase.New_date.ToString()).AddHours(5).ToShortDateString(); 
											if(DateTime.TryParse(AgreementExBase.New_date.ToString(), out _))
												contract.Date = DateTime.Parse(AgreementExBase.New_date.ToString()).AddHours(5).ToShortDateString();
											contract.ClientName = AccountBase.Name;
											if(AccountExBase.New_agent_type == 1 || AccountExBase.New_agent_type == 3)//физ. лицо
												contract.ClientSmallName = AccountExBase.new_smallname;
											//contract.ClientSmallName = GetSmallName(AccountExBase.new_smallname);
											else if(AccountExBase.New_agent_type == 2 )//юр лицо
												contract.ClientSmallName = GetSmallName(AccountExBase.New_req_boss_fio);
											contract.ObjectName = GuardObjectExBase.New_name;
											contract.ObjectAddress = GuardObjectExBase.New_addr_kladr;
											if(GuardObjectExBase.New_cost.HasValue) {
												contract.ObjectCost = GuardObjectExBase.New_cost.Value.ToString("F0", CultureInfo.CurrentCulture) + Environment.NewLine;
												contract.ObjectCost += "(" + numByWords.NumPhrase(ulong.Parse(GuardObjectExBase.New_cost.Value.ToString("F0", CultureInfo.CurrentCulture)), true) + ")";
											}
											else
												contract.ObjectCost = string.Empty;
											contract.ObjectMonthlyPay = String.Format("{0,-10:F}", GuardObjectExBase.New_monthlypay.ToString()).Replace(" ", "")/* + Environment.NewLine*/;
											contract.ObjectMonthlyPay += "(" + numByWords.NumPhrase(ulong.Parse(String.Format("{0,-10:F}", GuardObjectExBase.New_monthlypay.ToString())), true) + ")";
											contract.SendActs = AccountExBase.Send_acts;
											contract.PoryadokActs = AccountExBase.Poryadok_acts;
											if(AccountExBase.New_agent_type == 1) //физ.лицо
											{
												string pass_date = DateTime.TryParse(AccountExBase.New_pass_date.ToString(), out _) ? DateTime.Parse(AccountExBase.New_pass_date.ToString()).AddHours(5).ToShortDateString() + Environment.NewLine : "" + Environment.NewLine;
												contract.ClientInfo = AccountBase.Name + Environment.NewLine
													+ "ИНН " + AccountExBase.New_req_inn + Environment.NewLine
													+ "Адрес регистрации: " + AccountExBase.New_fact_addr_kladr + Environment.NewLine
													+ "Почтовый адрес: " + AccountExBase.New_fact_addr_kladr + Environment.NewLine
													+ "Паспорт Серия: " + AccountExBase.New_pass_serial + " Номер: " + AccountExBase.New_pass_number + Environment.NewLine
													+ "Выдан: " + AccountExBase.New_pass_issued + Environment.NewLine
													+ "Дата выдачи: " + pass_date
													+ "e-mail: " + AccountBase.EMailAddress1;
												contract.arrivar_day = GuardObjectExBase.New_arrival_time_day ?? "3-15";
												contract.arrivar_night = GuardObjectExBase.New_arrival_time_night ?? "3-12";
											}
											else if(AccountExBase.New_agent_type == 3) {//ИП
												string pass_date = DateTime.TryParse(AccountExBase.New_pass_date.ToString(), out _) ? DateTime.Parse(AccountExBase.New_pass_date.ToString()).AddHours(5).ToShortDateString() + Environment.NewLine : "" + Environment.NewLine;
												contract.ClientInfo = AccountBase.Name + Environment.NewLine
													+ "ИНН " + AccountExBase.New_req_inn + Environment.NewLine
													+ "ОГРН: " + AccountExBase.New_req_ogrn + Environment.NewLine
													+ "Адрес регистрации: " + AccountExBase.New_fact_addr_kladr + Environment.NewLine
													+ "Почтовый адрес: " + AccountExBase.New_fact_addr_kladr + Environment.NewLine
													+ "Паспорт Серия: " + AccountExBase.New_pass_serial + " Номер: " + AccountExBase.New_pass_number + Environment.NewLine
													+ "Выдан: " + AccountExBase.New_pass_issued + Environment.NewLine
													+ "Дата выдачи: " + pass_date
													+ "e-mail: " + AccountBase.EMailAddress1;
												contract.arrivar_day = GuardObjectExBase.New_arrival_time_day ?? "3-15";
												contract.arrivar_night = GuardObjectExBase.New_arrival_time_night ?? "3-12";
											}
											else if(AccountExBase.New_agent_type == 2 )//юр.лицо
											{
												string pass_date = DateTime.TryParse(AccountExBase.New_pass_date.ToString(), out _) ? DateTime.Parse(AccountExBase.New_pass_date.ToString()).AddHours(5).ToShortDateString() + Environment.NewLine : "" + Environment.NewLine;
												contract.ClientInfo = /*AccountBase.Name + Environment.NewLine
													+*/ "ИНН " + AccountExBase.New_req_inn + " КПП " + AccountExBase.New_req_kpp + Environment.NewLine
													+ "ОГРН: " + AccountExBase.New_req_ogrn + Environment.NewLine
													+ "Р/с: " + AccountExBase.New_bank_rs + Environment.NewLine
													+ AccountExBase.New_bank_name + Environment.NewLine
													+ "К/с: " + AccountExBase.New_bank_ks + Environment.NewLine
													+ "БИК: " + AccountExBase.New_bank_bik + Environment.NewLine
													+ "Юридический адрес: " + AccountExBase.New_fact_addr_kladr + Environment.NewLine
													+ "Почтовый адрес: " + AccountExBase.New_fact_addr_kladr + Environment.NewLine
													+ "e-mail: " + AccountBase.EMailAddress1 + Environment.NewLine
													+ AccountExBase.New_req_status + Environment.NewLine
													+ AccountExBase.new_smallname;
												contract.arrivar_day = GuardObjectExBase.New_arrival_time_day ?? "3-12";
												contract.arrivar_night = GuardObjectExBase.New_arrival_time_night ?? "3-9";
											}
											else
												contract.ClientInfo = AccountBase.Name + Environment.NewLine
													+ "ИНН " + AccountExBase.New_req_inn + Environment.NewLine
													+ "Адрес регистрации: " + AccountExBase.New_fact_addr_kladr + Environment.NewLine
													+ "Почтовый адрес: " + AccountExBase.New_fact_addr_kladr + Environment.NewLine
													+ "e-mail: " + AccountBase.EMailAddress1;
											using(SystemUserBaseContext systemUserBaseContext = new SystemUserBaseContext()) {
												contract.OwningUser = systemUserBaseContext.SystemUserBase.FirstOrDefault(x => x.SystemUserId == AgreementBase.OwningUser).FullName+" "+
													systemUserBaseContext.SystemUserBase.FirstOrDefault(x => x.SystemUserId == AgreementBase.OwningUser).MobilePhone;
											}
											if(bool.Parse(GuardObjectExBase.New_protection_os.ToString())) {
												contract.ObjectTypeService += "охрана имущества" + Environment.NewLine;
												contract.os = "■ охрана имущества" + Environment.NewLine;
												contract.SignalingOS = "Охранная сигнализация" + Environment.NewLine;
											}
											else
												contract.os = "";
											if(bool.Parse(GuardObjectExBase.New_protection_ps.ToString())) {
												contract.ObjectTypeService += "реагирование на сигналы пожар" + Environment.NewLine;
												contract.ps = "■ реагирование на сигналы пожар" + Environment.NewLine;
												contract.SignalingPS = "Пожарная сигнализация" + Environment.NewLine;
											}
											else
												contract.ps = "";
											if(bool.Parse(GuardObjectExBase.New_protection_trs.ToString())) {
												contract.ObjectTypeService += "экстренный выезд по тревожной сигнализации" + Environment.NewLine;
												contract.trs = "■ экстренный выезд по тревожной сигнализации" + Environment.NewLine;
												contract.SignalingTRS = "Тревожная сигнализация" + Environment.NewLine;
											}
											else
												contract.trs = "";
											contract.ObjectTypeService += "эксплуатационное обслуживание средств сигнализации (за исключением пожарной сигнализации)" + Environment.NewLine;
											contract.service_security = "■ эксплуатационное обслуживание средств сигнализации (за исключением пожарной сигнализации)" + Environment.NewLine;
											if(GuardObjectExBase.New_dogovor_type == 1) {//стандартный тип договора
												contract.BlockTypeFull = "■ c полной блокировкой";
												contract.BlockTypeDistinct = "□ c усеченной блокировкой";
											}
											else if(GuardObjectExBase.New_dogovor_type == 2) {
												contract.BlockTypeDistinct = "■ c усеченной блокировкой";
												contract.BlockTypeFull = "□ c полной блокировкой";
											}
											else {
												contract.BlockTypeFull = "□ c полной блокировкой";
												contract.BlockTypeDistinct = "□ c усеченной блокировкой";
											}
											decimal _sum = 0;
											int c = 1;
											using(DeviceExBaseContext deviceExBaseContext = new DeviceExBaseContext()) {
												foreach(RentDeviceExBase item in Rent_DeviceExBase.Where(x=>x.New_device_rent_device!=null)) {
													contract.DeviceName += c.ToString() + ". " + deviceExBaseContext.DeviceExBase.FirstOrDefault(x => x.New_deviceId == item.New_device_rent_device).New_name + Environment.NewLine;
													contract.DeviceCount += item.New_qty + Environment.NewLine;
													if(item.New_price.HasValue)
														contract.DeviceSum += (item.New_qty * Decimal.Round(decimal.Parse(item.New_price.ToString()), 2)) + Environment.NewLine;
													else
														contract.DeviceSum += Environment.NewLine;
													if(!string.IsNullOrEmpty(item.New_price.ToString()) && !string.IsNullOrEmpty(item.New_qty.ToString()))
														_sum += Decimal.Round((decimal)(item.New_qty * item.New_price), 2);
													c++;
												}
											}
											contract.AllCount = Rent_DeviceExBase.Count.ToString();
											contract.AllSum = _sum.ToString();
											contract.Coexecutors = executorsInfos;
											contract.PositionAndSmallName= AccountExBase.New_req_status + Environment.NewLine + AccountExBase.new_smallname;
											contract.SmallFirmName = AccountExBase.new_smallname;
											string tmp = TypesTemplates.FirstOrDefault(x => x.ttmp_ctp_ID == SelectedContractTypes.ctp_ID).ttmp_tmp;
											WordDocument.Exchange(contract, tmp, @"\\server-nass\Install\ИСХОДНИКИ\Шаблоны договоров\tmp\" + contract.Number + " " + contract.ClientName + ".docx");
										}
									}
								}
							}
						}
					}
				}
			}
		}

		public string GetGenitive(string str) {
			if(string.IsNullOrEmpty(str))
				return null;
			else {
				string[] words = new string[10];
				words = str.Split(' ');
				if(words.Length == 4) {
					var Genetive = new Petrovich() {
						AutoDetectGender = true,
						LastName = words[1],
						FirstName = words[2],
						MiddleName = words[3]
					};
					var inflected = Genetive.InflectTo(Case.Genitive);

					var GenetivePosition = new Petrovich() {
						AutoDetectGender = true,
						FirstName = words[0].ToLower()
					};
					return GenetivePosition.InflectFirstNameTo(Case.Genitive) + " " + inflected.LastName + " " + inflected.FirstName + " " + inflected.MiddleName;
				}
				else
					return null;
			}
		}

		public string GetSmallName(string FullName) {
			string ret = null;
			if(string.IsNullOrEmpty(FullName))
				return FullName;
			else {
				string[] _n = FullName.Split(' ');
				if(_n.Count() > 0) {
					ret += _n[0].ToString() + " ";
					for(int i = 0; i < _n.Count() - 1; i++)
						ret += _n[i + 1].Substring(0, 1) + ". ";
					return ret;
				}
				else
					return FullName;
			}
		}

		private void Form1_Load(object sender, EventArgs e) {
			using(ContractTypesContext contractTypesContext = new ContractTypesContext()) {
				ContractTypes = contractTypesContext.ContractTypes.Where(x => x.ctp_IsActive == true).ToList<ContractTypes>();
				cmb_ContractType.DataSource = ContractTypes;
				cmb_ContractType.DisplayMember = "ctp_Name";
			}
			using(TypesTemplatesContext typesTemplatesContext = new TypesTemplatesContext()) {
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
