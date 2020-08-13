using System;
using System.Collections.Generic;
using System.Data.Linq;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnionContractWF.Helpers {
    class Database {
        private static DataContext _context;
        //private static string _connstr;

        static readonly INIManager manager = new INIManager(Environment.CurrentDirectory + @"/Settings.ini");
        //static string ConnectionString { get; set; } = "Data Source=" + manager.GetPrivateString("Connection", "Server") +
        //                                            "; Database=" + manager.GetPrivateString("Connection", "DatabaseName") +
        //                                            "; User Id=" + manager.GetPrivateString("Connection", "Login") +
        //                                            "; Password=" + manager.GetPrivateString("Connection", "Password") + ";";

        public static string connstr_helper {
            get => "Data Source=" + manager.GetPrivateString("Connection Helper", "Server") +
                                                    "; Database=" + manager.GetPrivateString("Connection Helper", "DatabaseName") +
                                                    "; User Id=" + manager.GetPrivateString("Connection Helper", "Login") +
                                                    "; Password=" + manager.GetPrivateString("Connection Helper", "Password") + ";";
            set { }
        }
        public static string connstr_crm {
            get => "Data Source=" + manager.GetPrivateString("Connection CRM", "Server") +
                                                    "; Database=" + manager.GetPrivateString("Connection CRM", "DatabaseName") +
                                                    "; User Id=" + manager.GetPrivateString("Connection CRM", "Login") +
                                                    "; Password=" + manager.GetPrivateString("Connection CRM", "Password") + ";";
            set { }
        }
        public static string connstr {
            get => "Data Source=" + manager.GetPrivateString("Connection CRM", "Server") +
                                                    "; Database=" + manager.GetPrivateString("Connection CRM", "DatabaseName") +
                                                    "; User Id=" + manager.GetPrivateString("Connection CRM", "Login") +
                                                    "; Password=" + manager.GetPrivateString("Connection CRM", "Password") + ";";
            set { }
        }

        /// <summary>
        /// Получчение Контекста базы данных
        /// </summary>
        /// <returns> Контекст _context</returns>
        public DataContext GetDataContext(bool isCRM) {
            if (isCRM) {
                if (!string.IsNullOrEmpty(connstr_crm)) {
                    return _context = new DataContext(connstr_crm);
                }
            }
            else {
                if (!string.IsNullOrEmpty(connstr_helper)) {
                    return _context = new DataContext(connstr_helper);
                }
            }
            //throw new ArgumentNullException(Resources.connstr_nullOrEmpty);            
            throw new ArgumentNullException("Строка подключения пустая");
        }
        public Database() {

        }
        /// <summary>
        /// Установка подключения к Базе Данных по строке подключения
        /// </summary>
        /// <param name="connectionString">Строка подключения</param>
        public Database(string connectionString) {
            connstr = connectionString;
        }
        /// <summary>
        /// Установка подключения к Базе Данных по параметрам
        /// </summary>
        /// <param name="ServerName">Название сервера или Ip-Address</param>
        /// <param name="DatabaseName">Название Базы Данных</param>
        /// <param name="DBLogin">Логин для подключения к Базе данных</param>
        /// <param name="DbPass">Пароль для подключения к базе данных</param>
        //public DB(string ServerName, string DatabaseName, string DBLogin, string DbPass) {
        //    manager.GetPrivateString("Connection", "Server");
        //    manager.GetPrivateString("Connection", "DatabaseName");
        //    manager.GetPrivateString("Connection", "Login");
        //    manager.GetPrivateString("Connection", "Password");
        //}
        /// <summary>
        /// Получение записей из базы данных по модели
        /// </summary>
        /// <typeparam name="T">Тип обьектоа (класс)</typeparam>
        /// <returns>List<T> обьектов из БД</returns>
        public List<T> GetDbTable<T>(bool isCRM) where T : class {
            using (_context = GetDataContext(isCRM)) {
                return _context.GetTable<T>().ToList();
            }
        }
        /// <summary>
        /// Вставка записи в базу данных по модели
        /// </summary>
        /// <typeparam name="T"> Тип обьекта (класс)</typeparam>
        /// <param name="ent">Обьект вставки <Т> </param>
        public void InsertDbRow<T>(T ent, bool isCRM) where T : class {
            using (_context = GetDataContext(isCRM)) {
                _context.GetTable<T>().InsertOnSubmit(ent);
                _context.SubmitChanges();
            }
        }
    }
}
