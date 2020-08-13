using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace UnionContractWF.Helpers {
    class INIManager {
        public INIManager(string aPath) {
            Path = aPath;
        }
        public INIManager() : this("") { }

        public string GetPrivateString(string aSection, string aKey) {
            StringBuilder buffer = new StringBuilder(SIZE);

            GetPrivateString(aSection, aKey, null, buffer, SIZE, Path);

            return buffer.ToString();
        }

        public void WritePrivateString(string aSection, string aKey, string aValue) {
            WritePrivateString(aSection, aKey, aValue, Path);
        }

        public string Path { get; set; } = null;

        private const int SIZE = 1024;

        [DllImport("kernel32.dll", EntryPoint = "GetPrivateProfileString", CharSet = CharSet.Unicode)]
        private static extern int GetPrivateString(string section, string key, string def, StringBuilder buffer, int size, string path);

        [DllImport("kernel32.dll", EntryPoint = "WritePrivateProfileString", CharSet = CharSet.Unicode)]
        private static extern int WritePrivateString(string section, string key, string str, string path);
    }
}
