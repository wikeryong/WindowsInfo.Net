using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;

namespace WindowsInfo.Net
{

    /// <summary>
    /// windows api 名称
    /// </summary>
    public enum WindowsAPIType
    {
        /// <summary>
        /// 内存
        /// </summary>
        Win32_PhysicalMemory,
        /// <summary>
        /// cpu
        /// </summary>
        Win32_Processor,
        /// <summary>
        /// 硬盘
        /// </summary>
        win32_DiskDrive,
        /// <summary>
        /// 电脑型号
        /// </summary>
        Win32_ComputerSystemProduct,
        /// <summary>
        /// 分辨率
        /// </summary>
        Win32_DesktopMonitor,
        /// <summary>
        /// 显卡
        /// </summary>
        Win32_VideoController,
        /// <summary>
        /// 操作系统
        /// </summary>
        Win32_OperatingSystem

    }
    public enum WindowsAPIKeys
    {
        /// <summary>
        /// 名称
        /// </summary>
        Name,
        /// <summary>
        /// 显卡芯片
        /// </summary>
        VideoProcessor,
        /// <summary>
        /// 显存大小
        /// </summary>
        AdapterRAM,
        /// <summary>
        /// 分辨率宽
        /// </summary>
        ScreenWidth,
        /// <summary>
        /// 分辨率高
        /// </summary>
        ScreenHeight,
        /// <summary>
        /// 电脑型号
        /// </summary>
        Version,
        /// <summary>
        /// 硬盘容量
        /// </summary>
        Size,
        /// <summary>
        /// 内存容量
        /// </summary>
        Capacity,
        /// <summary>
        /// cpu核心数
        /// </summary>
        NumberOfCores
    }
    class Program
    {

        public static string ToGB(double size, double mod)
        {
            String[] units = new String[] { "B", "KB", "MB", "GB", "TB", "PB" };
            int i = 0;
            while (size >= mod)
            {
                size /= mod;
                i++;
            }
            return Math.Round(size) + units[i];
        }

        /// <summary>
        /// 获取硬盘容量
        /// </summary>
        public static string GetDiskSize()
        {
            string result = string.Empty;
            StringBuilder sb = new StringBuilder();
            try
            {
                string hdId = string.Empty;
                ManagementClass hardDisk = new ManagementClass(WindowsAPIType.win32_DiskDrive.ToString());
                ManagementObjectCollection hardDiskC = hardDisk.GetInstances();
                foreach (ManagementObject m in hardDiskC)
                {
                    long capacity = Convert.ToInt64(m[WindowsAPIKeys.Size.ToString()].ToString());
                    sb.Append(ToGB(capacity, 1000.0) + "+");
                }
                result = sb.ToString().TrimEnd('+');
            }
            catch
            {

            }
            return result;
        }
        static void Main(string[] args)
        {
            try
            {
                string filename = "info.txt";

                Console.Write("请输入部门：");
                string dept = Console.ReadLine();


                Console.Write("请输入姓名：");
                string name = Console.ReadLine();
                FileStream fs = new FileStream(dept + "_" + name + "_"+filename, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.Default);

                Console.WriteLine("正在获取信息，请稍后...");

                // 系统信息
                sw.WriteLine("---------  系统信息  ---------");
                SystemInfo systemInfo = new SystemInfo();
                sw.WriteLine("计算机名：" + systemInfo.GetComputerName());
                sw.WriteLine("登录用户名：" + systemInfo.GetUserName());
                sw.WriteLine("操作系统类型：" + systemInfo.GetSystemType());
                sw.WriteLine("系统安装日期：" + GetInstallDate());
                sw.WriteLine("使用部门：" + dept);
                sw.WriteLine("使用人：" + name);
                sw.WriteLine("\n");

                // 硬件信息
                sw.WriteLine("---------  硬件信息  ---------");
                HardwareInfo hardwareInfo = new HardwareInfo();
                sw.WriteLine("本机的MAC地址：" + hardwareInfo.GetLocalMac());

                Console.WriteLine("主板序列号：{0,20:s}", hardwareInfo.GetBIOSSerialNumber());
                sw.WriteLine("主板序列号：" + hardwareInfo.GetBIOSSerialNumber());
                sw.WriteLine("主板制造厂商：" + hardwareInfo.GetBoardManufacturer());
                sw.WriteLine("主板编号：" + hardwareInfo.GetBoardID());
                sw.WriteLine("主板编号：" + hardwareInfo.GetBoardID());
                sw.WriteLine("主板型号：" + hardwareInfo.GetBoardType());

                Console.WriteLine("CPU序列号：{0,20:s}", hardwareInfo.GetCPUSerialNumber());
                sw.WriteLine("CPU序列号：" + hardwareInfo.GetCPUSerialNumber());
                sw.WriteLine("CPU编号：" + hardwareInfo.GetCPUID());
                sw.WriteLine("CPU版本信息：" + hardwareInfo.GetCPUVersion());
                sw.WriteLine("CPU名称信息：" + hardwareInfo.GetCPUName());
                sw.WriteLine("CPU制造厂商：" + hardwareInfo.GetCPUManufacturer());

                //sw.WriteLine("物理硬盘序列号：" + hardwareInfo.GetHardDiskSerialNumber());


                sw.WriteLine("硬盘容量：" + GetDiskSize());

                List<string> hardDiskSerialList = hardwareInfo.GetHardDiskSerialNumber();
                for (int i = 1; i <= hardDiskSerialList.Count; i++)
                {
                    sw.WriteLine("物理硬盘 " + i + " 序列号：" + hardDiskSerialList[i-1]);
                }

                List<string> diskSerialList = hardwareInfo.GetDiskSerialNumber();
                for(int i = 1; i <= diskSerialList.Count; i++)
                { 
                    sw.WriteLine("磁盘 "+i+" 序列号：" + diskSerialList[i-1]);
                }
                

                sw.WriteLine();
                List<ManagementObject> netCardAddrList = hardwareInfo.GetNetCardMACAddress();
                foreach(var s in netCardAddrList)
                {
                    sw.WriteLine("--网卡地址：" + s["Description"]);
                    sw.WriteLine("  MAC：" + s["MACAddress"] + "描述：");
                }
                sw.WriteLine();
                List<ManagementObject> macs = hardwareInfo.GetAllMacAddress();
                foreach(var s in macs)
                {
                    sw.WriteLine("--网卡硬件地址：" + s["Description"]);
                    sw.WriteLine("  MAC："+s["MACAddress"]+"描述：");
                }

                string totalMemory = hardwareInfo.GetPhysicalMemory();
                sw.WriteLine("物理内存：" + totalMemory+"，大小（GB）："+ ToGB(float.Parse(totalMemory), 1024));

                sw.WriteLine("显卡PNPDeviceID：" + hardwareInfo.GetVideoPNPID());
                sw.WriteLine("声卡PNPDeviceID：" + hardwareInfo.GetSoundPNPID());
                sw.WriteLine("\n");

                // 网络信息
                sw.WriteLine("---------  网络信息  ---------");
                NetworkInfo networkInfo = new NetworkInfo();
                sw.WriteLine("IP地址：" + networkInfo.GetIPAddress());
                string[] LocalIpAddress = networkInfo.GetLocalIpAddress();
                foreach (var item in LocalIpAddress)
                {
                    sw.WriteLine("本地ip地址：" + item);
                }
                string[] ExtenalIpAddress = networkInfo.GetExtenalIpAddress();
                foreach (var item in ExtenalIpAddress)
                {
                    sw.WriteLine("外网ip地址：" + item);
                }
                sw.WriteLine("\n");

                // 软件信息
                sw.WriteLine("---------  软件信息  ---------");
                Dictionary<string, string> softDictionary = SoftwaresInfo.GetSoftWares();
                // 把软件名、版本号写入文件
                foreach (var item in softDictionary)
                {
                    //Console.WriteLine(item.Key + "\t" + item.Value);
                    sw.WriteLine(item.Key + "\t" + item.Value);
                }
                sw.WriteLine("\n");

                sw.Flush();
                sw.Close();
                fs.Close();

                Console.WriteLine("信息文件保存在：" + Environment.CurrentDirectory + @"\" +dept+"_"+name+ "_"+filename);
                Console.ReadKey();
            }
            catch (Exception e1)
            {
                Console.WriteLine("读取失败");
                Console.WriteLine(e1.Message);
                Console.WriteLine(e1.StackTrace.ToString());
                
                //System.Diagnostics.Debug.WriteLine("Program Main," + e1.Message.ToString());
            }
        }

        static string GetInstallDate()
        {
            System.Management.ObjectQuery MyQuery = new System.Management.ObjectQuery("SELECT * FROM Win32_OperatingSystem");
            //System.Management.ObjectQuery MyQuery = new System.Management.ObjectQuery("SELECT * FROM Win32_ComputerSystem");
            System.Management.ManagementScope MyScope = new System.Management.ManagementScope();
            ManagementObjectSearcher MySearch = new ManagementObjectSearcher(MyScope, MyQuery);
            ManagementObjectCollection MyCollection = MySearch.Get();
            string StrInfo = "";
            foreach (ManagementObject MyObject in MyCollection)
            {
                //显示系统基本信息
                StrInfo = MyObject.GetText(TextFormat.Mof);
                string[] MyString = { "" };
                //重新启动计算机
                //MyObject.InvokeMethod("Reboot",MyString);								
                //关闭计算机
                //MyObject.InvokeMethod("Shutdown",MyString);				
            }
            string InstallDate = StrInfo.Substring(StrInfo.LastIndexOf("InstallDate") + 15, 14);
            return InstallDate;

        }
    }
}
