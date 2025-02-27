using System;
using System.ServiceProcess;
using System.Timers;
using System.Data.SqlClient;
using System.Management;
using System.IO;

namespace EnvanterServisi
{
    public partial class EnvanterServisi : ServiceBase
    {

        string seriNo;
        string computerName;
        string ramGB;
        string diskGB;
        string macAddress;
        string userName;
        string islemci;
        string model;
        string asset;
        string serviceName = "MSSQLSERVER";
        Timer timer = new Timer();
        Timer timer2 = new Timer();


        public EnvanterServisi()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {

            Logger("Servis çalışmaya başladı." + DateTime.Now);
            if (!IsSqlServiceRunning(serviceName))
            {
                Logger("SQL servisi kapalı, başlatılıyor.");
                StartSqlService(serviceName);
            }
            else
            {
                Logger("SQL servisi zaten çalışıyor.");

            }
            EnvanterBilgileriniGonder();
            timer.Interval = 1000 * 60 * 5;
            timer.Elapsed += new ElapsedEventHandler(TimerElapsed);
            timer.Enabled = true;
            timer.Start();
            timer2.Interval = 1000 * 60 * 60;
            timer2.Elapsed += new ElapsedEventHandler(Timer2Elapsed);
            timer2.Enabled = true;
            timer2.Start();


        }

        protected override void OnStop()
        {
            Logger("Servis durdu." + DateTime.Now);
            timer.Stop();
        }


        private void Timer2Elapsed(object source, ElapsedEventArgs e)
        {
            File.Delete(AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\Log.txt");
        }

        private void TimerElapsed(object source, ElapsedEventArgs e)
        {
            Logger("Servis Çalışıyor." + DateTime.Now);
            if (DateTime.Now.Hour == 15 && DateTime.Now.Minute == 0)
            {
                EnvanterBilgileriniGonder();
            }
        }
        private void EnvanterBilgileriniGonder()
        {

            Logger("Servis çalışıyor." + DateTime.Now);

            EnvanterBilgileriniAl();

            string connectionString = @"Data Source=localhost;Initial Catalog=MyLocalDB;Integrated Security=True;Connect Timeout=30;Encrypt=False";
            DateTime anlikTarih = DateTime.Now;
            string anlikTarihTXT = anlikTarih.ToString("yyyy-MM-dd HH:mm:ss");



            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    Logger("Bağlantı başarılı." + DateTime.Now);
                    string assetFinder = "SELECT Asset FROM EnvanterTablosu WHERE [Seri No] LIKE @seriNo";
                    using (SqlCommand assetFinderKomut = new SqlCommand(assetFinder, conn))
                    {
                        assetFinderKomut.Parameters.AddWithValue("@seriNo", seriNo);
                        asset = assetFinderKomut.ExecuteScalar().ToString();
                    }


                    string finder = "SELECT COUNT(*) FROM EnvanterTablosu WHERE Asset LIKE @asset AND [Seri No] LIKE @seriNo";
                    using (SqlCommand finderKomut = new SqlCommand(finder, conn))
                    {
                        finderKomut.Parameters.AddWithValue("@asset", asset);
                        finderKomut.Parameters.AddWithValue("@seriNo", seriNo);

                        int count = (int)finderKomut.ExecuteScalar();

                        string creator = $"IF NOT EXISTS(SELECT * FROM sys.tables WHERE name = '{asset}')" +
                                                "BEGIN" +
                                                $"   CREATE TABLE \"{asset}\" " +
                                                "   (Asset TEXT," +
                                                "   [Seri No] TEXT," +
                                                "   [Bilgisayar Modeli] TEXT," +
                                                "   [Bilgisayar Adı] TEXT," +
                                                "   RAM TEXT," +
                                                "   [Disk Boyutu] TEXT," +
                                                "   [MAC Adresi] TEXT," +
                                                "   [İşlemci Modeli] TEXT," +
                                                "   [Kullanıcı] TEXT," +
                                                "   [Değişiklik Tarihi] TEXT);" +
                                                "END;";
                        using (SqlCommand creatorKomut = new SqlCommand(creator, conn))
                        {

                            creatorKomut.ExecuteNonQuery();
                        }
                        if (count <= 0)
                        {



                            string inserter = "INSERT INTO EnvanterTablosu " +
                                "(Asset, [Seri No], [Bilgisayar Modeli],[Bilgisayar Adı],RAM,[Disk Boyutu],[MAC Adresi],[İşlemci Modeli],[Kullanıcı],[Değişiklik Tarihi]) " +
                                "VALUES (@asset, @seriNo, @bilgisayarModeli, @bilgisayarAdi, @ram, @diskBoyutu, @macAdresi, @islemciModeli, @kullanici, @anlikTarih)";
                            using (SqlCommand inserterKomut = new SqlCommand(inserter, conn))
                            {
                                inserterKomut.Parameters.AddWithValue("@asset", asset);
                                inserterKomut.Parameters.AddWithValue("@seriNo", seriNo);
                                inserterKomut.Parameters.AddWithValue("@bilgisayarModeli", model);
                                inserterKomut.Parameters.AddWithValue("@bilgisayarAdi", computerName);
                                inserterKomut.Parameters.AddWithValue("@ram", ramGB);
                                inserterKomut.Parameters.AddWithValue("@diskBoyutu", diskGB);
                                inserterKomut.Parameters.AddWithValue("@macAdresi", macAddress);
                                inserterKomut.Parameters.AddWithValue("@islemciModeli", islemci);
                                inserterKomut.Parameters.AddWithValue("@kullanici", userName);
                                inserterKomut.Parameters.AddWithValue("@anlikTarih", anlikTarihTXT);
                                inserterKomut.ExecuteNonQuery();
                            }

                            string inserter2 = $"INSERT INTO \"{asset}\" " +
                                "(Asset, [Seri No], [Bilgisayar Modeli],[Bilgisayar Adı],RAM,[Disk Boyutu],[MAC Adresi],[İşlemci Modeli],[Kullanıcı],[Değişiklik Tarihi]) " +
                                "VALUES (@asset, @seriNo, @bilgisayarModeli, @bilgisayarAdi, @ram, @diskBoyutu, @macAdresi, @islemciModeli, @kullanici, @anlikTarih)";
                            using (SqlCommand inserterKomut = new SqlCommand(inserter2, conn))
                            {
                                inserterKomut.Parameters.AddWithValue("@asset", asset);
                                inserterKomut.Parameters.AddWithValue("@seriNo", seriNo);
                                inserterKomut.Parameters.AddWithValue("@bilgisayarModeli", model);
                                inserterKomut.Parameters.AddWithValue("@bilgisayarAdi", computerName);
                                inserterKomut.Parameters.AddWithValue("@ram", ramGB);
                                inserterKomut.Parameters.AddWithValue("@diskBoyutu", diskGB);
                                inserterKomut.Parameters.AddWithValue("@macAdresi", macAddress);
                                inserterKomut.Parameters.AddWithValue("@islemciModeli", islemci);
                                inserterKomut.Parameters.AddWithValue("@kullanici", userName);
                                inserterKomut.Parameters.AddWithValue("@anlikTarih", anlikTarihTXT);
                                inserterKomut.ExecuteNonQuery();
                            }
                        }
                        else
                        {
                            string inserter = "UPDATE EnvanterTablosu " +
                                "SET" +
                                "[Bilgisayar Modeli] = @bilgisayarModeli," +
                                "[Bilgisayar Adı] = @bilgisayarAdi," +
                                "RAM = @ram," +
                                "[Disk Boyutu] = @diskBoyutu," +
                                "[İşlemci Modeli] = @islemciModeli," +
                                "[Kullanıcı] = @kullanici," +
                                "[Değişiklik Tarihi] = @anlikTarih " +
                                "WHERE [Seri No] LIKE @seriNo AND [MAC Adresi] LIKE @macAdresi";
                            using (SqlCommand cmd2 = new SqlCommand(inserter, conn))
                            {
                                cmd2.Parameters.AddWithValue("@seriNo", seriNo);
                                cmd2.Parameters.AddWithValue("@bilgisayarModeli", model);
                                cmd2.Parameters.AddWithValue("@bilgisayarAdi", computerName);
                                cmd2.Parameters.AddWithValue("@ram", ramGB);
                                cmd2.Parameters.AddWithValue("@diskBoyutu", diskGB);
                                cmd2.Parameters.AddWithValue("@macAdresi", macAddress);
                                cmd2.Parameters.AddWithValue("@islemciModeli", islemci);
                                cmd2.Parameters.AddWithValue("@kullanici", userName);
                                cmd2.Parameters.AddWithValue("@anlikTarih", anlikTarihTXT);
                                cmd2.ExecuteNonQuery();

                            }

                            string inserter2 = $"INSERT INTO \"{asset}\" " +
                                "(Asset, [Seri No], [Bilgisayar Modeli],[Bilgisayar Adı],RAM,[Disk Boyutu],[MAC Adresi],[İşlemci Modeli],[Kullanıcı],[Değişiklik Tarihi]) " +
                                "VALUES (@asset, @seriNo, @bilgisayarModeli, @bilgisayarAdi, @ram, @diskBoyutu, @macAdresi, @islemciModeli, @kullanici, @anlikTarih)";
                            using (SqlCommand inserterKomut = new SqlCommand(inserter2, conn))
                            {
                                inserterKomut.Parameters.AddWithValue("@asset", "000.000.00000");
                                inserterKomut.Parameters.AddWithValue("@seriNo", seriNo);
                                inserterKomut.Parameters.AddWithValue("@bilgisayarModeli", model);
                                inserterKomut.Parameters.AddWithValue("@bilgisayarAdi", computerName);
                                inserterKomut.Parameters.AddWithValue("@ram", ramGB);
                                inserterKomut.Parameters.AddWithValue("@diskBoyutu", diskGB);
                                inserterKomut.Parameters.AddWithValue("@macAdresi", macAddress);
                                inserterKomut.Parameters.AddWithValue("@islemciModeli", islemci);
                                inserterKomut.Parameters.AddWithValue("@kullanici", userName);
                                inserterKomut.Parameters.AddWithValue("@anlikTarih", anlikTarihTXT);
                                inserterKomut.ExecuteNonQuery();
                            }
                        }

                    }
                    Logger("Veriler DB'ye eklendi." + DateTime.Now);
                }
                catch (Exception ex)
                {
                    Logger("Bağlantı başarısız." + ex.Message + DateTime.Now);
                }
            }
        }
        public void EnvanterBilgileriniAl()
        {
            ulong ramCapacity;
            ulong diskCapacity = 0;


            ManagementObjectSearcher biosSearcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BIOS");
            foreach (ManagementObject obj in biosSearcher.Get())
            {
                seriNo = obj["SerialNumber"].ToString();


            }

            ManagementObjectSearcher compSearcher = new ManagementObjectSearcher("SELECT Name, Model, TotalPhysicalMemory, UserName FROM Win32_ComputerSystem");
            foreach (ManagementObject obj in compSearcher.Get())
            {
                computerName = obj["Name"].ToString();
                userName = obj["UserName"].ToString();
                userName = userName.Split('\\')[1];
                model = obj["Model"].ToString();
                ramCapacity = (ulong)obj["TotalPhysicalMemory"];
                ramGB = Math.Ceiling((decimal)(ramCapacity) / 1073741824).ToString("F2") + " GB";

            }

            ManagementObjectSearcher diskSearcher = new ManagementObjectSearcher("SELECT * FROM Win32_LogicalDisk WHERE DriveType = 3");
            foreach (ManagementObject obj in diskSearcher.Get())
            {
                if (obj["Size"] != null)
                {
                    ulong diskSize = (ulong)obj["Size"];
                    decimal diskDecimal = (decimal)diskSize;
                    diskCapacity += (ulong)Math.Ceiling(diskDecimal);

                }
                diskGB = Math.Ceiling((decimal)(diskCapacity) / 1073741824).ToString("F2") + " GB";
            }

            ManagementObjectSearcher macSearcher = new ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionStatus = 2");
            foreach (ManagementObject obj in macSearcher.Get())
            {
                macAddress = obj["MacAddress"].ToString();

            }


            ManagementObjectSearcher procSearcher = new ManagementObjectSearcher("SELECT Name FROM Win32_Processor");
            foreach (ManagementObject obj in procSearcher.Get())
            {
                islemci = obj["Name"].ToString();

            }




        }

        private void Logger(string mesaj)
        {
            string dosyaYolu = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(dosyaYolu))
            {
                Directory.CreateDirectory(dosyaYolu);
            }

            string textYolu = dosyaYolu + "\\Log.txt";

            if (!File.Exists(textYolu))
            {
                using (StreamWriter sw = File.CreateText(textYolu))
                {
                    sw.WriteLine(mesaj);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(textYolu))
                {
                    sw.WriteLine(mesaj);
                }
            }
        }

        private bool IsSqlServiceRunning(string serviceName)
        {
            try
            {
                ServiceController sc = new ServiceController(serviceName);
                return sc.Status == ServiceControllerStatus.Running;

            }
            catch (Exception ex)
            {
                Logger("SQL servisi çalışmıyor." + ex.Message);
                return false;
            }
        }

        private void StartSqlService(string serviceName)
        {
            try
            {
                ServiceController sc = new ServiceController(serviceName);
                if (sc.Status == ServiceControllerStatus.Stopped || sc.Status == ServiceControllerStatus.Paused)
                {
                    sc.Start();
                    sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
                    Logger("SQL servisi başlatıldı.");
                }
            }
            catch (Exception ex)
            {
                Logger("Sql servisi başlatılamadı." + ex.Message);
            }
        }


    }
}
