using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using DF.BusEntry.DAL;
using RCD;

namespace GJOPENCASERECORD.ENTITY
{
    public static class AppInfo
    {
        /// <summary>
        /// 连接字符串
        /// </summary>
        public static string ConnectString = "";
        /// <summary>
        /// 数据库类型
        /// </summary>
        public static string DatabaseType = "";
        /// <summary>
        /// log文件位置
        /// </summary>
        public static string LogPosition = "";              
        /// <summary>
        /// 目标文件目录
        /// </summary>
        public static string PathName = "";
        /// <summary>
        /// 备份文件目录
        /// </summary>
        public static string BackUpRecords = "";

        //public static OracleInstance oracleDb = new OracleInstance();

        public static void Init()
        {
            try
            {
                GetAppSettings();
                //RCD.RCDB.ConnectString = ConnectString;
                //RCD.RCDB.InitValue("oledb", DatabaseType);
                //ORACLEHelper.connectionString = ConnectString;
                //oracleDb.ConnectString = ConnectString;
                //oracleDb.openConnection();
            }
            catch (Exception ex)
            {
                WriteLogs("初始化系统参数异常:" + ex.Message);
            }
        }

        private static void GetAppSettings()
        {
            DatabaseType = System.Configuration.ConfigurationManager.AppSettings.GetValues("DatabaseType")[0];
            ConnectString = System.Configuration.ConfigurationManager.AppSettings.GetValues("OleOracleConnect")[0];
            LogPosition = System.Configuration.ConfigurationManager.AppSettings.GetValues("LogPosition")[0];
            PathName = ConfigurationManager.AppSettings["pathName"];
            BackUpRecords = ConfigurationManager.AppSettings["BackUpRecords"];
        }

        /// <summary>
        /// 写入log文件
        /// </summary>
        /// <param name="logContent">log内容</param>
        public static void WriteLogs(string logContent)
        {
            try
            {
                if (!System.IO.Directory.Exists(LogPosition))
                {
                    System.IO.Directory.CreateDirectory(LogPosition);
                }
                System.IO.FileStream fs = new System.IO.FileStream(LogPosition + "//Sys" + DateTime.Now.ToString("yyyy-MM-dd") + ".log", System.IO.FileMode.Append);
                System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.GetEncoding("GB2312"));
                sw.Write(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + logContent + "\r\n");
                sw.Close();
                fs.Close();
            }
            catch
            { }
        }
    }
    
}
