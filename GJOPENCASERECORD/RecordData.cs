using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DF.BusEntry.DAL;
using GJOPENCASERECORD.ENTITY;

namespace GJOPENCASERECORD
{
    public class RecordData
    {


        /// <summary>
        /// 查询出目录中所有的excel并插入数据库
        /// 1. 判断目录是否存在，不存在则创建目录
        /// 2. 查询出目录中所有的excel
        /// 2.1 获取excel中数据datatable
        /// 2.2 将datatable转化为list<entity> 格式
        /// 3. 分别插入数据库
        /// 4. 将文件转移到另一个文件夹备份
        /// </summary>
        public void dealDataOnTime()
        {
            AppInfo.Init();
            dirExistOrCreate(AppInfo.PathName);
            List<string> fileList = getFileName(AppInfo.PathName);
            if (fileList.Count > 0)
            {
                Task[] tasks = new Task[fileList.Count];
                try
                {
                    for (int i = 0; i < fileList.Count; i++)
                    {
                        dealDataAsync(fileList[i], ref tasks[i]);
                    }
                    Task.WaitAll(tasks);
                }
                catch (Exception err)
                {
                    AppInfo.WriteLogs(err.Message);
                }
            }
        }

        /// <summary>
        /// 存在多个excel文件时，异步处理数据
        /// </summary>
        /// <param name="filePath"></param>
        public void dealDataAsync(string filePath, ref Task task)
        {
            DataTable dt = null;
            List<OpenRecord> entityList = null;
            task = Task.Factory.StartNew(() =>
            {
                bool ERRFLAG = false;
                ExcelHelper helper = new ExcelHelper(filePath);
                dt = helper.excelToDataTable(true);
                if (dt != null)
                {
                    entityList = datatableToList(dt,ref ERRFLAG);
                    deleteSameData(entityList, ref ERRFLAG);
                    insertDataToDataBase(entityList, ref ERRFLAG);
                    moveFileDir(filePath, ref ERRFLAG);
                }
            });
        }
        /// <summary>
        /// 根据时间 判断数据库中是否存在当天数据，
        /// 如果存在，删除
        /// </summary>
        /// <param name="entityList"></param>
        public void deleteSameData(List<OpenRecord> entityList,ref bool ERRFLAG)
        {
            if (entityList.Count>0)
            {
                try
                {
                    string recordDate = DateTime.Parse(entityList.First().recordTime, new DateTimeFormatInfo()
                    {
                        FullDateTimePattern = "yyyy-MM-dd HH:mi:ss"
                    }).ToString("yyyy-MM-dd");
                    string sql_delete = "delete gj_opencaserecords t where to_char(t.记录时间,'yyyy-MM-dd') = '"+recordDate+"'";
                    int count = ORACLEHelper.ExecuteSql(sql_delete);
                }
                catch (Exception err)
                {
                    ERRFLAG = true;
                    AppInfo.WriteLogs(err.Message);
                }
                
            }
        }
        /// <summary>
        /// 将文件转移到某个目录
        /// 文件名重新命名为 原文件名+MM+DD+HH+mm
        /// </summary>
        /// <param name="fullPath"></param>
        public void moveFileDir(string fullPath, ref bool ERRFLAG)
        {
            if (!ERRFLAG)
            {
                DateTime dtime = DateTime.Now;
                string[] dirs = fullPath.Split('\\');
                //文件名
                string fileName = dirs[dirs.Count() - 1];
                int lastIndex = fileName.LastIndexOf('.');
                //文件名
                string name = fileName.Substring(0, lastIndex);
                //后缀
                string exName = fileName.Substring(lastIndex, fileName.Count() - lastIndex);
                fileName = name + dtime.ToString("MMddHHmm") + exName;
                string destPath = AppInfo.BackUpRecords + "\\" + fileName;
                dirExistOrCreate(AppInfo.BackUpRecords);
                Directory.Move(fullPath, destPath);
            }
        }
        /// <summary>
        /// 判断目录（文件夹）是否存在，不存在则创建
        /// </summary>
        /// <param name="path"></param>
        private void dirExistOrCreate(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
        /// <summary>
        /// 返回目录下所有excel文件完全路径list
        /// null 或 list
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public List<string> getFileName(string path)
        {
            List<string> fileList = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories).ToList();
            if (fileList.Count > 0)
            {
                for (int i = 0; i < fileList.Count; i++)
                {
                    if (!isExcelFile(fileList[i]))
                    {
                        fileList.RemoveAt(i);
                        i--;
                    }
                }
            }
            return fileList;
        }

        /// <summary>
        /// 是否是EXCEL文件
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private bool isExcelFile(string filePath)
        {
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            byte[] b = new byte[4];
            string temstr = "";
            //将文件流读取的文件写入到字节数组
            if (fs.Length > 0)
            {
                fs.Read(b, 0, 4);

                fs.Close();

                for (int i = 0; i < b.Length; i++)
                {
                    temstr += Convert.ToString(b[i], 16);
                }
            }
            if (temstr.ToUpper() == "D0CF11E0" || temstr.ToUpper() == "504B34")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 根据列名称将datatable转化list(模版列名不能变化）
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private List<OpenRecord> datatableToList(DataTable dt,ref bool ERRFLAG)
        {
            List<OpenRecord> entityList = new List<OpenRecord>();
            if (dt != null)
            {
                removeEmptyDataRow(ref dt);
                try
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        OpenRecord entity = new OpenRecord();
                        entity.lineName = row["线路名称"].ToString() + "路";
                        entity.carNum = row["车辆编号"].ToString();
                        entity.owiner = row["钥匙卡/持卡人"].ToString();
                        entity.ouCardNum = row["换出内胆编号"].ToString();
                        entity.inCardNum = row["换入内胆编号"].ToString();
                        entity.recordTime = DateTime.Parse(row["记录时间"].ToString(), new DateTimeFormatInfo()
                        {
                            FullDateTimePattern = "MM/dd/yyyy HH:mm"
                        }).ToString("yyyy-MM-dd HH:mm");
                        entityList.Add(entity);
                    }
                }
                catch (Exception err)
                {
                    ERRFLAG = true;
                    AppInfo.WriteLogs(err.Message);
                }
            }
            return entityList;
        }
        /// <summary>
        /// 删除空白行
        /// </summary>
        /// <param name="dt"></param>
        private void removeEmptyDataRow(ref DataTable dt)
        {
            if (dt != null)
            {

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    bool emptyFlag = true;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (!string.IsNullOrEmpty(dt.Rows[i][j].ToString().Trim()))
                        {
                            emptyFlag = false;
                        }
                    }
                    if (emptyFlag)
                    {
                        dt.Rows.RemoveAt(i);
                        i--;
                    }
                }
            }

        }
        /// <summary>
        /// 向数据库插入数据
        /// </summary>
        /// <param name="entityList"></param>
        private void insertDataToDataBase(List<OpenRecord> entityList, ref bool ERRFLAG)
        {
            if (ERRFLAG)
            {
                return;
            }
            int minCount = 0;
            try
            {
                while (minCount < entityList.Count)
                {
                    int maxCount = entityList.Count > 30 + minCount ? 30 + minCount : entityList.Count;
                    StringBuilder sql = new StringBuilder("insert into gj_opencaserecords (线路名称,车牌号,持卡人,换出内胆编号,换入内胆编号,记录时间) select * from ( ");
                    for (int i = minCount; i < maxCount; i++)
                    {
                        OpenRecord entity = entityList[i];
                        sql.Append("select '").Append(entity.lineName + "' 线路名称,'").Append(entity.carNum + "' 车牌号,'").Append(entity.owiner + "' 持卡人,'").Append(entity.ouCardNum + "' 换出内胆编号,'").Append(entity.inCardNum + " ' 换入内胆编号,to_date('").Append(entity.recordTime).Append("','yyyy-MM-dd HH24:mi:ss') 记录时间 from dual ");
                        if (i < maxCount - 1)
                        {
                            sql.Append(" union all ");
                        }
                    }
                    sql.Append(" )");
                    //AppInfo.WriteLogs(sql.ToString());
                    //ORACLEHelper helper = new ORACLEHelper();
                    ORACLEHelper.ExecuteSql(sql.ToString());
                    //bool result = RCD.RCDB.Execute(sql.ToString());
                    minCount += 30;
                }
            }
            catch (Exception err)
            {
                ERRFLAG = true;
                AppInfo.WriteLogs(err.Message);
            }


        }
    }
}
