using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace GJOPENCASERECORD.ENTITY
{
    public class ExcelHelper
    {
        private string fileName = "";
        private IWorkbook workBook = null;
        private FileStream fs = null;
        public ExcelHelper(string fileName)
        {
            this.fileName = fileName;
        }
        /// <summary>  
        /// 将excel中的数据导入到DataTable中  
        /// </summary>  
        /// <param name="sheetName">excel工作薄sheet的名称</param>  
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>  
        /// <returns>返回的DataTable</returns> 
        public DataTable excelToDataTable(bool isFirstRowColumn, string sheetName = null)
        {
            ISheet sheet = null;
            DataTable dt_sheet = new DataTable();
            int startRow = 0;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                workBook = WorkbookFactory.Create(fs);
                /*
                if (fileName.IndexOf(".xlsx") > 0)
                {
                    workBook = new XSSFWorkbook();
                }
                else if (fileName.IndexOf(".xls") > 0)
                {
                    workBook = new HSSFWorkbook(fs);
                }
                 */
                if (sheetName != null)
                {
                    sheet = workBook.GetSheet(sheetName);
                    if (sheet == null)
                    {
                        sheet = workBook.GetSheetAt(0);
                    }
                }
                else
                {
                    sheet = workBook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数  

                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    dt_sheet.Columns.Add(column);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }

                    //最后一列的标号  
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　　　　　　  

                        DataRow dataRow = dt_sheet.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null  
                                dataRow[j] = row.GetCell(j).ToString();
                        }
                        dt_sheet.Rows.Add(dataRow);
                    }
                }
                fs.Close();
                return dt_sheet;
            }
            catch (Exception ex)
            {
                if (fs!=null)
                {
                    fs.Close();
                }
                AppInfo.WriteLogs(ex.Message);
                return null; 
            }
        }
    }
}
