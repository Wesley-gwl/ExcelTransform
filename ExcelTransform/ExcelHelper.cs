using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTransform
{
    public class ExcelHelper
    {
        private static IWorkbook workbook = null;
        private static FileStream fs = null;

        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="fileName">excel文件路径</param>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public static DataTable ExcelToDataTable(string fileName, string sheetName, bool isFirstRowColumn)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook(fs);

                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
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
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; i++)
                        {
                            DataColumn column = new DataColumn(i.ToString());
                            data.Columns.Add(column);
                        }
                        startRow = sheet.FirstRowNum;
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    //当前有数据的行(合并主行)
                    IRow dataRowMain = null;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　
                        var isMarge = false;
                        if (row.GetCell(0) != null)
                        {
                            //赋值主行
                            dataRowMain = row;
                        }
                        else
                        {
                            //此行被合并，获取个别数据获取主行数据
                            isMarge = true;
                        }
                        DataRow dataRow = data.NewRow();
                        for (int j = 0; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                            {
                                dataRow[j] = row.GetCell(j).ToString();
                            }
                            else if (isMarge && dataRowMain.GetCell(j) != null)
                            {
                                dataRow[j] = dataRowMain.GetCell(j).ToString();
                            }
                        }
                        data.Rows.Add(dataRow);
                    }
                }

                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 将DataTable(DataSet)导出到Execl文档
        /// </summary>
        /// <param name="dataSet">传入一个DataSet</param>
        /// <param name="Outpath">导出路径（可以不加扩展名，不加默认为.xls）</param>
        /// <returns>返回一个Bool类型的值，表示是否导出成功</returns>
        /// True表示导出成功，Flase表示导出失败
        public static bool DataTableToExcel(DataTable dt, string Outpath)
        {
            bool result = false;
            try
            {
                int sheetIndex = 0;
                //根据输出路径的扩展名判断workbook的实例类型
                IWorkbook workbook = null;
                string pathExtensionName = Outpath.Trim().Substring(Outpath.Length - 5);
                if (pathExtensionName.Contains(".xlsx"))
                {
                    workbook = new XSSFWorkbook();
                }
                else if (pathExtensionName.Contains(".xls"))
                {
                    workbook = new HSSFWorkbook();
                }
                else
                {
                    Outpath = Directory.GetCurrentDirectory() + "/" + Outpath.Trim() + ".xls";
                    workbook = new HSSFWorkbook();
                }
                if (dt != null && dt.Rows.Count > 0)
                {
                    ISheet sheet = workbook.CreateSheet(string.IsNullOrEmpty(dt.TableName) ? ("sheet" + sheetIndex) : dt.TableName);//创建一个名称为Sheet0的表
                    int rowCount = dt.Rows.Count;//行数
                    int columnCount = dt.Columns.Count;//列数

                    ICellStyle normalStyle = workbook.CreateCellStyle();//普通格的样式,待完善
                    normalStyle.Alignment = HorizontalAlignment.Center;
                    normalStyle.VerticalAlignment = VerticalAlignment.Center;
                    normalStyle.WrapText = true;
                    //设置列头
                    IRow row = sheet.CreateRow(0);//excel第一行设为列头
                    for (int c = 0; c < columnCount; c++)
                    {
                        ICell cell = row.CreateCell(c);
                        cell.SetCellValue(dt.Columns[c].ColumnName);
                        sheet.AutoSizeColumn(c);
                    }

                    //设置每行每列的单元格,
                    for (int i = 0; i < rowCount; i++)
                    {
                        row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < columnCount; j++)
                        {
                            ICell cell = row.CreateCell(j);//excel第二行开始写入数据
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                            cell.CellStyle = normalStyle;
                        }
                    }
                }
                //向outPath输出数据
                using (FileStream fs = File.OpenWrite(Outpath))
                {
                    workbook.Write(fs);//向打开的这个xls文件中写入数据
                    result = true;
                }
                return result;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}