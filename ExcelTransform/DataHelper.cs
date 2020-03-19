using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTransform
{
    public class DataHelper
    {
        /// <summary>
        /// databale分解数据表 当前第一行重复
        /// </summary>
        /// <param name="originalTab">需要分解的表</param>
        /// <param name="rowsNum">每个表包含的数据量</param>
        /// <param name="reRowNumber">重复行</param>
        /// <param name="reName">修改重复行单元格数据清空</param>
        /// <returns></returns>
        public static DataSet SplitDataTable(DataTable originalTab, int rowsNum, int? reRowNumber = null, string reName = null)
        {
            //获取所需创建的表数量
            int tableNum = originalTab.Rows.Count / rowsNum;

            //获取数据余数
            int remainder = originalTab.Rows.Count % rowsNum;
            if (remainder != 0) tableNum += 1;
            DataSet ds = new DataSet();

            //如果只需要创建1个表，直接将原始表存入DataSet
            if (tableNum == 0)
            {
                ds.Tables.Add(originalTab);
            }
            else
            {
                DataTable[] tableSlice = new DataTable[tableNum];

                //Save orginal columns into new table.
                for (int c = 0; c < tableNum; c++)
                {
                    tableSlice[c] = new DataTable();
                    foreach (DataColumn dc in originalTab.Columns)
                    {
                        tableSlice[c].Columns.Add(dc.ColumnName, dc.DataType);
                    }
                }
                //Import Rows
                for (int i = 0; i < tableNum; i++)
                {
                    // if the current table is not the last one
                    if (i != tableNum - 1)
                    {
                        for (int j = i * rowsNum; j < ((i + 1) * rowsNum); j++)
                        {
                            if (reRowNumber != null && j % rowsNum == 0)
                            {
                                tableSlice[i].ImportRow(originalTab.Rows[j]);//多加一行
                                if (reName != null)
                                {
                                    tableSlice[i].Rows[0][reName] = "";//改变SKU防止
                                }
                            }
                            tableSlice[i].ImportRow(originalTab.Rows[j]);
                        }
                    }
                    else
                    {
                        for (int k = i * rowsNum; k < (i * rowsNum + remainder); k++)
                        {
                            tableSlice[i].ImportRow(originalTab.Rows[k]);
                        }
                    }
                }

                //add all tables into a dataset
                foreach (DataTable dt in tableSlice)
                {
                    ds.Tables.Add(dt);
                }
            }
            return ds;
        }

        /// <summary>
        /// datagridview 转 datatable
        /// </summary>
        /// <param name="dgv"></param>
        /// <returns></returns>
        public static DataTable GetDgvToTable(DataGridView dgv)
        {
            var dt = new DataTable();

            // 列强制转换
            for (int count = 0; count < dgv.Columns.Count; count++)
            {
                DataColumn dc = new DataColumn(dgv.Columns[count].Name.ToString());
                dt.Columns.Add(dc);
            }

            // 循环行
            for (int count = 0; count < dgv.Rows.Count; count++)
            {
                DataRow dr = dt.NewRow();
                for (int countsub = 0; countsub < dgv.Columns.Count; countsub++)
                {
                    dr[countsub] = Convert.ToString(dgv.Rows[count].Cells[countsub].Value);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}