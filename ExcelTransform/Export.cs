using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTransform
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 打开读取文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnOpen_Click(object sender, EventArgs e)
        {
            try
            {
                OFDExcel.ShowDialog();
                var path = OFDExcel.FileNames;  //获取openFileDialog控件选择的文件名数组   string strpath = "";
                tbtUrl.Text = path[0];
                var dt = ExcelHelper.ExcelToDataTable(path[0], null, true);
                LblTotalLine.Text = "当前总共" + dt.Rows.Count.ToString() + "行";
                LblSuccessLine.Text = "当前总共" + 0.ToString() + "行";
                LblFailLine.Text = "当前总共" + 0.ToString() + "行";
                BindDate(DGVExport, dt);
            }
            catch
            {
                tbtUrl.Text = "请选择文件";
            }
        }

        /// <summary>
        /// 绑定展示dgv
        /// </summary>
        /// <param name="dt"></param>
        private void BindDate(DataGridView dgv, DataTable dt)
        {
            try
            {
                dgv.Columns.Clear();
                //添加新列
                foreach (DataColumn col in dt.Columns)
                {
                    dgv.Columns.Add(col.ColumnName, col.ColumnName);
                }
                dgv.Rows.Clear();
                dgv.Rows.Add(dt.Rows.Count);//增加同等数量的行数
                int i = 0;
                foreach (DataRow row in dt.Rows)//逐个读取单元格的内容；
                {
                    DataGridViewRow r1 = dgv.Rows[i];
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        r1.Cells[j].Value = row[j];
                    }
                    i++;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 转换
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnTransfrom_Click(object sender, EventArgs e)
        {
            var filepath = tbtUrl.Text;
            if (string.IsNullOrEmpty(filepath) || filepath == "请选择文件")
            {
                MessageBox.Show("转换失败，请选择文件");
            }
            else if (string.IsNullOrEmpty(TxtBuyUrl.Text))
            {
                MessageBox.Show("请填写商品链接");
            }
            else
            {
                var dt = ExcelHelper.ExcelToDataTable(filepath, null, true);
                Transfrom(dt);
                MessageBox.Show("转换成功，请查看成功数据和失败数据");
            }
        }

        /// <summary>
        /// 新增解析列 有买家留言 不加属性SKU
        /// </summary>
        private void Transfrom(DataTable dt)
        {
            //获取当前列数
            var count = dt.Columns.Count;
            dt.Columns.Add("商品名称", Type.GetType("System.String"));//19
            dt.Columns.Add("商品数量", Type.GetType("System.String"));
            dt.Columns.Add("收货电话", Type.GetType("System.String"));
            dt.Columns.Add("商品链接", Type.GetType("System.String"));
            dt.Columns.Add("属性SKU", Type.GetType("System.String"));

            var dtSuccess = dt.Clone();
            var dtFail = dt.Clone();

            int i = 1;
            foreach (DataRow row in dt.Rows)//逐个读取单元格的内容；
            {
                if (!string.IsNullOrWhiteSpace(row[30].ToString()))
                {
                    row[count] = "客户留言,请手动下单";
                    dtFail.Rows.Add(row.ItemArray);
                    continue;
                }
                //商品名称
                row[count] = row[18];
                //商品数量
                var number = (Convert.ToInt32(row[20]) % 100);
                if (number > 0)
                {
                    //无法整除
                    row[count + 1] = "数量不正确,请核对";
                    dtFail.Rows.Add(row.ItemArray);
                    continue;
                }
                else
                {
                    row[count + 1] = (Convert.ToInt32(row[20]) / 100).ToString();
                }
                //收货电话
                row[count + 2] = row[17];
                //商品链接
                row[count + 3] = TxtBuyUrl.Text;
                //属性SKU
                var productName = row[18].ToString();
                if (productName.Contains("100支黑色"))
                {
                    row[count + 4] = "100支中性笔【黑色】";
                    dtSuccess.Rows.Add(row.ItemArray);
                    continue;
                }
                if (productName.Contains("100支红色 "))
                {
                    row[count + 4] = "100支中性笔【红色】";
                    dtSuccess.Rows.Add(row.ItemArray);
                    continue;
                }
                if (productName.Contains("100支蓝色"))
                {
                    row[count + 4] = "100支中性笔【蓝色】";
                    dtSuccess.Rows.Add(row.ItemArray);
                    continue;
                }
                if (productName.Contains("80黑"))
                {
                    row[count + 4] = "【80黑+10蓝10红】中性笔】";
                    dtSuccess.Rows.Add(row.ItemArray);
                    continue;
                }
                if (productName.Contains("90支黑"))
                {
                    row[count + 4] = "【90黑色+10红色】中性笔";
                    dtSuccess.Rows.Add(row.ItemArray);
                    continue;
                }

                if (productName.Contains("50支黑") || productName.Contains("50黑"))
                {
                    if (productName.Contains("50支红") || productName.Contains("50红"))
                    {
                        row[count + 4] = "【50支黑色+50红色】中性笔";
                        dtSuccess.Rows.Add(row.ItemArray);
                        continue;
                    }
                    if (productName.Contains("50支蓝") || productName.Contains("50蓝"))
                    {
                        row[count + 4] = "【50支黑色+50蓝色】中性笔";
                        dtSuccess.Rows.Add(row.ItemArray);
                        continue;
                    }
                }
                if ((productName.Contains("50支蓝") || productName.Contains("50蓝")) && (productName.Contains("50支红") || productName.Contains("50红")))
                {
                    row[count + 4] = "【50支红色+50蓝色】中性笔";
                    dtSuccess.Rows.Add(row.ItemArray);
                    continue;
                }
                row[count + 4] = "商品选择不存在";
                dtFail.Rows.Add(row.ItemArray);
                i++;
            }
            LblSuccessLine.Text = "当前总共" + dtSuccess.Rows.Count.ToString() + "行";
            LblFailLine.Text = "当前总共" + dtFail.Rows.Count.ToString() + "行";

            BindDate(DGVExportSuccess, dtSuccess);
            BindDate(DGVExportFail, dtFail);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="dgv"></param>
        /// <returns></returns>
        public DataTable GetDgvToTable(DataGridView dgv)
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

        /// <summary>
        /// 导出失败数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnExportFail_Click(object sender, EventArgs e)
        {
            var dt = GetDgvToTable(DGVExportFail);
            ExcelHelper.DataTableToExcel(dt, "导出成功" + DateTime.Now.ToString("yyyyMMddhhmmss"));
            MessageBox.Show("导出成功");
        }

        /// <summary>
        /// 导出成功数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnExportSuccess_Click(object sender, EventArgs e)
        {
            var dt = GetDgvToTable(DGVExportSuccess);
            ExcelHelper.DataTableToExcel(dt, "导出成功" + DateTime.Now.ToString("yyyyMMddhhmmss"));
            MessageBox.Show("导出成功");
        }
    }
}