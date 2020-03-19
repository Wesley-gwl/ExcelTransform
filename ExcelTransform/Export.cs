using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTransform
{
    public partial class FrmMain : Form
    {
        public Dictionary<string, string> dic = null;

        public FrmMain()
        {
            dic = new Dictionary<string, string>();
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
            else if (string.IsNullOrEmpty(TxtSKUUrl.Text))
            {
                MessageBox.Show("请先导入SKU规则");
            }
            else
            {
                var dt = ExcelHelper.ExcelToDataTable(filepath, null, true);
                Transfrom(dt);
                MessageBox.Show("转换成功，请查看成功数据和失败数据");
            }
        }

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
                var key = CalculateSKU(productName);
                if (key == "")
                {
                    row[count + 4] = "商品选择不存在";
                    dtFail.Rows.Add(row.ItemArray);
                }
                else
                {
                    row[count + 4] = key;
                    dtSuccess.Rows.Add(row.ItemArray);
                }
                i++;
            }
            LblSuccessLine.Text = "当前总共" + dtSuccess.Rows.Count.ToString() + "行";
            LblFailLine.Text = "当前总共" + dtFail.Rows.Count.ToString() + "行";

            BindDate(DGVExportSuccess, dtSuccess);
            BindDate(DGVExportFail, dtFail);
        }

        /// <summary>
        /// 导出失败数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnExportFail_Click(object sender, EventArgs e)
        {
            var dt = DataHelper.GetDgvToTable(DGVExportFail);
            if (dt.Rows.Count > 0)
            {
                ExcelHelper.DataTableToExcel(dt, "导出转换失败数据" + DateTime.Now.ToString("yyyyMMddhhmmss"));
                MessageBox.Show("导出成功");
            }
            else
            {
                MessageBox.Show("无需导出");
            }
        }

        /// <summary>
        /// 导出成功数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnExportSuccess_Click(object sender, EventArgs e)
        {
            var dt = DataHelper.GetDgvToTable(DGVExportSuccess);
            if (dt.Rows.Count > 0)
            {
                ExcelHelper.DataTableToExcel(dt, "导出转换成功数据" + DateTime.Now.ToString("yyyyMMddhhmmss"));
                MessageBox.Show("导出成功");
            }
            else
            {
                MessageBox.Show("无需导出");
            }
        }

        /// <summary>
        /// 分批导出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSplitSuccessExport_Click(object sender, EventArgs e)
        {
            var dt = DataHelper.GetDgvToTable(DGVExportSuccess);
            var count = 9;
            //分批导出
            if (dt.Rows.Count > 0)
            {
                var dataSet = DataHelper.SplitDataTable(dt, count, 0, "属性SKU");
                for (int i = 0; i < dataSet.Tables.Count; i++)
                {
                    ExcelHelper.DataTableToExcel(dataSet.Tables[i], "导出转换成功数据" + DateTime.Now.ToString("yyyyMMddhhmmss") + "[" + (i + 1).ToString() + "]");
                }
                MessageBox.Show("导出成功");
            }
            else
            {
                MessageBox.Show("无需导出");
            }
        }

        /// <summary>
        /// 选择sku规则文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSelectSKU_Click(object sender, EventArgs e)
        {
            try
            {
                OFDSKUtxt.ShowDialog();
                var path = OFDSKUtxt.FileNames;  //获取openFileDialog控件选择的文件名数组
                var txt = TxtHelper.Read(path[0]);
                dic.Clear();
                TxtSKUUrl.Text = path[0];
                var temp = txt.Trim().Split(',');
                foreach (var item in temp)
                {
                    if (item != "")
                    {
                        var s = item.Split('@');
                        if (s.Count() > 1)
                        {
                            dic.Add(s[0], s[1]);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("文本解析失败,请检查格式" + ex.Message.ToString());
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="productName"></param>
        private string CalculateSKU(string productName)
        {
            foreach (var item in dic)
            {
                if (ResolverCondition(item.Key, productName))
                {
                    return item.Value;
                }
            }
            return "";
        }

        /// <summary>
        /// 转换计算
        /// </summary>
        /// <param name="key"></param>
        /// <param name="productName"></param>
        /// <returns></returns>
        private bool ResolverCondition(string key, string productName)
        {
            if (key.Contains("AND"))
            {
                var temp = Regex.Split(key, "AND", RegexOptions.IgnoreCase);
                foreach (var item in temp)
                {
                    item.Replace('(', ' ');
                    item.Replace(')', ' ');
                    item.Trim();
                    if (!ResolverCondition(item, productName))
                    {
                        return false;
                    }
                }
            }
            else if (key.Contains("OR"))
            {
                var temp = Regex.Split(key, "OR", RegexOptions.IgnoreCase);
                foreach (var item in temp)
                {
                    item.Trim();
                    if (ResolverCondition(item, productName))
                    {
                        return true;
                    }
                }
                return false;
            }
            return productName.Contains(key);
        }
    }
}