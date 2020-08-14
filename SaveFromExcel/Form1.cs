using System;
using System.IO;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SqlServer.Server;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Runtime.InteropServices;
using System.Threading;
using OfficeOpenXml;
using System.Globalization;
using System.Drawing.Drawing2D;

namespace SaveFromExcel
{
    public partial class Form1 : Form
    {
        string conString = "Server=167.86.73.27; Database=lcdatabase; User Id=sa; Password=locked123$";
        public Form1()
        {
            InitializeComponent();
            initTables();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFirstType();
        }

        public void openFirstType()
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XLS files (*.xls, *.xlt)|*.xls;*.xlt|XLSX files (*.xlsx, *.xlsm, *.xltx, *.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|ODS files (*.ods, *.ots)|*.ods;*.ots|CSV files (*.csv, *.tsv)|*.csv;*.tsv|HTML files (*.html, *.htm)|*.html;*.htm";
            openFileDialog.FilterIndex = 2;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileInfo file = new FileInfo(openFileDialog.FileName);
                ExcelPackage excel = new ExcelPackage(file);
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                var worksheet = excel.Workbook.Worksheets[0];

                dataGridView1.RowCount = 0;

                dataGridView1.RowCount = worksheet.Dimension.Rows - 1;
                for (int k = 0; k < worksheet.Dimension.Rows - 1; k++)
                {
                    dataGridView1.Rows[k].Cells[0].Value = worksheet.Cells[k + 2, 1].Value;
                    dataGridView1.Rows[k].Cells[1].Value = worksheet.Cells[k + 2, 2].Value;
                    dataGridView1.Rows[k].Cells[2].Value = worksheet.Cells[k + 2, 3].Value;
                    dataGridView1.Rows[k].Cells[3].Value = worksheet.Cells[k + 2, 4].Value;
                    dataGridView1.Rows[k].Cells[4].Value = worksheet.Cells[k + 2, 5].Value;
                    dataGridView1.Rows[k].Cells[5].Value = worksheet.Cells[k + 2, 6].Value;
                    dataGridView1.Rows[k].Cells[6].Value = worksheet.Cells[k + 2, 7].Value;
                    dataGridView1.Rows[k].Cells[7].Value = worksheet.Cells[k + 2, 8].Value;
                    dataGridView1.Rows[k].Cells[8].Value = worksheet.Cells[k + 2, 9].Value;
                    dataGridView1.Rows[k].Cells[9].Value = worksheet.Cells[k + 2, 10].Value;
                    dataGridView1.Rows[k].Cells[10].Value = worksheet.Cells[k + 2, 11].Value;
                    dataGridView1.Rows[k].Cells[11].Value = worksheet.Cells[k + 2, 12].Value;
                    dataGridView1.Rows[k].Cells[12].Value = worksheet.Cells[k + 2, 13].Value;
                    dataGridView1.Rows[k].Cells[13].Value = worksheet.Cells[k + 2, 14].Value;
                    dataGridView1.Rows[k].Cells[14].Value = worksheet.Cells[k + 2, 15].Value;
                    dataGridView1.Rows[k].Cells[15].Value = worksheet.Cells[k + 2, 16].Value;
                    dataGridView1.Rows[k].Cells[16].Value = worksheet.Cells[k + 2, 17].Value;
                    dataGridView1.Rows[k].Cells[17].Value = worksheet.Cells[k + 2, 18].Value;
                    dataGridView1.Rows[k].Cells[18].Value = worksheet.Cells[k + 2, 19].Value;
                    dataGridView1.Rows[k].Cells[19].Value = worksheet.Cells[k + 2, 20].Value;
                    dataGridView1.Rows[k].Cells[20].Value = worksheet.Cells[k + 2, 21].Value;
                    dataGridView1.Rows[k].Cells[21].Value = worksheet.Cells[k + 2, 22].Value;
                    dataGridView1.Rows[k].Cells[22].Value = worksheet.Cells[k + 2, 23].Value;
                    dataGridView1.Rows[k].Cells[23].Value = worksheet.Cells[k + 2, 24].Value;
                    dataGridView1.Rows[k].Cells[24].Value = worksheet.Cells[k + 2, 25].Value;
                    dataGridView1.Rows[k].Cells[25].Value = worksheet.Cells[k + 2, 26].Value;
                    dataGridView1.Rows[k].Cells[26].Value = worksheet.Cells[k + 2, 27].Value;
                }
            }
        }

        public void openSecondType()
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XLS files (*.xls, *.xlt)|*.xls;*.xlt|XLSX files (*.xlsx, *.xlsm, *.xltx, *.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|ODS files (*.ods, *.ots)|*.ods;*.ots|CSV files (*.csv, *.tsv)|*.csv;*.tsv|HTML files (*.html, *.htm)|*.html;*.htm";
            openFileDialog.FilterIndex = 2;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileInfo file = new FileInfo(openFileDialog.FileName);
                ExcelPackage excel = new ExcelPackage(file);
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                var worksheet = excel.Workbook.Worksheets[0];

                dataGridView2.RowCount = 0;

                dataGridView2.RowCount = worksheet.Dimension.Rows - 1;
                for (int k = 0; k < worksheet.Dimension.Rows - 1; k++)
                {
                    dataGridView2.Rows[k].Cells[0].Value = worksheet.Cells[k + 2, 1].Value;
                    dataGridView2.Rows[k].Cells[1].Value = worksheet.Cells[k + 2, 2].Value;
                    dataGridView2.Rows[k].Cells[2].Value = worksheet.Cells[k + 2, 3].Value;
                    dataGridView2.Rows[k].Cells[3].Value = worksheet.Cells[k + 2, 4].Value;
                    dataGridView2.Rows[k].Cells[4].Value = worksheet.Cells[k + 2, 5].Value;
                    dataGridView2.Rows[k].Cells[5].Value = worksheet.Cells[k + 2, 6].Value;
                    dataGridView2.Rows[k].Cells[6].Value = worksheet.Cells[k + 2, 7].Value;
                    dataGridView2.Rows[k].Cells[7].Value = worksheet.Cells[k + 2, 8].Value;
                    dataGridView2.Rows[k].Cells[8].Value = worksheet.Cells[k + 2, 9].Value;
                    dataGridView2.Rows[k].Cells[9].Value = worksheet.Cells[k + 2, 10].Value;
                    dataGridView2.Rows[k].Cells[10].Value = worksheet.Cells[k + 2, 11].Value;
                    dataGridView2.Rows[k].Cells[11].Value = worksheet.Cells[k + 2, 12].Value;
                    dataGridView2.Rows[k].Cells[12].Value = worksheet.Cells[k + 2, 13].Value;
                    dataGridView2.Rows[k].Cells[13].Value = worksheet.Cells[k + 2, 14].Value;
                    dataGridView2.Rows[k].Cells[14].Value = worksheet.Cells[k + 2, 15].Value;
                    dataGridView2.Rows[k].Cells[15].Value = worksheet.Cells[k + 2, 16].Value;
                    dataGridView2.Rows[k].Cells[16].Value = worksheet.Cells[k + 2, 17].Value;
                    dataGridView2.Rows[k].Cells[17].Value = worksheet.Cells[k + 2, 18].Value;
                    dataGridView2.Rows[k].Cells[18].Value = worksheet.Cells[k + 2, 19].Value;
                    dataGridView2.Rows[k].Cells[19].Value = worksheet.Cells[k + 2, 20].Value;
                    dataGridView2.Rows[k].Cells[20].Value = worksheet.Cells[k + 2, 21].Value;
                    dataGridView2.Rows[k].Cells[21].Value = worksheet.Cells[k + 2, 22].Value;
                    dataGridView2.Rows[k].Cells[22].Value = worksheet.Cells[k + 2, 23].Value;
                    dataGridView2.Rows[k].Cells[23].Value = worksheet.Cells[k + 2, 24].Value;
                    dataGridView2.Rows[k].Cells[24].Value = worksheet.Cells[k + 2, 25].Value;
                    dataGridView2.Rows[k].Cells[25].Value = worksheet.Cells[k + 2, 26].Value;
                    dataGridView2.Rows[k].Cells[26].Value = worksheet.Cells[k + 2, 27].Value;
                }
            }
        }

        public void initTables()
        {
            dataGridView1.RowCount = 0;
            dataGridView1.ColumnCount = 27;

            dataGridView1.Columns[0].HeaderText = "#";
            dataGridView1.Columns[1].HeaderText = "Стандарт";
            dataGridView1.Columns[2].HeaderText = "Tok";
            dataGridView1.Columns[3].HeaderText = "Серия";
            dataGridView1.Columns[4].HeaderText = "Аппарат ";
            dataGridView1.Columns[5].HeaderText = "Модель";
            dataGridView1.Columns[6].HeaderText = "Отключ. Способность";
            dataGridView1.Columns[7].HeaderText = "Pole ";
            dataGridView1.Columns[8].HeaderText = "Кривая";
            dataGridView1.Columns[9].HeaderText = "А";
            dataGridView1.Columns[10].HeaderText = "mA";
            dataGridView1.Columns[11].HeaderText = "Сторона";
            dataGridView1.Columns[12].HeaderText = "REAR";
            dataGridView1.Columns[13].HeaderText = "Вариант";
            dataGridView1.Columns[14].HeaderText = "Iten code";
            dataGridView1.Columns[15].HeaderText = "Full name";
            dataGridView1.Columns[16].HeaderText = "Price";
            dataGridView1.Columns[17].HeaderText = "SIZE";
            dataGridView1.Columns[18].HeaderText = "Tочный вес";
            dataGridView1.Columns[19].HeaderText = "BRUTTO";
            dataGridView1.Columns[20].HeaderText = "NETTO";
            dataGridView1.Columns[21].HeaderText = "Mbox CBM";
            dataGridView1.Columns[22].HeaderText = "Mbox Qty";
            dataGridView1.Columns[23].HeaderText = "MBox Weight";
            dataGridView1.Columns[24].HeaderText = "QTY MBox/PBox";
            dataGridView1.Columns[25].HeaderText = "PBox Weight";
            dataGridView1.Columns[26].HeaderText = "Pbox CBM";

            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[15].Width = 200;

            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            dataGridView2.RowCount = 0;
            dataGridView2.ColumnCount = 27;

            dataGridView2.Columns[0].HeaderText = "#";
            dataGridView2.Columns[1].HeaderText = "Стандарт";
            dataGridView2.Columns[2].HeaderText = "Tok";
            dataGridView2.Columns[3].HeaderText = "Серия";
            dataGridView2.Columns[4].HeaderText = "Аппарат ";
            dataGridView2.Columns[5].HeaderText = "Модель";
            dataGridView2.Columns[6].HeaderText = "Размер";
            dataGridView2.Columns[7].HeaderText = "Отключ. Способность";
            dataGridView2.Columns[8].HeaderText = "Расцепитель";
            dataGridView2.Columns[9].HeaderText = "А";
            dataGridView2.Columns[10].HeaderText = "Pole";
            dataGridView2.Columns[11].HeaderText = "Сторона";
            dataGridView2.Columns[12].HeaderText = "REAR";
            dataGridView2.Columns[13].HeaderText = "Вариант";
            dataGridView2.Columns[14].HeaderText = "Iten code";
            dataGridView2.Columns[15].HeaderText = "Full name";
            dataGridView2.Columns[16].HeaderText = "Price";
            dataGridView2.Columns[17].HeaderText = "SIZE";
            dataGridView2.Columns[18].HeaderText = "Tочный вес";
            dataGridView2.Columns[19].HeaderText = "BRUTTO";
            dataGridView2.Columns[20].HeaderText = "NETTO";
            dataGridView2.Columns[21].HeaderText = "Mbox CBM";
            dataGridView2.Columns[22].HeaderText = "Mbox Qty";
            dataGridView2.Columns[23].HeaderText = "MBox Weight";
            dataGridView2.Columns[24].HeaderText = "QTY MBox/PBox";
            dataGridView2.Columns[25].HeaderText = "PBox Weight";
            dataGridView2.Columns[26].HeaderText = "Pbox CBM";

            dataGridView2.Columns[0].Width = 50;
            dataGridView2.Columns[15].Width = 200;

            foreach (DataGridViewColumn col in dataGridView2.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openSecondType();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView2.RowCount = 0;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.RowCount = 0;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if(dataGridView1.RowCount>0)
            {
                SaveFirstType();
            }
        }

        public void SaveFirstType()
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();

            string standard = "";
            string tok = "";
            string series = "";
            string apparat = "";
            string model = "";
            string disable_ability = "";
            string role = "";
            string curve = "";
            int a = 0;
            int ma = 0;
            string side = "";
            string rear = "";
            string variant = "";
            string item_code = "";
            string name = "";
            double price = 0;
            string size = "";
            string exact_weight = "";
            string brutto = "";
            string netto = "";
            string mbox_cbm = "";
            int mbox_qty = 0;
            string mbox_weight = "";
            int qty_mbox_per_pbox = 0;
            string pbox_weight = "";
            string pbox_cbm = "";

            int i = 0;
            for (i = 0; i < dataGridView1.Rows.Count; i++)
            {
                try { standard = dataGridView1.Rows[i].Cells[1].Value.ToString(); } catch { standard = ""; };
                try { tok = dataGridView1.Rows[i].Cells[2].Value.ToString(); } catch { tok = ""; };
                try { series = dataGridView1.Rows[i].Cells[3].Value.ToString(); } catch { series = ""; };
                try { apparat = dataGridView1.Rows[i].Cells[4].Value.ToString(); } catch { apparat = ""; };
                try { model = dataGridView1.Rows[i].Cells[5].Value.ToString(); } catch { model = ""; };
                try { disable_ability = dataGridView1.Rows[i].Cells[6].Value.ToString(); } catch { disable_ability = ""; };
                try { role = dataGridView1.Rows[i].Cells[7].Value.ToString(); } catch { role = ""; };
                try { curve = dataGridView1.Rows[i].Cells[8].Value.ToString(); } catch { curve = ""; };
                try { a = Convert.ToInt32(dataGridView1.Rows[i].Cells[9].Value.ToString()); } catch { a = 0; };
                try { ma = Convert.ToInt32(dataGridView1.Rows[i].Cells[10].Value.ToString()); } catch { ma = 0; };
                try { side = dataGridView1.Rows[i].Cells[11].Value.ToString(); } catch { side = ""; };
                try { rear = dataGridView1.Rows[i].Cells[12].Value.ToString(); } catch { rear = ""; };
                try { variant = dataGridView1.Rows[i].Cells[13].Value.ToString(); } catch { variant = ""; };
                try { item_code = dataGridView1.Rows[i].Cells[14].Value.ToString(); } catch { item_code = ""; };
                try { name = dataGridView1.Rows[i].Cells[15].Value.ToString(); } catch { name = ""; };
                try { price = Convert.ToDouble(dataGridView1.Rows[i].Cells[16].Value.ToString()); } catch { price = 0; };
                try { size = dataGridView1.Rows[i].Cells[17].Value.ToString(); } catch { size = ""; };
                try { exact_weight = dataGridView1.Rows[i].Cells[18].Value.ToString(); } catch { exact_weight = ""; };
                try { brutto = dataGridView1.Rows[i].Cells[19].Value.ToString(); } catch { brutto = ""; };
                try { netto = dataGridView1.Rows[i].Cells[20].Value.ToString(); } catch { netto = ""; };
                try { mbox_cbm = dataGridView1.Rows[i].Cells[21].Value.ToString(); } catch { mbox_cbm = ""; };
                try { mbox_qty = Convert.ToInt32(dataGridView1.Rows[i].Cells[22].Value.ToString()); } catch { mbox_qty = 0; };
                try { mbox_weight = dataGridView1.Rows[i].Cells[23].Value.ToString(); } catch { mbox_weight = ""; };
                try { qty_mbox_per_pbox = Convert.ToInt32(dataGridView1.Rows[i].Cells[24].Value.ToString()); } catch { qty_mbox_per_pbox = 0; };
                try { pbox_weight = dataGridView1.Rows[i].Cells[25].Value.ToString(); } catch { pbox_weight = ""; };
                try { pbox_cbm = dataGridView1.Rows[i].Cells[26].Value.ToString(); } catch { pbox_cbm = ""; };

                //
                SqlCommand command1 = new SqlCommand();
                command1.Connection = connection;
                command1.CommandType = CommandType.Text;
                command1.CommandText = "INSERT INTO items_first_type(standard, tok, series, apparat, model, disable_ability, role, curve, a, ma, side, rear, variant, item_code, name, price, size, exact_weight, brutto, netto, mbox_cbm, mbox_qty, mbox_weight, qty_mbox_per_pbox, pbox_weight, pbox_cbm) VALUES(@standard, @tok, @series, @apparat, @model, @disable_ability, @role, @curve, @a, @ma, @side, @rear, @variant, @item_code, @name, @price, @size, @exact_weight, @brutto, @netto, @mbox_cbm, @mbox_qty, @mbox_weight, @qty_mbox_per_pbox, @pbox_weight, @pbox_cbm)";
                command1.Parameters.AddWithValue("@standard", standard);
                command1.Parameters.AddWithValue("@tok", tok);
                command1.Parameters.AddWithValue("@series", series);
                command1.Parameters.AddWithValue("@apparat", apparat);
                command1.Parameters.AddWithValue("@model", model);
                command1.Parameters.AddWithValue("@disable_ability", disable_ability);
                command1.Parameters.AddWithValue("@role", role);
                command1.Parameters.AddWithValue("@curve", curve);
                command1.Parameters.AddWithValue("@a", a);
                command1.Parameters.AddWithValue("@ma", ma);
                command1.Parameters.AddWithValue("@side", side);
                command1.Parameters.AddWithValue("@rear", rear);
                command1.Parameters.AddWithValue("@variant", variant);
                command1.Parameters.AddWithValue("@item_code", item_code);
                command1.Parameters.AddWithValue("@name", name);
                command1.Parameters.AddWithValue("@price", price);
                command1.Parameters.AddWithValue("@size", size);
                command1.Parameters.AddWithValue("@exact_weight", exact_weight);
                command1.Parameters.AddWithValue("@brutto", brutto);
                command1.Parameters.AddWithValue("@netto", netto);
                command1.Parameters.AddWithValue("@mbox_cbm", mbox_cbm);
                command1.Parameters.AddWithValue("@mbox_qty", mbox_qty);
                command1.Parameters.AddWithValue("@mbox_weight", mbox_weight);
                command1.Parameters.AddWithValue("@qty_mbox_per_pbox", qty_mbox_per_pbox);
                command1.Parameters.AddWithValue("@pbox_weight", pbox_weight);
                command1.Parameters.AddWithValue("@pbox_cbm", pbox_cbm);

                command1.ExecuteNonQuery();
            }
            connection.Close();
            MessageBox.Show("OK", "Сообщение",
            MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        public void SaveSecondType()
        {
            SqlConnection connection = new SqlConnection(conString);
            connection.Open();

            string standard = "";
            string tok = "";
            string series = "";
            string apparat = "";
            string model = "";
            int capacity = 0;
            string disable_ability = "";
            string release = "";
            int a = 0;
            string role = "";
            string side = "";
            string rear = "";
            string variant = "";
            string item_code = "";
            string name = "";
            double price = 0;
            string size = "";
            string exact_weight = "";
            string brutto = "";
            string netto = "";
            string mbox_cbm = "";
            int mbox_qty = 0;
            string mbox_weight = "";
            int qty_mbox_per_pbox = 0;
            string pbox_weight = "";
            string pbox_cbm = "";

            int i = 0;
            for (i = 0; i < dataGridView2.Rows.Count; i++)
            {
                try { standard = dataGridView2.Rows[i].Cells[1].Value.ToString(); } catch { standard = ""; };
                try { tok = dataGridView2.Rows[i].Cells[2].Value.ToString(); } catch { tok = ""; };
                try { series = dataGridView2.Rows[i].Cells[3].Value.ToString(); } catch { series = ""; };
                try { apparat = dataGridView2.Rows[i].Cells[4].Value.ToString(); } catch { apparat = ""; };
                try { model = dataGridView2.Rows[i].Cells[5].Value.ToString(); } catch { model = ""; };
                try { capacity = Convert.ToInt32(dataGridView2.Rows[i].Cells[6].Value.ToString()); } catch { capacity = 0; };
                try { disable_ability = dataGridView2.Rows[i].Cells[7].Value.ToString(); } catch { disable_ability = ""; };
                try { release = dataGridView2.Rows[i].Cells[8].Value.ToString(); } catch { release = ""; };
                try { a = Convert.ToInt32(dataGridView2.Rows[i].Cells[9].Value.ToString()); } catch { a = 0; };
                try { role = dataGridView2.Rows[i].Cells[10].Value.ToString(); } catch { role = ""; };
                try { side = dataGridView2.Rows[i].Cells[11].Value.ToString(); } catch { side = ""; };
                try { rear = dataGridView2.Rows[i].Cells[12].Value.ToString(); } catch { rear = ""; };
                try { variant = dataGridView2.Rows[i].Cells[13].Value.ToString(); } catch { variant = ""; };
                try { item_code = dataGridView2.Rows[i].Cells[14].Value.ToString(); } catch { item_code = ""; };
                try { name = dataGridView2.Rows[i].Cells[15].Value.ToString(); } catch { name = ""; };
                try { price = Convert.ToDouble(dataGridView2.Rows[i].Cells[16].Value.ToString()); } catch { price = 0; };
                try { size = dataGridView2.Rows[i].Cells[17].Value.ToString(); } catch { size = ""; };
                try { exact_weight = dataGridView2.Rows[i].Cells[18].Value.ToString(); } catch { exact_weight = ""; };
                try { brutto = dataGridView2.Rows[i].Cells[19].Value.ToString(); } catch { brutto = ""; };
                try { netto = dataGridView2.Rows[i].Cells[20].Value.ToString(); } catch { netto = ""; };
                try { mbox_cbm = dataGridView2.Rows[i].Cells[21].Value.ToString(); } catch { mbox_cbm = ""; };
                try { mbox_qty = Convert.ToInt32(dataGridView2.Rows[i].Cells[22].Value.ToString()); } catch { mbox_qty = 0; };
                try { mbox_weight = dataGridView2.Rows[i].Cells[23].Value.ToString(); } catch { mbox_weight = ""; };
                try { qty_mbox_per_pbox = Convert.ToInt32(dataGridView2.Rows[i].Cells[24].Value.ToString()); } catch { qty_mbox_per_pbox = 0; };
                try { pbox_weight = dataGridView2.Rows[i].Cells[25].Value.ToString(); } catch { pbox_weight = ""; };
                try { pbox_cbm = dataGridView2.Rows[i].Cells[26].Value.ToString(); } catch { pbox_cbm = ""; };

                //
                SqlCommand command1 = new SqlCommand();
                command1.Connection = connection;
                command1.CommandType = CommandType.Text;
                command1.CommandText = "INSERT INTO items_second_type(standard, tok, series, apparat, model, capacity, disable_ability, release, a, role, side, rear, variant, item_code, name, price, size, exact_weight, brutto, netto, mbox_cbm, mbox_qty, mbox_weight, qty_mbox_per_pbox, pbox_weight, pbox_cbm) VALUES(@standard, @tok, @series, @apparat, @model, @capacity,  @disable_ability, @release, @a, @role, @side, @rear, @variant, @item_code, @name, @price, @size, @exact_weight, @brutto, @netto, @mbox_cbm, @mbox_qty, @mbox_weight, @qty_mbox_per_pbox, @pbox_weight, @pbox_cbm)";
                command1.Parameters.AddWithValue("@standard", standard);
                command1.Parameters.AddWithValue("@tok", tok);
                command1.Parameters.AddWithValue("@series", series);
                command1.Parameters.AddWithValue("@apparat", apparat);
                command1.Parameters.AddWithValue("@model", model);
                command1.Parameters.AddWithValue("@capacity", capacity);
                command1.Parameters.AddWithValue("@disable_ability", disable_ability);
                command1.Parameters.AddWithValue("@release", release);
                command1.Parameters.AddWithValue("@a", a);
                command1.Parameters.AddWithValue("@role", role);
                command1.Parameters.AddWithValue("@side", side);
                command1.Parameters.AddWithValue("@rear", rear);
                command1.Parameters.AddWithValue("@variant", variant);
                command1.Parameters.AddWithValue("@item_code", item_code);
                command1.Parameters.AddWithValue("@name", name);
                command1.Parameters.AddWithValue("@price", price);
                command1.Parameters.AddWithValue("@size", size);
                command1.Parameters.AddWithValue("@exact_weight", exact_weight);
                command1.Parameters.AddWithValue("@brutto", brutto);
                command1.Parameters.AddWithValue("@netto", netto);
                command1.Parameters.AddWithValue("@mbox_cbm", mbox_cbm);
                command1.Parameters.AddWithValue("@mbox_qty", mbox_qty);
                command1.Parameters.AddWithValue("@mbox_weight", mbox_weight);
                command1.Parameters.AddWithValue("@qty_mbox_per_pbox", qty_mbox_per_pbox);
                command1.Parameters.AddWithValue("@pbox_weight", pbox_weight);
                command1.Parameters.AddWithValue("@pbox_cbm", pbox_cbm);

                command1.ExecuteNonQuery();
            }
            connection.Close();
            MessageBox.Show("OK", "Сообщение",
            MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if(dataGridView2.RowCount>0)
            {
                SaveSecondType();
            }
        }
    }
}
