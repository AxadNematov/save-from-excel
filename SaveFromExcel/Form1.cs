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
        public Form1()
        {
            InitializeComponent();
            initTables();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XLS files (*.xls, *.xlt)|*.xls;*.xlt|XLSX files (*.xlsx, *.xlsm, *.xltx, *.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|ODS files (*.ods, *.ots)|*.ods;*.ots|CSV files (*.csv, *.tsv)|*.csv;*.tsv|HTML files (*.html, *.htm)|*.html;*.htm";
            openFileDialog.FilterIndex = 2;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileInfo file = new FileInfo(openFileDialog.FileName);
                ExcelPackage excel = new ExcelPackage(file);
                var worksheet = excel.Workbook.Worksheets[1];

                dataGridView1.RowCount = worksheet.Dimension.Rows;
                for (int k = 0; k < worksheet.Dimension.Rows - 1; k++)
                {
                    dataGridView1.Rows[k].Cells[1].Value = k + 1;
                    dataGridView1.Rows[k].Cells[2].Value = worksheet.Cells[k + 2, 2].Value;
                    
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
    }
}
