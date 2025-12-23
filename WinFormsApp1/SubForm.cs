using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinFormsApp1
{
    public partial class SubForm : Form
    {
        public SubForm()
        {
            InitializeComponent();
        }

        public void SetData1(DataTable dt)
        {
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = dt;
            dataGridView1.ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private readonly string _subGridFile =
    Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "WinFormsApp1",
        "SubGrid.xlsx"
    );

        private readonly string _subFormFlag =
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "WinFormsApp1",
                "SubFormOpen.flag"
            );

        
        private void SaveSubGridSilently()
        {
            try
            {
                if (dataGridView1.Rows.Count == 0)
                    return;

                Directory.CreateDirectory(Path.GetDirectoryName(_subGridFile));

                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Data");

                int col = 1;

                foreach (DataGridViewColumn c in dataGridView1.Columns)
                {
                    ws.Cell(1, col).Value = c.HeaderText;
                    ws.Cell(1, col).Style.Font.Bold = true;
                    col++;
                }

                int row = 2;

                foreach (DataGridViewRow r in dataGridView1.Rows)
                {
                    col = 1;
                    foreach (DataGridViewCell cell in r.Cells)
                    {
                        ws.Cell(row, col).Value = cell.FormattedValue?.ToString();
                        col++;
                    }
                    row++;
                }

                ws.Columns().AdjustToContents();
                wb.SaveAs(_subGridFile);
            }
            catch
            {
                // silent by design
            }
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SaveSubGridSilently();   // 🔥 REQUIRED
            base.OnFormClosing(e);
        }
        public void SaveNow()
        {
            SaveSubGridSilently();
        }

    }

}
