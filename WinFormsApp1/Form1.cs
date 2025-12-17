using ClosedXML.Excel;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using System.Data;
using System.Text;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        private XLWorkbook workbook;
        private IXLWorksheet sheet;
        private BindingSource bs = new BindingSource();
        private CheckBox headerCheckBox;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel Files (*.xlsx)|*.xlsx";

                if (ofd.ShowDialog() != DialogResult.OK)
                    return;

                workbook = new XLWorkbook(ofd.FileName);
                sheet = workbook.Worksheet(1);

                DataTable dt = new DataTable();

                // 1) Create columns from first row
                foreach (var cell in sheet.FirstRow().CellsUsed())
                {
                    dt.Columns.Add(cell.GetString().Trim());
                }

                // 2) Read Excel rows
                foreach (var row in sheet.RowsUsed().Skip(1))
                {
                    DataRow dr = dt.NewRow();

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        dr[i] = row.Cell(i + 1).GetValue<string>();
                    }

                    dt.Rows.Add(dr);
                }

                // 3) Convert Status column values (NO extra column)
                if (dt.Columns.Contains("status"))
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        string val = row["Status"]?.ToString().Trim();

                        if (val == "1")
                            row["status"] = "مدرج";
                        else if (val == "0")
                            row["status"] = "غير مدرج";
                    }
                }

                dataGridView1.AutoGenerateColumns = true;
                bs.DataSource = dt;
                dataGridView1.DataSource = bs;
                // Replace Status column with ComboBox
                if (dataGridView1.Columns.Contains("Status"))
                {
                    int colIndex = dataGridView1.Columns["Status"].Index;

                    // Remove the original text column
                    dataGridView1.Columns.Remove("Status");

                    // Create ComboBox column
                    DataGridViewComboBoxColumn statusCombo = new DataGridViewComboBoxColumn();
                    statusCombo.Name = "Status";
                    statusCombo.HeaderText = "Status";
                    statusCombo.DataPropertyName = "Status"; // bind to same column
                    statusCombo.DropDownWidth = 100;
                    statusCombo.FlatStyle = FlatStyle.Standard;

                    statusCombo.Items.Add("مدرج");
                    statusCombo.Items.Add("غير مدرج");

                    statusCombo.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;

                    // Insert it back in the same position
                    dataGridView1.Columns.Insert(colIndex, statusCombo);
                }
                if (dt.Columns.Contains("Status"))
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        string raw = row["Status"]?.ToString().Trim();

                        if (raw == "1")
                            row["Status"] = "مدرج";
                        else if (raw == "0")
                            row["Status"] = "غير مدرج";
                        else if (string.IsNullOrEmpty(raw))
                            row["Status"] = "غير مدرج"; // safe default
                    }
                }

                // Add checkbox column (only once)
                // Add checkbox column (only once)
                if (!dataGridView1.Columns.Contains("chk"))
                {
                    DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
                    chk.Name = "chk";
                    chk.HeaderText = "";
                    chk.Width = 40;
                    chk.ReadOnly = false;

                    dataGridView1.Columns.Insert(0, chk);
                }
                AddHeaderCheckBox();

                UpdateRowCount();
                MessageBox.Show("Excel loaded successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Excel: " + ex.Message);
            }


        }

        private async void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            string folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "ArabicPDFs"
            );
            Directory.CreateDirectory(folder);

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                bool isChecked = Convert.ToBoolean(row.Cells["chk"].Value ?? false);
                if (!isChecked) continue;

                string html = GenerateArabicHtml(row);
                string clientName = "Unknown";

                if (row.DataGridView.Columns.Contains("Name"))
                {
                    clientName = row.Cells["Name"].FormattedValue?.ToString();
                }

                clientName = MakeSafeFileName(clientName);

                string pdfPath = Path.Combine(folder, $"{clientName}.pdf");

                await HtmlToPdfAsync(html, pdfPath);
            }

            MessageBox.Show("Arabic PDFs created successfully.");
        }

        private string GenerateArabicHtml(DataGridViewRow row)
        {
            // 1️⃣ Get client name safely
            string clientName = "غير معروف";

            if (row.DataGridView.Columns.Contains("Name"))
            {
                clientName = row.Cells["Name"].FormattedValue?.ToString();
                if (string.IsNullOrWhiteSpace(clientName))
                    clientName = "غير معروف";
            }

            StringBuilder sb = new StringBuilder();

            sb.Append($@"
<!DOCTYPE html>
<html lang='ar' dir='rtl'>
<head>
<meta charset='utf-8'>
<style>
    body {{
        font-family: 'Segoe UI', Tahoma, Arial;
        background-color: #f4f6f8;
        margin: 0;
        padding: 40px;
        direction: rtl;
        text-align: right;
    }}

    .container {{
        max-width: 800px;
        margin: auto;
        background: #ffffff;
        padding: 30px;
        border-radius: 10px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1);
    }}

    .title {{
        text-align: center;
        font-size: 22px;
        font-weight: bold;
        margin-bottom: 30px;
        color: #333;
    }}

    .row {{
        display: flex;
        border-bottom: 1px solid #eee;
        padding: 12px 0;
    }}

    .label {{
        width: 35%;
        font-weight: bold;
        color: #555;
    }}

    .value {{
        width: 65%;
        color: #000;
    }}

    .footer {{
        text-align: center;
        margin-top: 30px;
        font-size: 12px;
        color: #888;
    }}
</style>
</head>
<body>

<div class='container'>
    <div class='title'>بيانات العميل : {clientName}</div>
");

            // 2️⃣ Render rows
            foreach (DataGridViewCell cell in row.Cells)
            {
                if (cell.OwningColumn.Name == "chk")
                    continue;

                sb.Append($@"
    <div class='row'>
        <div class='label'>{cell.OwningColumn.HeaderText}</div>
        <div class='value'>{cell.FormattedValue}</div>
    </div>");
            }

            sb.Append(@"
    <div class='footer'>
        تم إنشاء هذا الملف تلقائيًا
    </div>
</div>

</body>
</html>
");

            return sb.ToString();
        }

        private async Task HtmlToPdfAsync(string html, string pdfPath)
        {
            var tcs = new TaskCompletionSource<bool>();

            var webView = new Microsoft.Web.WebView2.WinForms.WebView2();
            webView.CreateControl(); // IMPORTANT
            await webView.EnsureCoreWebView2Async();

            string tempHtml = Path.Combine(
                Path.GetTempPath(),
                Guid.NewGuid().ToString() + ".html"
            );

            File.WriteAllText(tempHtml, html, Encoding.UTF8);

            webView.CoreWebView2.NavigationCompleted += async (s, e) =>
            {
                await webView.CoreWebView2.PrintToPdfAsync(pdfPath);
                tcs.SetResult(true);
            };

            webView.CoreWebView2.Navigate(tempHtml);

            await tcs.Task; // ⬅️ WAIT HERE (CRITICAL)

            webView.Dispose();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.EndEdit(); // commit combo edits

            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("No data to export.");
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Files (*.xlsx)|*.xlsx";
            sfd.FileName = "ExportedData.xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Data");

                int excelCol = 1;

                // 1️⃣ Write headers
                foreach (DataGridViewColumn col in dataGridView1.Columns)
                {
                    if (col.Name == "chk") // skip checkbox column
                        continue;

                    ws.Cell(1, excelCol).Value = col.HeaderText;
                    ws.Cell(1, excelCol).Style.Font.Bold = true;
                    excelCol++;
                }

                // 2️⃣ Write rows
                int excelRow = 2;

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    excelCol = 1;

                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.OwningColumn.Name == "chk")
                            continue;

                        ws.Cell(excelRow, excelCol).Value =
                            cell.FormattedValue?.ToString() ?? "";

                        excelCol++;
                    }

                    excelRow++;
                }

                ws.Columns().AdjustToContents();
                wb.SaveAs(sfd.FileName);
            }

            MessageBox.Show("Excel file exported successfully.");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView1.EndEdit(); // commit checkbox & combobox edits

            bool anyChecked = false;

            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Choose folder to save Excel files";

                if (fbd.ShowDialog() != DialogResult.OK)
                    return;

                string baseFolder = fbd.SelectedPath;

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    bool isChecked = Convert.ToBoolean(row.Cells["chk"].Value ?? false);
                    if (!isChecked)
                        continue;

                    anyChecked = true;

                    using (var wb = new XLWorkbook())
                    {
                        var ws = wb.Worksheets.Add("Data");

                        int excelCol = 1;

                        // 1️⃣ Headers
                        foreach (DataGridViewColumn col in dataGridView1.Columns)
                        {
                            if (col.Name == "chk")
                                continue;

                            ws.Cell(1, excelCol).Value = col.HeaderText;
                            ws.Cell(1, excelCol).Style.Font.Bold = true;
                            excelCol++;
                        }

                        // 2️⃣ Single row data
                        excelCol = 1;
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (cell.OwningColumn.Name == "chk")
                                continue;

                            ws.Cell(2, excelCol).Value =
                                cell.FormattedValue?.ToString() ?? "";

                            excelCol++;
                        }

                        ws.Columns().AdjustToContents();

                        // 3️⃣ Safe file name
                        string fileName = $"Row_{row.Index + 1}.xlsx";
                        string fullPath = Path.Combine(baseFolder, fileName);

                        wb.SaveAs(fullPath);
                    }
                }
            }

            if (!anyChecked)
                MessageBox.Show("No rows were selected.");
            else
                MessageBox.Show("Excel files created successfully.");
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (bs.DataSource == null)
                return;

            string text = txtSearch.Text.Replace("'", "''");

            if (string.IsNullOrWhiteSpace(text))
            {
                bs.RemoveFilter();
                return;
            }

            // Build filter dynamically for all columns except checkbox
            DataTable dt = (DataTable)bs.DataSource;

            var filters = new List<string>();

            foreach (DataColumn col in dt.Columns)
            {
                filters.Add($"CONVERT([{col.ColumnName}], 'System.String') LIKE '%{text}%'");
            }

            bs.Filter = string.Join(" OR ", filters);
        }

        private string MakeSafeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Unknown";

            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c.ToString(), "");
            }

            return name.Trim();
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name != "Status")
                return;

            if (e.Value == null)
                return;

            string status = e.Value.ToString();

            e.CellStyle.BackColor = Color.White;
            e.CellStyle.ForeColor = Color.Black;

            if (status == "مدرج")
            {
                e.CellStyle.BackColor = Color.FromArgb(46, 125, 50);
                e.CellStyle.ForeColor = Color.White;
            }
            else if (status == "غير مدرج")
            {
                e.CellStyle.BackColor = Color.FromArgb(198, 40, 40);
                e.CellStyle.ForeColor = Color.White;
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            if (dataGridView1.Columns[e.ColumnIndex].Name == "Status")
            {
                // Force redraw of this row immediately
                dataGridView1.InvalidateRow(e.RowIndex);
            }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView1.IsCurrentCellDirty)
            {
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void ApplyFilters()
        {
            if (bs.DataSource == null)
                return;

            List<string> filters = new List<string>();

            // 🔍 Text search (if you already have it)
            if (!string.IsNullOrWhiteSpace(txtSearch.Text))
            {
                string text = txtSearch.Text.Replace("'", "''");

                DataTable dt = (DataTable)bs.DataSource;
                var textFilters = new List<string>();

                foreach (DataColumn col in dt.Columns)
                {
                    textFilters.Add(
                        $"CONVERT([{col.ColumnName}], 'System.String') LIKE '%{text}%'"
                    );
                }

                filters.Add("(" + string.Join(" OR ", textFilters) + ")");
            }

            // 📅 Date range filter
            DateTime from = dateTimePicker1.Value.Date;
            DateTime to = dateTimePicker2.Value.Date.AddDays(1).AddTicks(-1); // include whole day

            string dateFilter =
                $"[Date] >= #{from:MM/dd/yyyy}# AND [Date] <= #{to:MM/dd/yyyy}#";

            filters.Add(dateFilter);

            // Apply everything
            bs.Filter = string.Join(" AND ", filters);
            UpdateRowCount();
        }

        private void AddHeaderCheckBox()
        {
            // Prevent adding it twice
            if (headerCheckBox != null)
                return;

            headerCheckBox = new CheckBox();
            headerCheckBox.Size = new Size(15, 15);
            headerCheckBox.BackColor = Color.Transparent;

            // Get header cell rectangle for chk column
            Rectangle rect = dataGridView1.GetCellDisplayRectangle(
                dataGridView1.Columns["chk"].Index, -1, true
            );

            // Center the checkbox in header cell
            headerCheckBox.Location = new Point(
                rect.X + (rect.Width - headerCheckBox.Width) / 2,
                rect.Y + (rect.Height - headerCheckBox.Height) / 2
            );

            headerCheckBox.CheckedChanged += HeaderCheckBox_CheckedChanged;

            dataGridView1.Controls.Add(headerCheckBox);
        }

        private void HeaderCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            bool isChecked = headerCheckBox.Checked;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // Respect filtering
                if (row.Visible)
                {
                    row.Cells["chk"].Value = isChecked;
                }
            }
        }

        private void UpdateRowCount()
        {
            int count = 0;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Visible)
                    count++;
            }

            label3.Text = $"Rows: {count}";

        }


   

     
    }

}
