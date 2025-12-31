using ClosedXML.Excel;
using Microsoft.Web.WebView2.WinForms;
using System.Data;
using System.Text.Json;

namespace WinFormsApp1
{
    public partial class SubForm : Form
    {
        private WebView2 webView;
        private DataTable _dt;

        private readonly string _subGridFile =
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "WinFormsApp1",
                "SubGrid.xlsx"
            );

        public SubForm()
        {
            InitializeComponent();

            // Hide WinForms grid (we don't use it)
            dataGridView1.Visible = false;

            webView = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(webView);
            webView.BringToFront();

            Load += async (_, __) =>
            {
                await InitWebViewAsync();
                SendDataToHtml();
            };
        }

        private async Task InitWebViewAsync()
        {
            await webView.EnsureCoreWebView2Async();

            string htmlPath = Path.Combine(Application.StartupPath, "UI", "subform.html");
            if (!File.Exists(htmlPath))
            {
                MessageBox.Show("SubForm HTML not found:\n" + htmlPath);
                return;
            }

            webView.CoreWebView2.Navigate(htmlPath);
        }

        public void SetData1(DataTable dt)
        {
            _dt = dt;
        }

        private void SendDataToHtml()
        {
            if (webView?.CoreWebView2 == null)
                return;

            if (_dt == null || _dt.Rows.Count == 0)
                return;

            var cols = _dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();
            var rows = new List<object>();

            foreach (DataRow r in _dt.Rows)
            {
                var obj = new Dictionary<string, string?>();
                foreach (string c in cols)
                    obj[c] = r[c]?.ToString();
                rows.Add(obj);
            }

            var payload = new { columns = cols, rows = rows };
            string json = JsonSerializer.Serialize(payload);

            webView.CoreWebView2.ExecuteScriptAsync($"renderSubGrid({json});");
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SaveSubGridSilently(); // required by your design
            base.OnFormClosing(e);
        }

        public void SaveNow() => SaveSubGridSilently();

        private void SaveSubGridSilently()
        {
            try
            {
                if (_dt == null || _dt.Rows.Count == 0)
                    return;

                Directory.CreateDirectory(Path.GetDirectoryName(_subGridFile)!);

                using var wb = new ClosedXML.Excel.XLWorkbook();
                var ws = wb.Worksheets.Add("Data");

                // headers
                for (int c = 0; c < _dt.Columns.Count; c++)
                {
                    ws.Cell(1, c + 1).Value = _dt.Columns[c].ColumnName;
                    ws.Cell(1, c + 1).Style.Font.Bold = true;
                }

                // rows
                for (int r = 0; r < _dt.Rows.Count; r++)
                {
                    for (int c = 0; c < _dt.Columns.Count; c++)
                        ws.Cell(r + 2, c + 1).Value = _dt.Rows[r][c]?.ToString() ?? "";
                }

                ws.Columns().AdjustToContents();
                wb.SaveAs(_subGridFile);
            }
            catch
            {
                // silent by design
            }
        }
    }
}
