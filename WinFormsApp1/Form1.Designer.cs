namespace WinFormsApp1
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
           
            DataGridViewCellStyle dataGridViewCellStyle3 = new DataGridViewCellStyle();
            dataGridView1 = new DataGridView();
            buttonImport = new Button();
            buttonRowPdf = new Button();
            buttonAllExcel = new Button();
            buttonRowExcel = new Button();
            txtSearch = new TextBox();
            topPanel = new Panel();
            button1 = new Button();
            dateTimePicker2 = new DateTimePicker();
            label2 = new Label();
            dateTimePicker1 = new DateTimePicker();
            label1 = new Label();
            lblSearch = new Label();
            label3 = new Label();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            topPanel.SuspendLayout();
            SuspendLayout();
            // 
            // dataGridView1
            // 
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.BorderStyle = BorderStyle.None;
            dataGridViewCellStyle3.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = Color.FromArgb(230, 233, 236);
            dataGridViewCellStyle3.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            dataGridViewCellStyle3.ForeColor = SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = DataGridViewTriState.True;
            dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle3;
            dataGridView1.Dock = DockStyle.Fill;
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.Location = new Point(0, 141);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowTemplate.Height = 28;
            dataGridView1.Size = new Size(900, 379);
            dataGridView1.TabIndex = 0;
            dataGridView1.CellFormatting += dataGridView1_CellFormatting;
            dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
            dataGridView1.CurrentCellDirtyStateChanged += dataGridView1_CurrentCellDirtyStateChanged;
            // 
            // buttonImport
            // 
            buttonImport.Location = new Point(10, 12);
            buttonImport.Name = "buttonImport";
            buttonImport.Size = new Size(120, 35);
            buttonImport.TabIndex = 0;
            buttonImport.Text = "📥 Import Excel";
            buttonImport.Click += button1_Click;
            // 
            // buttonRowPdf
            // 
            buttonRowPdf.Location = new Point(164, 61);
            buttonRowPdf.Name = "buttonRowPdf";
            buttonRowPdf.Size = new Size(120, 35);
            buttonRowPdf.TabIndex = 3;
            buttonRowPdf.Text = "\U0001f9fe PDF / Row";
            buttonRowPdf.Click += button2_Click;
            // 
            // buttonAllExcel
            // 
            buttonAllExcel.Location = new Point(0, 61);
            buttonAllExcel.Name = "buttonAllExcel";
            buttonAllExcel.Size = new Size(140, 35);
            buttonAllExcel.TabIndex = 1;
            buttonAllExcel.Text = "📤 Export All Excel";
            buttonAllExcel.Click += button3_Click;
            // 
            // buttonRowExcel
            // 
            buttonRowExcel.Location = new Point(164, 13);
            buttonRowExcel.Name = "buttonRowExcel";
            buttonRowExcel.Size = new Size(120, 35);
            buttonRowExcel.TabIndex = 2;
            buttonRowExcel.Text = "📄 Excel / Row";
            buttonRowExcel.Click += button4_Click;
            // 
            // txtSearch
            // 
            txtSearch.Location = new Point(368, 10);
            txtSearch.Name = "txtSearch";
            txtSearch.PlaceholderText = "Type to search...";
            txtSearch.Size = new Size(200, 23);
            txtSearch.TabIndex = 5;
            txtSearch.TextChanged += textBox1_TextChanged;
            // 
            // topPanel
            // 
            topPanel.BackColor = Color.FromArgb(245, 247, 250);
            topPanel.Controls.Add(button1);
            topPanel.Controls.Add(dateTimePicker2);
            topPanel.Controls.Add(label2);
            topPanel.Controls.Add(dateTimePicker1);
            topPanel.Controls.Add(label1);
            topPanel.Controls.Add(buttonImport);
            topPanel.Controls.Add(buttonAllExcel);
            topPanel.Controls.Add(buttonRowExcel);
            topPanel.Controls.Add(buttonRowPdf);
            topPanel.Controls.Add(lblSearch);
            topPanel.Controls.Add(txtSearch);
            topPanel.Dock = DockStyle.Top;
            topPanel.Location = new Point(0, 0);
            topPanel.Name = "topPanel";
            topPanel.Padding = new Padding(10);
            topPanel.Size = new Size(900, 141);
            topPanel.TabIndex = 1;
            // 
            // button1
            // 
            button1.Location = new Point(616, 13);
            button1.Name = "button1";
            button1.Size = new Size(120, 35);
            button1.TabIndex = 10;
            button1.Text = "📄 Apply dateFilter";
            button1.Click += button1_Click_1;
            // 
            // dateTimePicker2
            // 
            dateTimePicker2.Location = new Point(368, 73);
            dateTimePicker2.Name = "dateTimePicker2";
            dateTimePicker2.Size = new Size(200, 23);
            dateTimePicker2.TabIndex = 9;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(317, 77);
            label2.Name = "label2";
            label2.Size = new Size(29, 15);
            label2.TabIndex = 8;
            label2.Text = "To : ";
            // 
            // dateTimePicker1
            // 
            dateTimePicker1.Location = new Point(368, 44);
            dateTimePicker1.Name = "dateTimePicker1";
            dateTimePicker1.Size = new Size(200, 23);
            dateTimePicker1.TabIndex = 7;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(317, 44);
            label1.Name = "label1";
            label1.Size = new Size(44, 15);
            label1.TabIndex = 6;
            label1.Text = "From : ";
            // 
            // lblSearch
            // 
            lblSearch.AutoSize = true;
            lblSearch.Location = new Point(317, 13);
            lblSearch.Name = "lblSearch";
            lblSearch.Size = new Size(45, 15);
            lblSearch.TabIndex = 4;
            lblSearch.Text = "Search:";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(801, 496);
            label3.Name = "label3";
            label3.Size = new Size(41, 15);
            label3.TabIndex = 5;
            label3.Text = "Rows :";
           
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(900, 520);
            Controls.Add(label3);
            Controls.Add(dataGridView1);
            Controls.Add(topPanel);
            Font = new Font("Segoe UI", 9F);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Excel Manager";
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            topPanel.ResumeLayout(false);
            topPanel.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private DataGridView dataGridView1;
        private Panel topPanel;
        private Button buttonImport;
        private Button buttonRowPdf;
        private Button buttonAllExcel;
        private Button buttonRowExcel;
        private TextBox txtSearch;
        private Label lblSearch;
        private Label label1;
        private DateTimePicker dateTimePicker2;
        private Label label2;
        private DateTimePicker dateTimePicker1;
        private Button button1;
        private Label label3;
    }
}
