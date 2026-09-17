namespace ExcelImporter;

partial class Form1
{
    private System.ComponentModel.IContainer components = null!;
    private Button btnAddExcel = null!;
    private Button btnClear = null!;
    private Button btnSend = null!;
    private DataGridView gridRecords = null!;
    private Label lblStatus = null!;
    private Label lblDropHint = null!;
    private Panel panelToolbar = null!;
    private Panel panelDropZone = null!;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }

        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        components = new System.ComponentModel.Container();

        panelToolbar = new Panel();
        btnAddExcel = new Button();
        btnClear = new Button();
        btnSend = new Button();
        lblStatus = new Label();
        panelDropZone = new Panel();
        lblDropHint = new Label();
        gridRecords = new DataGridView();

        panelToolbar.SuspendLayout();
        panelDropZone.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)gridRecords).BeginInit();
        SuspendLayout();

        panelToolbar.Dock = DockStyle.Top;
        panelToolbar.Height = 56;
        panelToolbar.Padding = new Padding(12, 10, 12, 10);
        panelToolbar.BackColor = SystemColors.Control;
        panelToolbar.Controls.Add(btnAddExcel);
        panelToolbar.Controls.Add(btnClear);
        panelToolbar.Controls.Add(btnSend);
        panelToolbar.Controls.Add(lblStatus);

        btnAddExcel.Text = "Προσθήκη Excel...";
        btnAddExcel.Location = new Point(12, 12);
        btnAddExcel.Size = new Size(160, 32);
        btnAddExcel.Click += btnAddExcel_Click;

        btnClear.Text = "Καθαρισμός";
        btnClear.Location = new Point(184, 12);
        btnClear.Size = new Size(110, 32);
        btnClear.Enabled = false;
        btnClear.Click += btnClear_Click;

        btnSend.Text = "Αποστολή";
        btnSend.Location = new Point(306, 12);
        btnSend.Size = new Size(120, 32);
        btnSend.Enabled = false;
        btnSend.Click += btnSend_Click;

        lblStatus.AutoSize = true;
        lblStatus.Location = new Point(440, 18);
        lblStatus.Text = "Εγγραφές: 0";

        panelDropZone.Dock = DockStyle.Fill;
        panelDropZone.BackColor = Color.FromArgb(245, 245, 245);
        panelDropZone.Padding = new Padding(24);
        panelDropZone.Controls.Add(lblDropHint);

        lblDropHint.Dock = DockStyle.Fill;
        lblDropHint.TextAlign = ContentAlignment.MiddleCenter;
        lblDropHint.Font = new Font("Segoe UI", 14F, FontStyle.Regular);
        lblDropHint.ForeColor = SystemColors.GrayText;
        lblDropHint.Text = "Σύρετε αρχείο Excel εδώ\r\n(.xls ή .xlsx)";

        gridRecords.AllowUserToAddRows = false;
        gridRecords.AllowUserToDeleteRows = false;
        gridRecords.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        gridRecords.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        gridRecords.Dock = DockStyle.Fill;
        gridRecords.MultiSelect = false;
        gridRecords.ReadOnly = true;
        gridRecords.RowHeadersVisible = false;
        gridRecords.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        gridRecords.Visible = false;

        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(1100, 620);
        Controls.Add(panelDropZone);
        Controls.Add(gridRecords);
        Controls.Add(panelToolbar);
        MinimumSize = new Size(800, 450);
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Εισαγωγή Excel — Preview / Αποστολή συντομεύσεων";

        panelToolbar.ResumeLayout(false);
        panelToolbar.PerformLayout();
        panelDropZone.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)gridRecords).EndInit();
        ResumeLayout(false);
    }
}
