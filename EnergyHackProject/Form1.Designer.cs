namespace EnergyHackProject
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.button1_LEP_ImportFile = new System.Windows.Forms.Button();
            this.CorrectCellNameButton = new System.Windows.Forms.Button();
            this.textBoxSearch = new System.Windows.Forms.TextBox();
            this.labelSearchResult = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.labelSqlState = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.label_Info_DB = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.файлToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.импортToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.лЭПToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.пСToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.экспортToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.бДToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.подключитьсяToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.справкаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.button4_LEP_ExportBD = new System.Windows.Forms.Button();
            this.button5_LEP_ImportBD = new System.Windows.Forms.Button();
            this.button3__LEP_ExportFile = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.button7_PS_ExportBD = new System.Windows.Forms.Button();
            this.button6_PS_ImportBD = new System.Windows.Forms.Button();
            this.button8_PS_ExportFile = new System.Windows.Forms.Button();
            this.button9_PS_ImportFile = new System.Windows.Forms.Button();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.поToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tableLayoutPanel5 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel6 = new System.Windows.Forms.TableLayoutPanel();
            this.groupBox1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.tableLayoutPanel2.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.tableLayoutPanel4.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tableLayoutPanel5.SuspendLayout();
            this.tableLayoutPanel6.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1_LEP_ImportFile
            // 
            this.button1_LEP_ImportFile.Location = new System.Drawing.Point(4, 5);
            this.button1_LEP_ImportFile.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button1_LEP_ImportFile.Name = "button1_LEP_ImportFile";
            this.button1_LEP_ImportFile.Size = new System.Drawing.Size(125, 57);
            this.button1_LEP_ImportFile.TabIndex = 0;
            this.button1_LEP_ImportFile.Text = "Импорт из файла";
            this.button1_LEP_ImportFile.UseVisualStyleBackColor = true;
            this.button1_LEP_ImportFile.Click += new System.EventHandler(this.button1_LEP_ImportFile_Click);
            // 
            // CorrectCellNameButton
            // 
            this.CorrectCellNameButton.Location = new System.Drawing.Point(4, 138);
            this.CorrectCellNameButton.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.CorrectCellNameButton.Name = "CorrectCellNameButton";
            this.CorrectCellNameButton.Size = new System.Drawing.Size(191, 35);
            this.CorrectCellNameButton.TabIndex = 3;
            this.CorrectCellNameButton.Text = "Reserved";
            this.CorrectCellNameButton.UseVisualStyleBackColor = true;
            this.CorrectCellNameButton.Click += new System.EventHandler(this.CorrectCellNameButton_Click);
            // 
            // textBoxSearch
            // 
            this.textBoxSearch.Location = new System.Drawing.Point(4, 78);
            this.textBoxSearch.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxSearch.Name = "textBoxSearch";
            this.textBoxSearch.Size = new System.Drawing.Size(280, 27);
            this.textBoxSearch.TabIndex = 4;
            this.textBoxSearch.TextChanged += new System.EventHandler(this.textBoxSearch_TextChanged);
            // 
            // labelSearchResult
            // 
            this.labelSearchResult.AutoSize = true;
            this.labelSearchResult.Location = new System.Drawing.Point(0, 114);
            this.labelSearchResult.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelSearchResult.Name = "labelSearchResult";
            this.labelSearchResult.Size = new System.Drawing.Size(105, 20);
            this.labelSearchResult.TabIndex = 5;
            this.labelSearchResult.Text = "Type anything!";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.labelSqlState);
            this.groupBox1.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.groupBox1.Location = new System.Drawing.Point(1121, 5);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Size = new System.Drawing.Size(453, 35);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Подключение к базе данных:";
            // 
            // labelSqlState
            // 
            this.labelSqlState.AutoSize = true;
            this.labelSqlState.Location = new System.Drawing.Point(253, 0);
            this.labelSqlState.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelSqlState.Name = "labelSqlState";
            this.labelSqlState.Size = new System.Drawing.Size(96, 24);
            this.labelSqlState.TabIndex = 3;
            this.labelSqlState.Text = "Отключено";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(25, 5);
            this.button2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(169, 65);
            this.button2.TabIndex = 7;
            this.button2.Text = "CreateWord";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button3__LEP_ExportFile_Click_1);
            // 
            // label_Info_DB
            // 
            this.label_Info_DB.AutoSize = true;
            this.label_Info_DB.Location = new System.Drawing.Point(1048, 792);
            this.label_Info_DB.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label_Info_DB.Name = "label_Info_DB";
            this.label_Info_DB.Size = new System.Drawing.Size(0, 20);
            this.label_Info_DB.TabIndex = 8;
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.файлToolStripMenuItem,
            this.бДToolStripMenuItem,
            this.справкаToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(8, 3, 0, 3);
            this.menuStrip1.Size = new System.Drawing.Size(1685, 30);
            this.menuStrip1.TabIndex = 9;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // файлToolStripMenuItem
            // 
            this.файлToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.импортToolStripMenuItem,
            this.экспортToolStripMenuItem});
            this.файлToolStripMenuItem.Name = "файлToolStripMenuItem";
            this.файлToolStripMenuItem.Size = new System.Drawing.Size(59, 24);
            this.файлToolStripMenuItem.Text = "Файл";
            // 
            // импортToolStripMenuItem
            // 
            this.импортToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.лЭПToolStripMenuItem,
            this.пСToolStripMenuItem});
            this.импортToolStripMenuItem.Name = "импортToolStripMenuItem";
            this.импортToolStripMenuItem.Size = new System.Drawing.Size(148, 26);
            this.импортToolStripMenuItem.Text = "Импорт";
            // 
            // лЭПToolStripMenuItem
            // 
            this.лЭПToolStripMenuItem.Name = "лЭПToolStripMenuItem";
            this.лЭПToolStripMenuItem.Size = new System.Drawing.Size(122, 26);
            this.лЭПToolStripMenuItem.Text = "ЛЭП";
            this.лЭПToolStripMenuItem.Click += new System.EventHandler(this.button1_LEP_ImportFile_Click);
            // 
            // пСToolStripMenuItem
            // 
            this.пСToolStripMenuItem.Name = "пСToolStripMenuItem";
            this.пСToolStripMenuItem.Size = new System.Drawing.Size(122, 26);
            this.пСToolStripMenuItem.Text = "ПС";
            this.пСToolStripMenuItem.Click += new System.EventHandler(this.button9_PS_ImportFile_Click);
            // 
            // экспортToolStripMenuItem
            // 
            this.экспортToolStripMenuItem.Name = "экспортToolStripMenuItem";
            this.экспортToolStripMenuItem.Size = new System.Drawing.Size(148, 26);
            this.экспортToolStripMenuItem.Text = "Экспорт";
            // 
            // бДToolStripMenuItem
            // 
            this.бДToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.подключитьсяToolStripMenuItem});
            this.бДToolStripMenuItem.Name = "бДToolStripMenuItem";
            this.бДToolStripMenuItem.Size = new System.Drawing.Size(42, 24);
            this.бДToolStripMenuItem.Text = "БД";
            // 
            // подключитьсяToolStripMenuItem
            // 
            this.подключитьсяToolStripMenuItem.Name = "подключитьсяToolStripMenuItem";
            this.подключитьсяToolStripMenuItem.Size = new System.Drawing.Size(193, 26);
            this.подключитьсяToolStripMenuItem.Text = "Подключиться";
            this.подключитьсяToolStripMenuItem.Click += new System.EventHandler(this.подключитьсяToolStripMenuItem_Click);
            // 
            // справкаToolStripMenuItem
            // 
            this.справкаToolStripMenuItem.Name = "справкаToolStripMenuItem";
            this.справкаToolStripMenuItem.Size = new System.Drawing.Size(81, 24);
            this.справкаToolStripMenuItem.Text = "Справка";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(4, 5);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1677, 846);
            this.tabControl1.TabIndex = 10;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.tableLayoutPanel1);
            this.tabPage1.Location = new System.Drawing.Point(4, 29);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage1.Size = new System.Drawing.Size(1669, 813);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "ЛЭП";
            this.tabPage1.UseVisualStyleBackColor = true;
            this.tabPage1.Click += new System.EventHandler(this.tabPage1_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.dataGridView1, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(4, 5);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 77F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1661, 803);
            this.tableLayoutPanel1.TabIndex = 12;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(4, 82);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.Size = new System.Drawing.Size(1653, 716);
            this.dataGridView1.TabIndex = 1;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 5;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 133F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 133F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 133F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 133F));
            this.tableLayoutPanel2.Controls.Add(this.button4_LEP_ExportBD, 4, 0);
            this.tableLayoutPanel2.Controls.Add(this.button5_LEP_ImportBD, 1, 0);
            this.tableLayoutPanel2.Controls.Add(this.button3__LEP_ExportFile, 3, 0);
            this.tableLayoutPanel2.Controls.Add(this.button1_LEP_ImportFile, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(4, 5);
            this.tableLayoutPanel2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(1653, 67);
            this.tableLayoutPanel2.TabIndex = 0;
            // 
            // button4_LEP_ExportBD
            // 
            this.button4_LEP_ExportBD.Enabled = false;
            this.button4_LEP_ExportBD.Location = new System.Drawing.Point(1524, 5);
            this.button4_LEP_ExportBD.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button4_LEP_ExportBD.Name = "button4_LEP_ExportBD";
            this.button4_LEP_ExportBD.Size = new System.Drawing.Size(125, 57);
            this.button4_LEP_ExportBD.TabIndex = 3;
            this.button4_LEP_ExportBD.Text = "Сохранить в базу данных";
            this.button4_LEP_ExportBD.UseVisualStyleBackColor = true;
            this.button4_LEP_ExportBD.Click += new System.EventHandler(this.button4_LEP_ExportBD_Click);
            // 
            // button5_LEP_ImportBD
            // 
            this.button5_LEP_ImportBD.Enabled = false;
            this.button5_LEP_ImportBD.Location = new System.Drawing.Point(137, 5);
            this.button5_LEP_ImportBD.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button5_LEP_ImportBD.Name = "button5_LEP_ImportBD";
            this.button5_LEP_ImportBD.Size = new System.Drawing.Size(125, 57);
            this.button5_LEP_ImportBD.TabIndex = 4;
            this.button5_LEP_ImportBD.Text = "Импорт из базы данных";
            this.button5_LEP_ImportBD.UseVisualStyleBackColor = true;
            this.button5_LEP_ImportBD.Click += new System.EventHandler(this.button5_LEP_ImportBD_Click);
            // 
            // button3__LEP_ExportFile
            // 
            this.button3__LEP_ExportFile.Location = new System.Drawing.Point(1391, 5);
            this.button3__LEP_ExportFile.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button3__LEP_ExportFile.Name = "button3__LEP_ExportFile";
            this.button3__LEP_ExportFile.Size = new System.Drawing.Size(125, 57);
            this.button3__LEP_ExportFile.TabIndex = 2;
            this.button3__LEP_ExportFile.Text = "Сохранить в файл";
            this.button3__LEP_ExportFile.UseVisualStyleBackColor = true;
            this.button3__LEP_ExportFile.Click += new System.EventHandler(this.button3_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.tableLayoutPanel3);
            this.tabPage2.Location = new System.Drawing.Point(4, 29);
            this.tabPage2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabPage2.Size = new System.Drawing.Size(1669, 813);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "ПС";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 1;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.Controls.Add(this.dataGridView2, 0, 1);
            this.tableLayoutPanel3.Controls.Add(this.tableLayoutPanel4, 0, 0);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel3.Location = new System.Drawing.Point(4, 5);
            this.tableLayoutPanel3.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 2;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 77F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(1661, 803);
            this.tableLayoutPanel3.TabIndex = 13;
            // 
            // dataGridView2
            // 
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView2.Location = new System.Drawing.Point(4, 82);
            this.dataGridView2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowHeadersWidth = 51;
            this.dataGridView2.Size = new System.Drawing.Size(1653, 716);
            this.dataGridView2.TabIndex = 6;
            // 
            // tableLayoutPanel4
            // 
            this.tableLayoutPanel4.ColumnCount = 5;
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 133F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 133F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 133F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 133F));
            this.tableLayoutPanel4.Controls.Add(this.button7_PS_ExportBD, 4, 0);
            this.tableLayoutPanel4.Controls.Add(this.button6_PS_ImportBD, 1, 0);
            this.tableLayoutPanel4.Controls.Add(this.button8_PS_ExportFile, 3, 0);
            this.tableLayoutPanel4.Controls.Add(this.button9_PS_ImportFile, 0, 0);
            this.tableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel4.Location = new System.Drawing.Point(4, 5);
            this.tableLayoutPanel4.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            this.tableLayoutPanel4.RowCount = 1;
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel4.Size = new System.Drawing.Size(1653, 67);
            this.tableLayoutPanel4.TabIndex = 0;
            // 
            // button7_PS_ExportBD
            // 
            this.button7_PS_ExportBD.Enabled = false;
            this.button7_PS_ExportBD.Location = new System.Drawing.Point(1524, 5);
            this.button7_PS_ExportBD.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button7_PS_ExportBD.Name = "button7_PS_ExportBD";
            this.button7_PS_ExportBD.Size = new System.Drawing.Size(125, 57);
            this.button7_PS_ExportBD.TabIndex = 8;
            this.button7_PS_ExportBD.Text = "Сохранить в базу данных";
            this.button7_PS_ExportBD.UseVisualStyleBackColor = true;
            this.button7_PS_ExportBD.Click += new System.EventHandler(this.button7_PS_ExportBD_Click);
            // 
            // button6_PS_ImportBD
            // 
            this.button6_PS_ImportBD.Enabled = false;
            this.button6_PS_ImportBD.Location = new System.Drawing.Point(137, 5);
            this.button6_PS_ImportBD.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button6_PS_ImportBD.Name = "button6_PS_ImportBD";
            this.button6_PS_ImportBD.Size = new System.Drawing.Size(125, 57);
            this.button6_PS_ImportBD.TabIndex = 9;
            this.button6_PS_ImportBD.Text = "Импорт из базы данных";
            this.button6_PS_ImportBD.UseVisualStyleBackColor = true;
            this.button6_PS_ImportBD.Click += new System.EventHandler(this.button6_PS_ImportBD_Click);
            // 
            // button8_PS_ExportFile
            // 
            this.button8_PS_ExportFile.Location = new System.Drawing.Point(1391, 5);
            this.button8_PS_ExportFile.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button8_PS_ExportFile.Name = "button8_PS_ExportFile";
            this.button8_PS_ExportFile.Size = new System.Drawing.Size(125, 57);
            this.button8_PS_ExportFile.TabIndex = 7;
            this.button8_PS_ExportFile.Text = "Сохранить в файл";
            this.button8_PS_ExportFile.UseVisualStyleBackColor = true;
            this.button8_PS_ExportFile.Click += new System.EventHandler(this.button8_PS_ExportFile_Click);
            // 
            // button9_PS_ImportFile
            // 
            this.button9_PS_ImportFile.Location = new System.Drawing.Point(4, 5);
            this.button9_PS_ImportFile.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button9_PS_ImportFile.Name = "button9_PS_ImportFile";
            this.button9_PS_ImportFile.Size = new System.Drawing.Size(125, 57);
            this.button9_PS_ImportFile.TabIndex = 5;
            this.button9_PS_ImportFile.Text = "Импорт из файла";
            this.button9_PS_ImportFile.UseVisualStyleBackColor = true;
            this.button9_PS_ImportFile.Click += new System.EventHandler(this.button9_PS_ImportFile_Click);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.поToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(169, 28);
            this.contextMenuStrip1.Opening += new System.ComponentModel.CancelEventHandler(this.contextMenuStrip1_Opening);
            // 
            // поToolStripMenuItem
            // 
            this.поToolStripMenuItem.Name = "поToolStripMenuItem";
            this.поToolStripMenuItem.Size = new System.Drawing.Size(168, 24);
            this.поToolStripMenuItem.Text = "Примечания";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.CorrectCellNameButton);
            this.panel1.Controls.Add(this.textBoxSearch);
            this.panel1.Controls.Add(this.labelSearchResult);
            this.panel1.Location = new System.Drawing.Point(1654, 5);
            this.panel1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(19, 42);
            this.panel1.TabIndex = 11;
            // 
            // tableLayoutPanel5
            // 
            this.tableLayoutPanel5.ColumnCount = 1;
            this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel5.Controls.Add(this.tabControl1, 0, 0);
            this.tableLayoutPanel5.Controls.Add(this.tableLayoutPanel6, 0, 1);
            this.tableLayoutPanel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel5.Location = new System.Drawing.Point(0, 30);
            this.tableLayoutPanel5.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tableLayoutPanel5.Name = "tableLayoutPanel5";
            this.tableLayoutPanel5.RowCount = 2;
            this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 62F));
            this.tableLayoutPanel5.Size = new System.Drawing.Size(1685, 918);
            this.tableLayoutPanel5.TabIndex = 12;
            // 
            // tableLayoutPanel6
            // 
            this.tableLayoutPanel6.ColumnCount = 3;
            this.tableLayoutPanel6.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel6.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 533F));
            this.tableLayoutPanel6.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 27F));
            this.tableLayoutPanel6.Controls.Add(this.groupBox1, 1, 0);
            this.tableLayoutPanel6.Controls.Add(this.panel1, 2, 0);
            this.tableLayoutPanel6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel6.Location = new System.Drawing.Point(4, 861);
            this.tableLayoutPanel6.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tableLayoutPanel6.Name = "tableLayoutPanel6";
            this.tableLayoutPanel6.RowCount = 1;
            this.tableLayoutPanel6.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel6.Size = new System.Drawing.Size(1677, 52);
            this.tableLayoutPanel6.TabIndex = 11;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1685, 948);
            this.Controls.Add(this.tableLayoutPanel5);
            this.Controls.Add(this.label_Info_DB);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.ShowIcon = false;
            this.Text = "Energy Hack 3";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.tableLayoutPanel4.ResumeLayout(false);
            this.contextMenuStrip1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tableLayoutPanel5.ResumeLayout(false);
            this.tableLayoutPanel6.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1_LEP_ImportFile;
        private System.Windows.Forms.Button CorrectCellNameButton;
        private System.Windows.Forms.TextBox textBoxSearch;
        private System.Windows.Forms.Label labelSearchResult;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label labelSqlState;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label_Info_DB;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem файлToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem импортToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem лЭПToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem пСToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem экспортToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem бДToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem справкаToolStripMenuItem;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Button button3__LEP_ExportFile;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button button4_LEP_ExportBD;
        private System.Windows.Forms.Button button5_LEP_ImportBD;
        private System.Windows.Forms.Button button6_PS_ImportBD;
        private System.Windows.Forms.Button button7_PS_ExportBD;
        private System.Windows.Forms.Button button8_PS_ExportFile;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Button button9_PS_ImportFile;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem поToolStripMenuItem;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.ToolStripMenuItem подключитьсяToolStripMenuItem;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel4;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel5;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel6;
    }
}

