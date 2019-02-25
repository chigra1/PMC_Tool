namespace PCM_Tool
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btnCalculate = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.dGVmaterials = new System.Windows.Forms.DataGridView();
            this.itemNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.itemName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Description = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Supplier = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Delivery_time = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Product = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bomQty = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.safetyStock = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.stockQty = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.WIP = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OP_PO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Assembly_Quantity = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.diff = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MOQ = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.POQ = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Final_demand = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Unit_Price = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Total_price = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dGVboms = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnExit = new System.Windows.Forms.Button();
            this.btnExportAll = new System.Windows.Forms.Button();
            this.dGVerrors = new System.Windows.Forms.DataGridView();
            this.Item = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Detail = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnExportDemand = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.btnExportErrors = new System.Windows.Forms.Button();
            this.lbFoundBoms = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label2 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.informationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dGVmaterials)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dGVboms)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dGVerrors)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnCalculate
            // 
            this.btnCalculate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCalculate.Location = new System.Drawing.Point(399, 207);
            this.btnCalculate.Name = "btnCalculate";
            this.btnCalculate.Size = new System.Drawing.Size(118, 64);
            this.btnCalculate.TabIndex = 0;
            this.btnCalculate.Text = "Calculate material --->";
            this.btnCalculate.UseVisualStyleBackColor = true;
            this.btnCalculate.Click += new System.EventHandler(this.btnCalculate_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(8, 1);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(103, 20);
            this.label3.TabIndex = 7;
            this.label3.Text = "Calculation:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(44, 38);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(261, 20);
            this.label4.TabIndex = 11;
            this.label4.Text = "BOMs from Masterschedule.xls:";
            // 
            // dGVmaterials
            // 
            this.dGVmaterials.AllowUserToAddRows = false;
            this.dGVmaterials.AllowUserToDeleteRows = false;
            this.dGVmaterials.CausesValidation = false;
            this.dGVmaterials.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dGVmaterials.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.itemNumber,
            this.itemName,
            this.Description,
            this.Supplier,
            this.Delivery_time,
            this.Product,
            this.bomQty,
            this.safetyStock,
            this.stockQty,
            this.WIP,
            this.PO,
            this.OP_PO,
            this.Assembly_Quantity,
            this.diff,
            this.MOQ,
            this.POQ,
            this.Final_demand,
            this.Unit_Price,
            this.Total_price});
            this.dGVmaterials.Location = new System.Drawing.Point(12, 23);
            this.dGVmaterials.Name = "dGVmaterials";
            this.dGVmaterials.ReadOnly = true;
            this.dGVmaterials.Size = new System.Drawing.Size(743, 340);
            this.dGVmaterials.TabIndex = 16;
            // 
            // itemNumber
            // 
            this.itemNumber.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.itemNumber.HeaderText = "Item";
            this.itemNumber.Name = "itemNumber";
            this.itemNumber.ReadOnly = true;
            this.itemNumber.Width = 52;
            // 
            // itemName
            // 
            this.itemName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.itemName.HeaderText = "Material Number";
            this.itemName.Name = "itemName";
            this.itemName.ReadOnly = true;
            // 
            // Description
            // 
            this.Description.HeaderText = "Description";
            this.Description.Name = "Description";
            this.Description.ReadOnly = true;
            // 
            // Supplier
            // 
            this.Supplier.HeaderText = "Supplier";
            this.Supplier.Name = "Supplier";
            this.Supplier.ReadOnly = true;
            // 
            // Delivery_time
            // 
            this.Delivery_time.HeaderText = "Delivery Time";
            this.Delivery_time.Name = "Delivery_time";
            this.Delivery_time.ReadOnly = true;
            // 
            // Product
            // 
            this.Product.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Product.HeaderText = "BOMs";
            this.Product.Name = "Product";
            this.Product.ReadOnly = true;
            this.Product.Width = 61;
            // 
            // bomQty
            // 
            this.bomQty.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.bomQty.HeaderText = "Total Required Components";
            this.bomQty.Name = "bomQty";
            this.bomQty.ReadOnly = true;
            this.bomQty.Width = 150;
            // 
            // safetyStock
            // 
            this.safetyStock.HeaderText = "Safety Stock";
            this.safetyStock.Name = "safetyStock";
            this.safetyStock.ReadOnly = true;
            // 
            // stockQty
            // 
            this.stockQty.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.stockQty.HeaderText = "On Stock";
            this.stockQty.Name = "stockQty";
            this.stockQty.ReadOnly = true;
            this.stockQty.Width = 71;
            // 
            // WIP
            // 
            this.WIP.HeaderText = "WIP";
            this.WIP.Name = "WIP";
            this.WIP.ReadOnly = true;
            // 
            // PO
            // 
            this.PO.HeaderText = "PO";
            this.PO.Name = "PO";
            this.PO.ReadOnly = true;
            // 
            // OP_PO
            // 
            this.OP_PO.HeaderText = "OP_PO";
            this.OP_PO.Name = "OP_PO";
            this.OP_PO.ReadOnly = true;
            // 
            // Assembly_Quantity
            // 
            this.Assembly_Quantity.HeaderText = "Assembly_Quantity";
            this.Assembly_Quantity.Name = "Assembly_Quantity";
            this.Assembly_Quantity.ReadOnly = true;
            // 
            // diff
            // 
            this.diff.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.diff.HeaderText = "Difference";
            this.diff.Name = "diff";
            this.diff.ReadOnly = true;
            this.diff.Width = 81;
            // 
            // MOQ
            // 
            this.MOQ.HeaderText = "MOQ";
            this.MOQ.Name = "MOQ";
            this.MOQ.ReadOnly = true;
            // 
            // POQ
            // 
            this.POQ.HeaderText = "POQ";
            this.POQ.Name = "POQ";
            this.POQ.ReadOnly = true;
            // 
            // Final_demand
            // 
            this.Final_demand.HeaderText = "Final Demand";
            this.Final_demand.Name = "Final_demand";
            this.Final_demand.ReadOnly = true;
            // 
            // Unit_Price
            // 
            this.Unit_Price.HeaderText = "Unit Price";
            this.Unit_Price.Name = "Unit_Price";
            this.Unit_Price.ReadOnly = true;
            // 
            // Total_price
            // 
            this.Total_price.HeaderText = "Total Price";
            this.Total_price.Name = "Total_price";
            this.Total_price.ReadOnly = true;
            // 
            // dGVboms
            // 
            this.dGVboms.AllowUserToAddRows = false;
            this.dGVboms.AllowUserToDeleteRows = false;
            this.dGVboms.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dGVboms.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3});
            this.dGVboms.Location = new System.Drawing.Point(30, 60);
            this.dGVboms.Name = "dGVboms";
            this.dGVboms.ReadOnly = true;
            this.dGVboms.Size = new System.Drawing.Size(348, 376);
            this.dGVboms.TabIndex = 18;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.dataGridViewTextBoxColumn1.HeaderText = "Item";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            this.dataGridViewTextBoxColumn1.Width = 50;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.HeaderText = "BOM";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.ReadOnly = true;
            this.dataGridViewTextBoxColumn2.Width = 120;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.HeaderText = "Quantity";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.ReadOnly = true;
            // 
            // btnExit
            // 
            this.btnExit.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExit.Location = new System.Drawing.Point(1152, 718);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(140, 39);
            this.btnExit.TabIndex = 19;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnExportAll
            // 
            this.btnExportAll.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnExportAll.BackgroundImage")));
            this.btnExportAll.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnExportAll.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnExportAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportAll.Location = new System.Drawing.Point(313, 366);
            this.btnExportAll.Name = "btnExportAll";
            this.btnExportAll.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.btnExportAll.Size = new System.Drawing.Size(189, 46);
            this.btnExportAll.TabIndex = 20;
            this.btnExportAll.Text = "Export whole list to .xlsx";
            this.btnExportAll.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnExportAll.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnExportAll.UseVisualStyleBackColor = true;
            this.btnExportAll.Click += new System.EventHandler(this.btnExportAll_Click);
            // 
            // dGVerrors
            // 
            this.dGVerrors.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dGVerrors.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Item,
            this.Detail});
            this.dGVerrors.Location = new System.Drawing.Point(19, 45);
            this.dGVerrors.Name = "dGVerrors";
            this.dGVerrors.Size = new System.Drawing.Size(1248, 157);
            this.dGVerrors.TabIndex = 21;
            // 
            // Item
            // 
            this.Item.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Item.HeaderText = "Id";
            this.Item.Name = "Item";
            this.Item.Width = 41;
            // 
            // Detail
            // 
            this.Detail.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Detail.HeaderText = "Detail";
            this.Detail.Name = "Detail";
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.btnExportDemand);
            this.panel1.Controls.Add(this.dGVmaterials);
            this.panel1.Controls.Add(this.btnExportAll);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Location = new System.Drawing.Point(523, 36);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(769, 417);
            this.panel1.TabIndex = 22;
            // 
            // btnExportDemand
            // 
            this.btnExportDemand.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnExportDemand.BackgroundImage")));
            this.btnExportDemand.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnExportDemand.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnExportDemand.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportDemand.Location = new System.Drawing.Point(508, 366);
            this.btnExportDemand.Name = "btnExportDemand";
            this.btnExportDemand.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.btnExportDemand.Size = new System.Drawing.Size(247, 46);
            this.btnExportDemand.TabIndex = 21;
            this.btnExportDemand.Text = "Export materials on demand to .xlsx";
            this.btnExportDemand.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnExportDemand.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnExportDemand.UseVisualStyleBackColor = true;
            this.btnExportDemand.Click += new System.EventHandler(this.btnExportDemand_Click);
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Location = new System.Drawing.Point(12, 36);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(381, 417);
            this.panel2.TabIndex = 23;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(16, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(200, 20);
            this.label1.TabIndex = 24;
            this.label1.Text = "Errors during execution:";
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.btnExportErrors);
            this.panel3.Controls.Add(this.label1);
            this.panel3.Controls.Add(this.dGVerrors);
            this.panel3.Location = new System.Drawing.Point(12, 494);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1280, 217);
            this.panel3.TabIndex = 25;
            // 
            // btnExportErrors
            // 
            this.btnExportErrors.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnExportErrors.BackgroundImage")));
            this.btnExportErrors.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnExportErrors.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnExportErrors.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportErrors.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExportErrors.Location = new System.Drawing.Point(222, 3);
            this.btnExportErrors.Name = "btnExportErrors";
            this.btnExportErrors.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.btnExportErrors.Size = new System.Drawing.Size(162, 41);
            this.btnExportErrors.TabIndex = 22;
            this.btnExportErrors.Text = "Export Errors to .xlsx";
            this.btnExportErrors.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnExportErrors.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnExportErrors.UseVisualStyleBackColor = true;
            this.btnExportErrors.Click += new System.EventHandler(this.btnExportErrors_Click);
            // 
            // lbFoundBoms
            // 
            this.lbFoundBoms.AutoSize = true;
            this.lbFoundBoms.Location = new System.Drawing.Point(12, 456);
            this.lbFoundBoms.Name = "lbFoundBoms";
            this.lbFoundBoms.Size = new System.Drawing.Size(75, 13);
            this.lbFoundBoms.TabIndex = 26;
            this.lbFoundBoms.Text = "Found BOMs: ";
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(847, 462);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(432, 23);
            this.progressBar1.TabIndex = 27;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(751, 462);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(90, 16);
            this.label2.TabIndex = 28;
            this.label2.Text = "Progress Bar:";
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.settingsToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1304, 24);
            this.menuStrip1.TabIndex = 29;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(92, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(61, 20);
            this.settingsToolStripMenuItem.Text = "Settings";
            this.settingsToolStripMenuItem.Click += new System.EventHandler(this.settingsToolStripMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.informationToolStripMenuItem});
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.aboutToolStripMenuItem.Text = "Help";
            // 
            // informationToolStripMenuItem
            // 
            this.informationToolStripMenuItem.Name = "informationToolStripMenuItem";
            this.informationToolStripMenuItem.Size = new System.Drawing.Size(107, 22);
            this.informationToolStripMenuItem.Text = "About";
            this.informationToolStripMenuItem.Click += new System.EventHandler(this.informationToolStripMenuItem_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(105, 726);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(0, 13);
            this.label5.TabIndex = 30;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(677, 726);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(0, 13);
            this.label6.TabIndex = 31;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1304, 779);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.lbFoundBoms);
            this.Controls.Add(this.dGVboms);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnCalculate);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.menuStrip1);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PMC Tool v.0.2.0";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dGVmaterials)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dGVboms)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dGVerrors)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCalculate;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DataGridView dGVmaterials;
        private System.Windows.Forms.DataGridView dGVboms;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Button btnExportAll;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridView dGVerrors;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label lbFoundBoms;
        private System.Windows.Forms.Button btnExportDemand;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem informationToolStripMenuItem;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnExportErrors;
        private System.Windows.Forms.DataGridViewTextBoxColumn Item;
        private System.Windows.Forms.DataGridViewTextBoxColumn Detail;
        private System.Windows.Forms.DataGridViewTextBoxColumn itemNumber;
        private System.Windows.Forms.DataGridViewTextBoxColumn itemName;
        private System.Windows.Forms.DataGridViewTextBoxColumn Description;
        private System.Windows.Forms.DataGridViewTextBoxColumn Supplier;
        private System.Windows.Forms.DataGridViewTextBoxColumn Delivery_time;
        private System.Windows.Forms.DataGridViewTextBoxColumn Product;
        private System.Windows.Forms.DataGridViewTextBoxColumn bomQty;
        private System.Windows.Forms.DataGridViewTextBoxColumn safetyStock;
        private System.Windows.Forms.DataGridViewTextBoxColumn stockQty;
        private System.Windows.Forms.DataGridViewTextBoxColumn WIP;
        private System.Windows.Forms.DataGridViewTextBoxColumn PO;
        private System.Windows.Forms.DataGridViewTextBoxColumn OP_PO;
        private System.Windows.Forms.DataGridViewTextBoxColumn Assembly_Quantity;
        private System.Windows.Forms.DataGridViewTextBoxColumn diff;
        private System.Windows.Forms.DataGridViewTextBoxColumn MOQ;
        private System.Windows.Forms.DataGridViewTextBoxColumn POQ;
        private System.Windows.Forms.DataGridViewTextBoxColumn Final_demand;
        private System.Windows.Forms.DataGridViewTextBoxColumn Unit_Price;
        private System.Windows.Forms.DataGridViewTextBoxColumn Total_price;
        private System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;
    }
}

