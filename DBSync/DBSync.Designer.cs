namespace DBSync
{
    partial class DBSync
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DBSync));
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.tbMainBaseName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cbShop = new System.Windows.Forms.ComboBox();
            this.gbMain = new System.Windows.Forms.GroupBox();
            this.gbShop = new System.Windows.Forms.GroupBox();
            this.tbLastDateSync = new System.Windows.Forms.TextBox();
            this.tbShopBaseName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.gbSync = new System.Windows.Forms.GroupBox();
            this.label27 = new System.Windows.Forms.Label();
            this.label28 = new System.Windows.Forms.Label();
            this.tbInvoiceMoveOutChange2 = new System.Windows.Forms.TextBox();
            this.label29 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.tbInvoiceMoveInChange2 = new System.Windows.Forms.TextBox();
            this.label26 = new System.Windows.Forms.Label();
            this.cbNoDeleteDocuments = new System.Windows.Forms.CheckBox();
            this.cbVerifyMainBase = new System.Windows.Forms.CheckBox();
            this.label13 = new System.Windows.Forms.Label();
            this.tbPriceChange2 = new System.Windows.Forms.TextBox();
            this.tbTovarChange2 = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.cbDeleteAutoInvoice = new System.Windows.Forms.CheckBox();
            this.btnFixRemains = new System.Windows.Forms.Button();
            this.label25 = new System.Windows.Forms.Label();
            this.label24 = new System.Windows.Forms.Label();
            this.label23 = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.tbInvoiceReturnChange2 = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.tbInvoiceInChange2 = new System.Windows.Forms.TextBox();
            this.tbInvoiceOutChange2 = new System.Windows.Forms.TextBox();
            this.tbContragentChange2 = new System.Windows.Forms.TextBox();
            this.tbContragentChange = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.tbDiscountChange = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.tbPriceChange = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.tbTovarChange = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.listView1 = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.backgroundWorker2 = new System.ComponentModel.BackgroundWorker();
            this.backgroundWorker3 = new System.ComponentModel.BackgroundWorker();
            this.btnRefreshDataSite = new System.Windows.Forms.Button();
            this.backgroundWorker4 = new System.ComponentModel.BackgroundWorker();
            this.bgVerifyTovar = new System.ComponentModel.BackgroundWorker();
            this.bgVerifyInvoice = new System.ComponentModel.BackgroundWorker();
            this.btnYML = new System.Windows.Forms.Button();
            this.backgroundWorker5 = new System.ComponentModel.BackgroundWorker();
            this.gbMain.SuspendLayout();
            this.gbShop.SuspendLayout();
            this.gbSync.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(395, 16);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(104, 25);
            this.button1.TabIndex = 0;
            this.button1.Text = "Контроль";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(395, 15);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(104, 25);
            this.button2.TabIndex = 1;
            this.button2.Text = "Подключиться";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(87, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Основная база:";
            // 
            // tbMainBaseName
            // 
            this.tbMainBaseName.Location = new System.Drawing.Point(99, 18);
            this.tbMainBaseName.Name = "tbMainBaseName";
            this.tbMainBaseName.Size = new System.Drawing.Size(290, 20);
            this.tbMainBaseName.TabIndex = 3;
            this.tbMainBaseName.Text = "Srvr=\"milch.eimb.ru\";Ref=\"base\";Usr=\"Admin\";Pwd=\"131313\"";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(54, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Магазин:";
            // 
            // cbShop
            // 
            this.cbShop.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbShop.FormattingEnabled = true;
            this.cbShop.Location = new System.Drawing.Point(99, 19);
            this.cbShop.Name = "cbShop";
            this.cbShop.Size = new System.Drawing.Size(290, 21);
            this.cbShop.TabIndex = 5;
            this.cbShop.SelectedIndexChanged += new System.EventHandler(this.cbShop_SelectedIndexChanged);
            // 
            // gbMain
            // 
            this.gbMain.Controls.Add(this.tbMainBaseName);
            this.gbMain.Controls.Add(this.button2);
            this.gbMain.Controls.Add(this.label1);
            this.gbMain.Location = new System.Drawing.Point(12, 7);
            this.gbMain.Name = "gbMain";
            this.gbMain.Size = new System.Drawing.Size(515, 49);
            this.gbMain.TabIndex = 6;
            this.gbMain.TabStop = false;
            // 
            // gbShop
            // 
            this.gbShop.Controls.Add(this.tbLastDateSync);
            this.gbShop.Controls.Add(this.tbShopBaseName);
            this.gbShop.Controls.Add(this.label3);
            this.gbShop.Controls.Add(this.label2);
            this.gbShop.Controls.Add(this.cbShop);
            this.gbShop.Controls.Add(this.button1);
            this.gbShop.Enabled = false;
            this.gbShop.Location = new System.Drawing.Point(12, 63);
            this.gbShop.Name = "gbShop";
            this.gbShop.Size = new System.Drawing.Size(515, 79);
            this.gbShop.TabIndex = 7;
            this.gbShop.TabStop = false;
            // 
            // tbLastDateSync
            // 
            this.tbLastDateSync.Location = new System.Drawing.Point(395, 46);
            this.tbLastDateSync.Name = "tbLastDateSync";
            this.tbLastDateSync.ReadOnly = true;
            this.tbLastDateSync.Size = new System.Drawing.Size(104, 20);
            this.tbLastDateSync.TabIndex = 9;
            // 
            // tbShopBaseName
            // 
            this.tbShopBaseName.Location = new System.Drawing.Point(99, 46);
            this.tbShopBaseName.Name = "tbShopBaseName";
            this.tbShopBaseName.Size = new System.Drawing.Size(290, 20);
            this.tbShopBaseName.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 49);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(87, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "База магазина:";
            // 
            // gbSync
            // 
            this.gbSync.Controls.Add(this.label27);
            this.gbSync.Controls.Add(this.label28);
            this.gbSync.Controls.Add(this.tbInvoiceMoveOutChange2);
            this.gbSync.Controls.Add(this.label29);
            this.gbSync.Controls.Add(this.label14);
            this.gbSync.Controls.Add(this.label20);
            this.gbSync.Controls.Add(this.tbInvoiceMoveInChange2);
            this.gbSync.Controls.Add(this.label26);
            this.gbSync.Controls.Add(this.cbNoDeleteDocuments);
            this.gbSync.Controls.Add(this.cbVerifyMainBase);
            this.gbSync.Controls.Add(this.label13);
            this.gbSync.Controls.Add(this.tbPriceChange2);
            this.gbSync.Controls.Add(this.tbTovarChange2);
            this.gbSync.Controls.Add(this.label19);
            this.gbSync.Controls.Add(this.cbDeleteAutoInvoice);
            this.gbSync.Controls.Add(this.btnFixRemains);
            this.gbSync.Controls.Add(this.label25);
            this.gbSync.Controls.Add(this.label24);
            this.gbSync.Controls.Add(this.label23);
            this.gbSync.Controls.Add(this.label22);
            this.gbSync.Controls.Add(this.label21);
            this.gbSync.Controls.Add(this.label18);
            this.gbSync.Controls.Add(this.label17);
            this.gbSync.Controls.Add(this.label16);
            this.gbSync.Controls.Add(this.label15);
            this.gbSync.Controls.Add(this.tbInvoiceReturnChange2);
            this.gbSync.Controls.Add(this.button3);
            this.gbSync.Controls.Add(this.label9);
            this.gbSync.Controls.Add(this.tbInvoiceInChange2);
            this.gbSync.Controls.Add(this.tbInvoiceOutChange2);
            this.gbSync.Controls.Add(this.tbContragentChange2);
            this.gbSync.Controls.Add(this.tbContragentChange);
            this.gbSync.Controls.Add(this.label12);
            this.gbSync.Controls.Add(this.label11);
            this.gbSync.Controls.Add(this.label10);
            this.gbSync.Controls.Add(this.tbDiscountChange);
            this.gbSync.Controls.Add(this.label8);
            this.gbSync.Controls.Add(this.tbPriceChange);
            this.gbSync.Controls.Add(this.label7);
            this.gbSync.Controls.Add(this.label6);
            this.gbSync.Controls.Add(this.label5);
            this.gbSync.Controls.Add(this.tbTovarChange);
            this.gbSync.Controls.Add(this.label4);
            this.gbSync.Enabled = false;
            this.gbSync.Location = new System.Drawing.Point(12, 148);
            this.gbSync.Name = "gbSync";
            this.gbSync.Size = new System.Drawing.Size(516, 366);
            this.gbSync.TabIndex = 8;
            this.gbSync.TabStop = false;
            this.gbSync.Text = "Данные для синхронизации";
            // 
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label27.Location = new System.Drawing.Point(325, 240);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(21, 13);
            this.label27.TabIndex = 58;
            this.label27.Text = "<=";
            // 
            // label28
            // 
            this.label28.AutoSize = true;
            this.label28.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label28.ForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.label28.Location = new System.Drawing.Point(225, 240);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(33, 13);
            this.label28.TabIndex = 57;
            this.label28.Text = "<...>";
            // 
            // tbInvoiceMoveOutChange2
            // 
            this.tbInvoiceMoveOutChange2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbInvoiceMoveOutChange2.Location = new System.Drawing.Point(353, 237);
            this.tbInvoiceMoveOutChange2.Name = "tbInvoiceMoveOutChange2";
            this.tbInvoiceMoveOutChange2.ReadOnly = true;
            this.tbInvoiceMoveOutChange2.Size = new System.Drawing.Size(146, 20);
            this.tbInvoiceMoveOutChange2.TabIndex = 56;
            this.tbInvoiceMoveOutChange2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label29
            // 
            this.label29.AutoSize = true;
            this.label29.Location = new System.Drawing.Point(6, 240);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(120, 13);
            this.label29.TabIndex = 55;
            this.label29.Text = "Документы списания:";
            this.label29.DoubleClick += new System.EventHandler(this.label6_DoubleClick);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label14.Location = new System.Drawing.Point(325, 216);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(21, 13);
            this.label14.TabIndex = 54;
            this.label14.Text = "<=";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label20.ForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.label20.Location = new System.Drawing.Point(225, 216);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(33, 13);
            this.label20.TabIndex = 53;
            this.label20.Text = "<...>";
            // 
            // tbInvoiceMoveInChange2
            // 
            this.tbInvoiceMoveInChange2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbInvoiceMoveInChange2.Location = new System.Drawing.Point(353, 213);
            this.tbInvoiceMoveInChange2.Name = "tbInvoiceMoveInChange2";
            this.tbInvoiceMoveInChange2.ReadOnly = true;
            this.tbInvoiceMoveInChange2.Size = new System.Drawing.Size(146, 20);
            this.tbInvoiceMoveInChange2.TabIndex = 52;
            this.tbInvoiceMoveInChange2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(6, 216);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(149, 13);
            this.label26.TabIndex = 51;
            this.label26.Text = "Документы оприходования:";
            this.label26.DoubleClick += new System.EventHandler(this.label6_DoubleClick);
            // 
            // cbNoDeleteDocuments
            // 
            this.cbNoDeleteDocuments.AutoSize = true;
            this.cbNoDeleteDocuments.Enabled = false;
            this.cbNoDeleteDocuments.Location = new System.Drawing.Point(63, 338);
            this.cbNoDeleteDocuments.Name = "cbNoDeleteDocuments";
            this.cbNoDeleteDocuments.Size = new System.Drawing.Size(227, 17);
            this.cbNoDeleteDocuments.TabIndex = 50;
            this.cbNoDeleteDocuments.Text = "не удалять документы без флага Отчет";
            this.cbNoDeleteDocuments.UseVisualStyleBackColor = true;
            // 
            // cbVerifyMainBase
            // 
            this.cbVerifyMainBase.AutoSize = true;
            this.cbVerifyMainBase.Enabled = false;
            this.cbVerifyMainBase.Location = new System.Drawing.Point(63, 320);
            this.cbVerifyMainBase.Name = "cbVerifyMainBase";
            this.cbVerifyMainBase.Size = new System.Drawing.Size(160, 17);
            this.cbVerifyMainBase.TabIndex = 49;
            this.cbVerifyMainBase.Text = "сверять с основной базой";
            this.cbVerifyMainBase.UseVisualStyleBackColor = true;
            this.cbVerifyMainBase.CheckedChanged += new System.EventHandler(this.cbVerifyMainBase_CheckedChanged);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label13.Location = new System.Drawing.Point(320, 72);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(28, 13);
            this.label13.TabIndex = 48;
            this.label13.Text = "<=>";
            // 
            // tbPriceChange2
            // 
            this.tbPriceChange2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbPriceChange2.Location = new System.Drawing.Point(353, 69);
            this.tbPriceChange2.Name = "tbPriceChange2";
            this.tbPriceChange2.ReadOnly = true;
            this.tbPriceChange2.Size = new System.Drawing.Size(146, 20);
            this.tbPriceChange2.TabIndex = 47;
            this.tbPriceChange2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // tbTovarChange2
            // 
            this.tbTovarChange2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbTovarChange2.Location = new System.Drawing.Point(353, 45);
            this.tbTovarChange2.Name = "tbTovarChange2";
            this.tbTovarChange2.ReadOnly = true;
            this.tbTovarChange2.Size = new System.Drawing.Size(146, 20);
            this.tbTovarChange2.TabIndex = 46;
            this.tbTovarChange2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label19.Location = new System.Drawing.Point(320, 48);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(28, 13);
            this.label19.TabIndex = 45;
            this.label19.Text = "<=>";
            // 
            // cbDeleteAutoInvoice
            // 
            this.cbDeleteAutoInvoice.AutoSize = true;
            this.cbDeleteAutoInvoice.Enabled = false;
            this.cbDeleteAutoInvoice.Location = new System.Drawing.Point(63, 302);
            this.cbDeleteAutoInvoice.Name = "cbDeleteAutoInvoice";
            this.cbDeleteAutoInvoice.Size = new System.Drawing.Size(287, 17);
            this.cbDeleteAutoInvoice.TabIndex = 44;
            this.cbDeleteAutoInvoice.Text = "удалять документы поступления и инвентаризации";
            this.cbDeleteAutoInvoice.UseVisualStyleBackColor = true;
            this.cbDeleteAutoInvoice.CheckedChanged += new System.EventHandler(this.cbDeleteAutoInvoice_CheckedChanged);
            // 
            // btnFixRemains
            // 
            this.btnFixRemains.Enabled = false;
            this.btnFixRemains.Location = new System.Drawing.Point(353, 302);
            this.btnFixRemains.Name = "btnFixRemains";
            this.btnFixRemains.Size = new System.Drawing.Size(146, 53);
            this.btnFixRemains.TabIndex = 43;
            this.btnFixRemains.Text = "Зафиксировать остатки";
            this.btnFixRemains.UseVisualStyleBackColor = true;
            this.btnFixRemains.Click += new System.EventHandler(this.btnFixRemains_Click);
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label25.Location = new System.Drawing.Point(325, 192);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(21, 13);
            this.label25.TabIndex = 42;
            this.label25.Text = "<=";
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label24.Location = new System.Drawing.Point(325, 168);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(21, 13);
            this.label24.TabIndex = 41;
            this.label24.Text = "<=";
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label23.Location = new System.Drawing.Point(325, 144);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(21, 13);
            this.label23.TabIndex = 40;
            this.label23.Text = "<=";
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label22.Location = new System.Drawing.Point(320, 120);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(28, 13);
            this.label22.TabIndex = 39;
            this.label22.Text = "<=>";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label21.Location = new System.Drawing.Point(325, 96);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(21, 13);
            this.label21.TabIndex = 38;
            this.label21.Text = "=>";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label18.ForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.label18.Location = new System.Drawing.Point(225, 144);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(33, 13);
            this.label18.TabIndex = 35;
            this.label18.Text = "<...>";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label17.ForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.label17.Location = new System.Drawing.Point(225, 192);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(33, 13);
            this.label17.TabIndex = 34;
            this.label17.Text = "<...>";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label16.ForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.label16.Location = new System.Drawing.Point(225, 168);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(33, 13);
            this.label16.TabIndex = 33;
            this.label16.Text = "<...>";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label15.ForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.label15.Location = new System.Drawing.Point(410, 96);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(33, 13);
            this.label15.TabIndex = 32;
            this.label15.Text = "<...>";
            // 
            // tbInvoiceReturnChange2
            // 
            this.tbInvoiceReturnChange2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbInvoiceReturnChange2.Location = new System.Drawing.Point(353, 189);
            this.tbInvoiceReturnChange2.Name = "tbInvoiceReturnChange2";
            this.tbInvoiceReturnChange2.ReadOnly = true;
            this.tbInvoiceReturnChange2.Size = new System.Drawing.Size(146, 20);
            this.tbInvoiceReturnChange2.TabIndex = 29;
            this.tbInvoiceReturnChange2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(353, 263);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(146, 32);
            this.button3.TabIndex = 8;
            this.button3.Text = "Синхронизировать";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(6, 192);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(119, 13);
            this.label9.TabIndex = 27;
            this.label9.Text = "Документы возврата:";
            this.label9.DoubleClick += new System.EventHandler(this.label6_DoubleClick);
            // 
            // tbInvoiceInChange2
            // 
            this.tbInvoiceInChange2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbInvoiceInChange2.Location = new System.Drawing.Point(353, 141);
            this.tbInvoiceInChange2.Name = "tbInvoiceInChange2";
            this.tbInvoiceInChange2.ReadOnly = true;
            this.tbInvoiceInChange2.Size = new System.Drawing.Size(146, 20);
            this.tbInvoiceInChange2.TabIndex = 26;
            this.tbInvoiceInChange2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // tbInvoiceOutChange2
            // 
            this.tbInvoiceOutChange2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbInvoiceOutChange2.Location = new System.Drawing.Point(353, 165);
            this.tbInvoiceOutChange2.Name = "tbInvoiceOutChange2";
            this.tbInvoiceOutChange2.ReadOnly = true;
            this.tbInvoiceOutChange2.Size = new System.Drawing.Size(146, 20);
            this.tbInvoiceOutChange2.TabIndex = 25;
            this.tbInvoiceOutChange2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // tbContragentChange2
            // 
            this.tbContragentChange2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbContragentChange2.Location = new System.Drawing.Point(353, 117);
            this.tbContragentChange2.Name = "tbContragentChange2";
            this.tbContragentChange2.ReadOnly = true;
            this.tbContragentChange2.Size = new System.Drawing.Size(146, 20);
            this.tbContragentChange2.TabIndex = 24;
            this.tbContragentChange2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // tbContragentChange
            // 
            this.tbContragentChange.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbContragentChange.Location = new System.Drawing.Point(169, 117);
            this.tbContragentChange.Name = "tbContragentChange";
            this.tbContragentChange.ReadOnly = true;
            this.tbContragentChange.Size = new System.Drawing.Size(146, 20);
            this.tbContragentChange.TabIndex = 23;
            this.tbContragentChange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(6, 120);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(76, 13);
            this.label12.TabIndex = 22;
            this.label12.Text = "Контрагенты:";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label11.Location = new System.Drawing.Point(374, 20);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(97, 13);
            this.label11.TabIndex = 21;
            this.label11.Text = "База магазина";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label10.Location = new System.Drawing.Point(193, 20);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(97, 13);
            this.label10.TabIndex = 20;
            this.label10.Text = "Основная база";
            // 
            // tbDiscountChange
            // 
            this.tbDiscountChange.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbDiscountChange.Location = new System.Drawing.Point(169, 93);
            this.tbDiscountChange.Name = "tbDiscountChange";
            this.tbDiscountChange.ReadOnly = true;
            this.tbDiscountChange.Size = new System.Drawing.Size(146, 20);
            this.tbDiscountChange.TabIndex = 17;
            this.tbDiscountChange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 96);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(47, 13);
            this.label8.TabIndex = 16;
            this.label8.Text = "Скидки:";
            // 
            // tbPriceChange
            // 
            this.tbPriceChange.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbPriceChange.Location = new System.Drawing.Point(169, 69);
            this.tbPriceChange.Name = "tbPriceChange";
            this.tbPriceChange.ReadOnly = true;
            this.tbPriceChange.Size = new System.Drawing.Size(146, 20);
            this.tbPriceChange.TabIndex = 14;
            this.tbPriceChange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(6, 72);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(145, 13);
            this.label7.TabIndex = 13;
            this.label7.Text = "Документы установки цен:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 144);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(136, 13);
            this.label6.TabIndex = 11;
            this.label6.Text = "Документы поступления:";
            this.label6.DoubleClick += new System.EventHandler(this.label6_DoubleClick);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 168);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(132, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Документы реализации:";
            this.label5.DoubleClick += new System.EventHandler(this.label6_DoubleClick);
            // 
            // tbTovarChange
            // 
            this.tbTovarChange.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbTovarChange.Location = new System.Drawing.Point(169, 45);
            this.tbTovarChange.Name = "tbTovarChange";
            this.tbTovarChange.ReadOnly = true;
            this.tbTovarChange.Size = new System.Drawing.Size(146, 20);
            this.tbTovarChange.TabIndex = 8;
            this.tbTovarChange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(7, 48);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(135, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "Номенклатуры / бренды:";
            this.label4.Click += new System.EventHandler(this.label4_Click);
            this.label4.DoubleClick += new System.EventHandler(this.label4_DoubleClick);
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // listView1
            // 
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
            this.listView1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.listView1.GridLines = true;
            this.listView1.Location = new System.Drawing.Point(12, 520);
            this.listView1.MultiSelect = false;
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(515, 130);
            this.listView1.SmallImageList = this.imageList1;
            this.listView1.TabIndex = 0;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Сообщения";
            this.columnHeader1.Width = 486;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "msg.bmp");
            this.imageList1.Images.SetKeyName(1, "err.bmp");
            this.imageList1.Images.SetKeyName(2, "success.bmp");
            // 
            // backgroundWorker2
            // 
            this.backgroundWorker2.WorkerReportsProgress = true;
            this.backgroundWorker2.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker2_DoWork);
            this.backgroundWorker2.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker2_ProgressChanged);
            this.backgroundWorker2.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker2_RunWorkerCompleted);
            // 
            // backgroundWorker3
            // 
            this.backgroundWorker3.WorkerReportsProgress = true;
            this.backgroundWorker3.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker3_DoWork);
            this.backgroundWorker3.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker3_ProgressChanged);
            this.backgroundWorker3.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker3_RunWorkerCompleted);
            // 
            // btnRefreshDataSite
            // 
            this.btnRefreshDataSite.Enabled = false;
            this.btnRefreshDataSite.Location = new System.Drawing.Point(365, 656);
            this.btnRefreshDataSite.Name = "btnRefreshDataSite";
            this.btnRefreshDataSite.Size = new System.Drawing.Size(162, 30);
            this.btnRefreshDataSite.TabIndex = 4;
            this.btnRefreshDataSite.Text = "Обновить данные на сайте";
            this.btnRefreshDataSite.UseVisualStyleBackColor = true;
            this.btnRefreshDataSite.Click += new System.EventHandler(this.btnRefreshDataSite_Click);
            // 
            // backgroundWorker4
            // 
            this.backgroundWorker4.WorkerReportsProgress = true;
            this.backgroundWorker4.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker4_DoWork);
            this.backgroundWorker4.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker4_ProgressChanged);
            this.backgroundWorker4.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker4_RunWorkerCompleted);
            // 
            // bgVerifyTovar
            // 
            this.bgVerifyTovar.WorkerReportsProgress = true;
            this.bgVerifyTovar.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgVerifyTovar_DoWork);
            this.bgVerifyTovar.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bgVerifyTovar_ProgressChanged);
            this.bgVerifyTovar.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bgVerifyTovar_RunWorkerCompleted);
            // 
            // bgVerifyInvoice
            // 
            this.bgVerifyInvoice.WorkerReportsProgress = true;
            this.bgVerifyInvoice.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bVerifyInvoice_DoWork);
            this.bgVerifyInvoice.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bgVerifyTovar_ProgressChanged);
            this.bgVerifyInvoice.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bgVerifyTovar_RunWorkerCompleted);
            // 
            // btnYML
            // 
            this.btnYML.Enabled = false;
            this.btnYML.Location = new System.Drawing.Point(12, 656);
            this.btnYML.Name = "btnYML";
            this.btnYML.Size = new System.Drawing.Size(162, 30);
            this.btnYML.TabIndex = 9;
            this.btnYML.Text = "Обновить файл для Яндекса";
            this.btnYML.UseVisualStyleBackColor = true;
            this.btnYML.Click += new System.EventHandler(this.btnYML_Click);
            // 
            // backgroundWorker5
            // 
            this.backgroundWorker5.WorkerReportsProgress = true;
            this.backgroundWorker5.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker5_DoWork);
            this.backgroundWorker5.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker5_ProgressChanged);
            this.backgroundWorker5.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker5_RunWorkerCompleted);
            // 
            // DBSync
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(539, 692);
            this.Controls.Add(this.btnYML);
            this.Controls.Add(this.btnRefreshDataSite);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.gbSync);
            this.Controls.Add(this.gbShop);
            this.Controls.Add(this.gbMain);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DBSync";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Синхронизация баз 1С";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DBSync_FormClosing);
            this.Load += new System.EventHandler(this.DBSync_Load);
            this.gbMain.ResumeLayout(false);
            this.gbMain.PerformLayout();
            this.gbShop.ResumeLayout(false);
            this.gbShop.PerformLayout();
            this.gbSync.ResumeLayout(false);
            this.gbSync.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbMainBaseName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbShop;
        private System.Windows.Forms.GroupBox gbMain;
        private System.Windows.Forms.GroupBox gbShop;
        private System.Windows.Forms.GroupBox gbSync;
        private System.Windows.Forms.TextBox tbShopBaseName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tbTovarChange;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tbPriceChange;
        private System.Windows.Forms.Label label7;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.TextBox tbDiscountChange;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox tbInvoiceReturnChange2;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox tbInvoiceInChange2;
        private System.Windows.Forms.TextBox tbInvoiceOutChange2;
        private System.Windows.Forms.TextBox tbContragentChange2;
        private System.Windows.Forms.TextBox tbContragentChange;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Button btnFixRemains;
        private System.ComponentModel.BackgroundWorker backgroundWorker2;
        private System.Windows.Forms.CheckBox cbDeleteAutoInvoice;
        private System.ComponentModel.BackgroundWorker backgroundWorker3;
        private System.Windows.Forms.TextBox tbTovarChange2;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox tbPriceChange2;
        private System.Windows.Forms.CheckBox cbVerifyMainBase;
        private System.Windows.Forms.Button btnRefreshDataSite;
        private System.ComponentModel.BackgroundWorker backgroundWorker4;
        private System.Windows.Forms.CheckBox cbNoDeleteDocuments;
        private System.Windows.Forms.Label label27;
        private System.Windows.Forms.Label label28;
        private System.Windows.Forms.TextBox tbInvoiceMoveOutChange2;
        private System.Windows.Forms.Label label29;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.TextBox tbInvoiceMoveInChange2;
        private System.Windows.Forms.Label label26;
        private System.ComponentModel.BackgroundWorker bgVerifyTovar;
        private System.ComponentModel.BackgroundWorker bgVerifyInvoice;
        private System.Windows.Forms.TextBox tbLastDateSync;
        private System.Windows.Forms.Button btnYML;
        private System.ComponentModel.BackgroundWorker backgroundWorker5;
    }
}

