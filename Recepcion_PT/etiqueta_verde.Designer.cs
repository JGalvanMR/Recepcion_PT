namespace Recepcion_PT
{
    partial class etiqueta_verde
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(etiqueta_verde));
            this.txtclave = new System.Windows.Forms.TextBox();
            this.lbclave = new System.Windows.Forms.Label();
            this.txtrecibo = new System.Windows.Forms.TextBox();
            this.lbrecibo = new System.Windows.Forms.Label();
            this.txtviaje = new System.Windows.Forms.TextBox();
            this.lbviaje = new System.Windows.Forms.Label();
            this.txtprov = new System.Windows.Forms.TextBox();
            this.lbprov = new System.Windows.Forms.Label();
            this.txtrchtbl = new System.Windows.Forms.TextBox();
            this.lbrchtbl = new System.Windows.Forms.Label();
            this.txthora = new System.Windows.Forms.TextBox();
            this.lbhora = new System.Windows.Forms.Label();
            this.DGV4 = new System.Windows.Forms.DataGridView();
            this.prod_clave = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.prod_nombre = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.numero_tarimas = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cajas_x_tarima = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.esp = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.imp = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.fec_cad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.grad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.country = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Items = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnactu = new System.Windows.Forms.Button();
            this.DGV5 = new System.Windows.Forms.DataGridView();
            this.clave_producto = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.descrip = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.box = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tarimas_total = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.especi = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cad_fec = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.imp_tar = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.id_SSCC = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Item = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnvalidareti = new System.Windows.Forms.Button();
            this.btnimpeti = new System.Windows.Forms.Button();
            this.eti_verde = new System.Drawing.Printing.PrintDocument();
            this.lbgrado = new System.Windows.Forms.Label();
            this.cbgrado = new System.Windows.Forms.ComboBox();
            this.LblHrCap = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.DGV4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV5)).BeginInit();
            this.SuspendLayout();
            // 
            // txtclave
            // 
            this.txtclave.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtclave.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtclave.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtclave.Location = new System.Drawing.Point(107, 7);
            this.txtclave.MaxLength = 10;
            this.txtclave.Name = "txtclave";
            this.txtclave.ReadOnly = true;
            this.txtclave.Size = new System.Drawing.Size(391, 27);
            this.txtclave.TabIndex = 439;
            // 
            // lbclave
            // 
            this.lbclave.AutoSize = true;
            this.lbclave.BackColor = System.Drawing.Color.Transparent;
            this.lbclave.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbclave.ForeColor = System.Drawing.Color.White;
            this.lbclave.Location = new System.Drawing.Point(14, 10);
            this.lbclave.Name = "lbclave";
            this.lbclave.Size = new System.Drawing.Size(46, 16);
            this.lbclave.TabIndex = 438;
            this.lbclave.Text = "Clave:";
            // 
            // txtrecibo
            // 
            this.txtrecibo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtrecibo.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtrecibo.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtrecibo.Location = new System.Drawing.Point(107, 59);
            this.txtrecibo.MaxLength = 10;
            this.txtrecibo.Name = "txtrecibo";
            this.txtrecibo.ReadOnly = true;
            this.txtrecibo.Size = new System.Drawing.Size(391, 27);
            this.txtrecibo.TabIndex = 441;
            // 
            // lbrecibo
            // 
            this.lbrecibo.AutoSize = true;
            this.lbrecibo.BackColor = System.Drawing.Color.Transparent;
            this.lbrecibo.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbrecibo.ForeColor = System.Drawing.Color.White;
            this.lbrecibo.Location = new System.Drawing.Point(14, 62);
            this.lbrecibo.Name = "lbrecibo";
            this.lbrecibo.Size = new System.Drawing.Size(54, 16);
            this.lbrecibo.TabIndex = 440;
            this.lbrecibo.Text = "Recibo:";
            // 
            // txtviaje
            // 
            this.txtviaje.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtviaje.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtviaje.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtviaje.Location = new System.Drawing.Point(107, 107);
            this.txtviaje.MaxLength = 10;
            this.txtviaje.Name = "txtviaje";
            this.txtviaje.ReadOnly = true;
            this.txtviaje.Size = new System.Drawing.Size(391, 27);
            this.txtviaje.TabIndex = 443;
            // 
            // lbviaje
            // 
            this.lbviaje.AutoSize = true;
            this.lbviaje.BackColor = System.Drawing.Color.Transparent;
            this.lbviaje.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbviaje.ForeColor = System.Drawing.Color.White;
            this.lbviaje.Location = new System.Drawing.Point(14, 111);
            this.lbviaje.Name = "lbviaje";
            this.lbviaje.Size = new System.Drawing.Size(43, 16);
            this.lbviaje.TabIndex = 442;
            this.lbviaje.Text = "Viaje:";
            // 
            // txtprov
            // 
            this.txtprov.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtprov.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtprov.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtprov.Location = new System.Drawing.Point(722, 7);
            this.txtprov.MaxLength = 10;
            this.txtprov.Name = "txtprov";
            this.txtprov.ReadOnly = true;
            this.txtprov.Size = new System.Drawing.Size(391, 27);
            this.txtprov.TabIndex = 445;
            // 
            // lbprov
            // 
            this.lbprov.AutoSize = true;
            this.lbprov.BackColor = System.Drawing.Color.Transparent;
            this.lbprov.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbprov.ForeColor = System.Drawing.Color.White;
            this.lbprov.Location = new System.Drawing.Point(629, 10);
            this.lbprov.Name = "lbprov";
            this.lbprov.Size = new System.Drawing.Size(74, 16);
            this.lbprov.TabIndex = 444;
            this.lbprov.Text = "Proveedor:";
            // 
            // txtrchtbl
            // 
            this.txtrchtbl.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtrchtbl.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtrchtbl.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtrchtbl.Location = new System.Drawing.Point(722, 59);
            this.txtrchtbl.MaxLength = 10;
            this.txtrchtbl.Name = "txtrchtbl";
            this.txtrchtbl.ReadOnly = true;
            this.txtrchtbl.Size = new System.Drawing.Size(391, 27);
            this.txtrchtbl.TabIndex = 447;
            // 
            // lbrchtbl
            // 
            this.lbrchtbl.AutoSize = true;
            this.lbrchtbl.BackColor = System.Drawing.Color.Transparent;
            this.lbrchtbl.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbrchtbl.ForeColor = System.Drawing.Color.White;
            this.lbrchtbl.Location = new System.Drawing.Point(593, 62);
            this.lbrchtbl.Name = "lbrchtbl";
            this.lbrchtbl.Size = new System.Drawing.Size(105, 16);
            this.lbrchtbl.TabIndex = 446;
            this.lbrchtbl.Text = "Rancho - tabla:";
            // 
            // txthora
            // 
            this.txthora.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txthora.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txthora.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txthora.Location = new System.Drawing.Point(722, 107);
            this.txthora.MaxLength = 10;
            this.txthora.Name = "txthora";
            this.txthora.ReadOnly = true;
            this.txthora.Size = new System.Drawing.Size(391, 27);
            this.txthora.TabIndex = 449;
            // 
            // lbhora
            // 
            this.lbhora.AutoSize = true;
            this.lbhora.BackColor = System.Drawing.Color.Transparent;
            this.lbhora.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbhora.ForeColor = System.Drawing.Color.White;
            this.lbhora.Location = new System.Drawing.Point(652, 111);
            this.lbhora.Name = "lbhora";
            this.lbhora.Size = new System.Drawing.Size(42, 16);
            this.lbhora.TabIndex = 448;
            this.lbhora.Text = "Hora:";
            // 
            // DGV4
            // 
            this.DGV4.AllowUserToAddRows = false;
            this.DGV4.AllowUserToDeleteRows = false;
            this.DGV4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV4.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.prod_clave,
            this.prod_nombre,
            this.numero_tarimas,
            this.cajas_x_tarima,
            this.esp,
            this.imp,
            this.fec_cad,
            this.grad,
            this.country,
            this.Items});
            this.DGV4.Location = new System.Drawing.Point(17, 167);
            this.DGV4.Name = "DGV4";
            this.DGV4.Size = new System.Drawing.Size(1097, 173);
            this.DGV4.TabIndex = 450;
            this.DGV4.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DGV4_CellContentClick);
            this.DGV4.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.DGV4_CellEndEdit);
            // 
            // prod_clave
            // 
            this.prod_clave.HeaderText = "Código";
            this.prod_clave.Name = "prod_clave";
            this.prod_clave.ReadOnly = true;
            // 
            // prod_nombre
            // 
            this.prod_nombre.HeaderText = "Descripción";
            this.prod_nombre.MaxInputLength = 50;
            this.prod_nombre.Name = "prod_nombre";
            this.prod_nombre.ReadOnly = true;
            // 
            // numero_tarimas
            // 
            this.numero_tarimas.HeaderText = "Num. Tarimas";
            this.numero_tarimas.Name = "numero_tarimas";
            this.numero_tarimas.ReadOnly = true;
            // 
            // cajas_x_tarima
            // 
            this.cajas_x_tarima.HeaderText = "Cajas x Tar";
            this.cajas_x_tarima.Name = "cajas_x_tarima";
            // 
            // esp
            // 
            this.esp.HeaderText = "Especificación";
            this.esp.MaxInputLength = 40;
            this.esp.Name = "esp";
            // 
            // imp
            // 
            this.imp.HeaderText = "Imp";
            this.imp.Name = "imp";
            // 
            // fec_cad
            // 
            this.fec_cad.HeaderText = "Fecha cad";
            this.fec_cad.Name = "fec_cad";
            this.fec_cad.Visible = false;
            // 
            // grad
            // 
            this.grad.HeaderText = "Grado";
            this.grad.Name = "grad";
            // 
            // country
            // 
            this.country.HeaderText = "pais";
            this.country.Name = "country";
            // 
            // Items
            // 
            this.Items.HeaderText = "Items";
            this.Items.Name = "Items";
            // 
            // btnactu
            // 
            this.btnactu.Location = new System.Drawing.Point(416, 347);
            this.btnactu.Name = "btnactu";
            this.btnactu.Size = new System.Drawing.Size(296, 27);
            this.btnactu.TabIndex = 451;
            this.btnactu.Text = "Actualizar";
            this.btnactu.UseVisualStyleBackColor = true;
            this.btnactu.Click += new System.EventHandler(this.btnactu_Click);
            // 
            // DGV5
            // 
            this.DGV5.AllowUserToAddRows = false;
            this.DGV5.AllowUserToDeleteRows = false;
            this.DGV5.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV5.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.clave_producto,
            this.descrip,
            this.box,
            this.tar,
            this.tarimas_total,
            this.especi,
            this.cad_fec,
            this.imp_tar,
            this.id_SSCC,
            this.Item});
            this.DGV5.Location = new System.Drawing.Point(17, 397);
            this.DGV5.Name = "DGV5";
            this.DGV5.Size = new System.Drawing.Size(1097, 173);
            this.DGV5.TabIndex = 452;
            this.DGV5.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.DGV5_CellEndEdit);
            // 
            // clave_producto
            // 
            this.clave_producto.HeaderText = "";
            this.clave_producto.Name = "clave_producto";
            this.clave_producto.ReadOnly = true;
            // 
            // descrip
            // 
            this.descrip.HeaderText = "Header1";
            this.descrip.Name = "descrip";
            this.descrip.ReadOnly = true;
            // 
            // box
            // 
            this.box.HeaderText = "Cajas";
            this.box.Name = "box";
            // 
            // tar
            // 
            this.tar.HeaderText = "Header1";
            this.tar.Name = "tar";
            this.tar.ReadOnly = true;
            // 
            // tarimas_total
            // 
            this.tarimas_total.HeaderText = "Header1";
            this.tarimas_total.Name = "tarimas_total";
            this.tarimas_total.ReadOnly = true;
            // 
            // especi
            // 
            this.especi.HeaderText = "Header1";
            this.especi.MaxInputLength = 40;
            this.especi.Name = "especi";
            // 
            // cad_fec
            // 
            this.cad_fec.HeaderText = "Fecha cad";
            this.cad_fec.Name = "cad_fec";
            this.cad_fec.Visible = false;
            // 
            // imp_tar
            // 
            this.imp_tar.HeaderText = "Imp";
            this.imp_tar.Name = "imp_tar";
            // 
            // id_SSCC
            // 
            this.id_SSCC.HeaderText = "id_SSCC";
            this.id_SSCC.Name = "id_SSCC";
            // 
            // Item
            // 
            this.Item.HeaderText = "Items";
            this.Item.Name = "Item";
            // 
            // btnvalidareti
            // 
            this.btnvalidareti.Enabled = false;
            this.btnvalidareti.Location = new System.Drawing.Point(301, 590);
            this.btnvalidareti.Name = "btnvalidareti";
            this.btnvalidareti.Size = new System.Drawing.Size(233, 27);
            this.btnvalidareti.TabIndex = 453;
            this.btnvalidareti.Text = "Validar etiquetas";
            this.btnvalidareti.UseVisualStyleBackColor = true;
            this.btnvalidareti.Click += new System.EventHandler(this.btnvalidareti_Click);
            // 
            // btnimpeti
            // 
            this.btnimpeti.Enabled = false;
            this.btnimpeti.Location = new System.Drawing.Point(594, 590);
            this.btnimpeti.Name = "btnimpeti";
            this.btnimpeti.Size = new System.Drawing.Size(233, 27);
            this.btnimpeti.TabIndex = 454;
            this.btnimpeti.Text = "Imprimir etiquetas";
            this.btnimpeti.UseVisualStyleBackColor = true;
            this.btnimpeti.Click += new System.EventHandler(this.btnimpeti_Click);
            // 
            // eti_verde
            // 
            this.eti_verde.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDocument2_PrintPage);
            // 
            // lbgrado
            // 
            this.lbgrado.AutoSize = true;
            this.lbgrado.BackColor = System.Drawing.Color.Transparent;
            this.lbgrado.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbgrado.ForeColor = System.Drawing.Color.Black;
            this.lbgrado.Location = new System.Drawing.Point(855, 352);
            this.lbgrado.Name = "lbgrado";
            this.lbgrado.Size = new System.Drawing.Size(123, 16);
            this.lbgrado.TabIndex = 455;
            this.lbgrado.Text = "Seleccionar grado:";
            this.lbgrado.Visible = false;
            // 
            // cbgrado
            // 
            this.cbgrado.FormattingEnabled = true;
            this.cbgrado.Items.AddRange(new object[] {
            "GRADO 2",
            "GRADO 3",
            "GRADO 4"});
            this.cbgrado.Location = new System.Drawing.Point(984, 351);
            this.cbgrado.Name = "cbgrado";
            this.cbgrado.Size = new System.Drawing.Size(129, 23);
            this.cbgrado.TabIndex = 456;
            this.cbgrado.Visible = false;
            this.cbgrado.SelectionChangeCommitted += new System.EventHandler(this.comboBox1_SelectionChangeCommitted);
            // 
            // LblHrCap
            // 
            this.LblHrCap.AutoSize = true;
            this.LblHrCap.BackColor = System.Drawing.Color.Transparent;
            this.LblHrCap.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblHrCap.ForeColor = System.Drawing.Color.White;
            this.LblHrCap.Location = new System.Drawing.Point(722, 144);
            this.LblHrCap.Name = "LblHrCap";
            this.LblHrCap.Size = new System.Drawing.Size(91, 16);
            this.LblHrCap.TabIndex = 457;
            this.LblHrCap.Text = "Hora Captura";
            // 
            // etiqueta_verde
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(1129, 634);
            this.Controls.Add(this.LblHrCap);
            this.Controls.Add(this.cbgrado);
            this.Controls.Add(this.lbgrado);
            this.Controls.Add(this.btnimpeti);
            this.Controls.Add(this.btnvalidareti);
            this.Controls.Add(this.DGV5);
            this.Controls.Add(this.btnactu);
            this.Controls.Add(this.DGV4);
            this.Controls.Add(this.txthora);
            this.Controls.Add(this.lbhora);
            this.Controls.Add(this.txtrchtbl);
            this.Controls.Add(this.lbrchtbl);
            this.Controls.Add(this.txtprov);
            this.Controls.Add(this.lbprov);
            this.Controls.Add(this.txtviaje);
            this.Controls.Add(this.lbviaje);
            this.Controls.Add(this.txtrecibo);
            this.Controls.Add(this.lbrecibo);
            this.Controls.Add(this.txtclave);
            this.Controls.Add(this.lbclave);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "etiqueta_verde";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Impresión de etiquetas";
            this.Load += new System.EventHandler(this.etiqueta_verde_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGV4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV5)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtclave;
        private System.Windows.Forms.Label lbclave;
        private System.Windows.Forms.TextBox txtrecibo;
        private System.Windows.Forms.Label lbrecibo;
        private System.Windows.Forms.TextBox txtviaje;
        private System.Windows.Forms.Label lbviaje;
        private System.Windows.Forms.TextBox txtprov;
        private System.Windows.Forms.Label lbprov;
        private System.Windows.Forms.TextBox txtrchtbl;
        private System.Windows.Forms.Label lbrchtbl;
        private System.Windows.Forms.TextBox txthora;
        private System.Windows.Forms.Label lbhora;
        private System.Windows.Forms.DataGridView DGV4;
        private System.Windows.Forms.Button btnactu;
        private System.Windows.Forms.DataGridView DGV5;
        private System.Windows.Forms.Button btnvalidareti;
        private System.Windows.Forms.Button btnimpeti;
        private System.Drawing.Printing.PrintDocument eti_verde;
        private System.Windows.Forms.Label lbgrado;
        private System.Windows.Forms.ComboBox cbgrado;
        private System.Windows.Forms.Label LblHrCap;
        private System.Windows.Forms.DataGridViewTextBoxColumn prod_clave;
        private System.Windows.Forms.DataGridViewTextBoxColumn prod_nombre;
        private System.Windows.Forms.DataGridViewTextBoxColumn numero_tarimas;
        private System.Windows.Forms.DataGridViewTextBoxColumn cajas_x_tarima;
        private System.Windows.Forms.DataGridViewTextBoxColumn esp;
        private System.Windows.Forms.DataGridViewCheckBoxColumn imp;
        private System.Windows.Forms.DataGridViewTextBoxColumn fec_cad;
        private System.Windows.Forms.DataGridViewTextBoxColumn grad;
        private System.Windows.Forms.DataGridViewTextBoxColumn country;
        private System.Windows.Forms.DataGridViewTextBoxColumn Items;
        private System.Windows.Forms.DataGridViewTextBoxColumn clave_producto;
        private System.Windows.Forms.DataGridViewTextBoxColumn descrip;
        private System.Windows.Forms.DataGridViewTextBoxColumn box;
        private System.Windows.Forms.DataGridViewTextBoxColumn tar;
        private System.Windows.Forms.DataGridViewTextBoxColumn tarimas_total;
        private System.Windows.Forms.DataGridViewTextBoxColumn especi;
        private System.Windows.Forms.DataGridViewTextBoxColumn cad_fec;
        private System.Windows.Forms.DataGridViewCheckBoxColumn imp_tar;
        private System.Windows.Forms.DataGridViewTextBoxColumn id_SSCC;
        private System.Windows.Forms.DataGridViewTextBoxColumn Item;
    }
}