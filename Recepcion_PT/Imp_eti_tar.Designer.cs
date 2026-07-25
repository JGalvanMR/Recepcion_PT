namespace Recepcion_PT
{
    partial class Imp_eti_tar
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Imp_eti_tar));
            this.btnsaveout = new System.Windows.Forms.Button();
            this.btnvalidareti = new System.Windows.Forms.Button();
            this.DGV = new System.Windows.Forms.DataGridView();
            this.clave_prod = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.nom_prod = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.num_tar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.eti = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.num_eti = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.eti_fin = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.imp = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.gra = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.country = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cbselprod = new System.Windows.Forms.ComboBox();
            this.lbselprod = new System.Windows.Forms.Label();
            this.txtrecibo = new System.Windows.Forms.TextBox();
            this.lbrecibo = new System.Windows.Forms.Label();
            this.eti_blanca = new System.Drawing.Printing.PrintDocument();
            this.cbgrado = new System.Windows.Forms.ComboBox();
            this.lbgrado = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.DGV)).BeginInit();
            this.SuspendLayout();
            // 
            // btnsaveout
            // 
            this.btnsaveout.Location = new System.Drawing.Point(612, 251);
            this.btnsaveout.Name = "btnsaveout";
            this.btnsaveout.Size = new System.Drawing.Size(233, 27);
            this.btnsaveout.TabIndex = 643;
            this.btnsaveout.Text = "Cancelar";
            this.btnsaveout.UseVisualStyleBackColor = true;
            this.btnsaveout.Click += new System.EventHandler(this.btnsaveout_Click);
            // 
            // btnvalidareti
            // 
            this.btnvalidareti.Enabled = false;
            this.btnvalidareti.Location = new System.Drawing.Point(134, 251);
            this.btnvalidareti.Name = "btnvalidareti";
            this.btnvalidareti.Size = new System.Drawing.Size(233, 27);
            this.btnvalidareti.TabIndex = 642;
            this.btnvalidareti.Text = "Imprimir etiquetas";
            this.btnvalidareti.UseVisualStyleBackColor = true;
            this.btnvalidareti.Click += new System.EventHandler(this.btnvalidareti_Click);
            // 
            // DGV
            // 
            this.DGV.AllowUserToAddRows = false;
            this.DGV.AllowUserToDeleteRows = false;
            this.DGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.clave_prod,
            this.nom_prod,
            this.num_tar,
            this.eti,
            this.num_eti,
            this.eti_fin,
            this.imp,
            this.gra,
            this.country});
            this.DGV.Location = new System.Drawing.Point(12, 71);
            this.DGV.Name = "DGV";
            this.DGV.RowHeadersVisible = false;
            this.DGV.Size = new System.Drawing.Size(954, 174);
            this.DGV.TabIndex = 641;
            // 
            // clave_prod
            // 
            this.clave_prod.HeaderText = "Clave";
            this.clave_prod.Name = "clave_prod";
            this.clave_prod.ReadOnly = true;
            // 
            // nom_prod
            // 
            this.nom_prod.HeaderText = "Producto";
            this.nom_prod.Name = "nom_prod";
            this.nom_prod.ReadOnly = true;
            // 
            // num_tar
            // 
            this.num_tar.HeaderText = "# Tarima";
            this.num_tar.Name = "num_tar";
            this.num_tar.ReadOnly = true;
            // 
            // eti
            // 
            this.eti.HeaderText = "Etiquetas";
            this.eti.Name = "eti";
            this.eti.ReadOnly = true;
            // 
            // num_eti
            // 
            this.num_eti.HeaderText = "Etiqueta inicial";
            this.num_eti.Name = "num_eti";
            // 
            // eti_fin
            // 
            this.eti_fin.HeaderText = "Etiqueta final";
            this.eti_fin.Name = "eti_fin";
            // 
            // imp
            // 
            this.imp.HeaderText = "";
            this.imp.Name = "imp";
            // 
            // gra
            // 
            this.gra.HeaderText = "Grado";
            this.gra.Name = "gra";
            // 
            // country
            // 
            this.country.HeaderText = "País";
            this.country.Name = "country";
            // 
            // cbselprod
            // 
            this.cbselprod.DropDownHeight = 100;
            this.cbselprod.DropDownWidth = 500;
            this.cbselprod.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbselprod.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.cbselprod.FormattingEnabled = true;
            this.cbselprod.IntegralHeight = false;
            this.cbselprod.Location = new System.Drawing.Point(393, 23);
            this.cbselprod.Name = "cbselprod";
            this.cbselprod.Size = new System.Drawing.Size(293, 24);
            this.cbselprod.TabIndex = 640;
            this.cbselprod.Visible = false;
            this.cbselprod.SelectionChangeCommitted += new System.EventHandler(this.cbselprod_SelectionChangeCommitted);
            // 
            // lbselprod
            // 
            this.lbselprod.AutoSize = true;
            this.lbselprod.BackColor = System.Drawing.Color.Transparent;
            this.lbselprod.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbselprod.ForeColor = System.Drawing.Color.White;
            this.lbselprod.Location = new System.Drawing.Point(236, 26);
            this.lbselprod.Name = "lbselprod";
            this.lbselprod.Size = new System.Drawing.Size(150, 16);
            this.lbselprod.TabIndex = 639;
            this.lbselprod.Text = "Selecccionar producto:";
            this.lbselprod.Visible = false;
            // 
            // txtrecibo
            // 
            this.txtrecibo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtrecibo.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtrecibo.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtrecibo.Location = new System.Drawing.Point(133, 23);
            this.txtrecibo.MaxLength = 10;
            this.txtrecibo.Name = "txtrecibo";
            this.txtrecibo.Size = new System.Drawing.Size(79, 27);
            this.txtrecibo.TabIndex = 638;
            this.txtrecibo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtrecibo_KeyPress);
            // 
            // lbrecibo
            // 
            this.lbrecibo.AutoSize = true;
            this.lbrecibo.BackColor = System.Drawing.Color.Transparent;
            this.lbrecibo.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbrecibo.ForeColor = System.Drawing.Color.White;
            this.lbrecibo.Location = new System.Drawing.Point(53, 26);
            this.lbrecibo.Name = "lbrecibo";
            this.lbrecibo.Size = new System.Drawing.Size(53, 16);
            this.lbrecibo.TabIndex = 637;
            this.lbrecibo.Text = "Recibo:";
            // 
            // eti_blanca
            // 
            this.eti_blanca.BeginPrint += new System.Drawing.Printing.PrintEventHandler(this.eti_blanca_BeginPrint);
            this.eti_blanca.EndPrint += new System.Drawing.Printing.PrintEventHandler(this.eti_blanca_EndPrint);
            this.eti_blanca.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.eti_blanca_PrintPage);
            // 
            // cbgrado
            // 
            this.cbgrado.FormattingEnabled = true;
            this.cbgrado.Items.AddRange(new object[] {
            "GRADO 2",
            "GRADO 3",
            "GRADO 4"});
            this.cbgrado.Location = new System.Drawing.Point(825, 24);
            this.cbgrado.Name = "cbgrado";
            this.cbgrado.Size = new System.Drawing.Size(129, 23);
            this.cbgrado.TabIndex = 645;
            this.cbgrado.Visible = false;
            this.cbgrado.SelectionChangeCommitted += new System.EventHandler(this.cbgrado_SelectionChangeCommitted);
            // 
            // lbgrado
            // 
            this.lbgrado.AutoSize = true;
            this.lbgrado.BackColor = System.Drawing.Color.Transparent;
            this.lbgrado.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbgrado.ForeColor = System.Drawing.Color.White;
            this.lbgrado.Location = new System.Drawing.Point(696, 25);
            this.lbgrado.Name = "lbgrado";
            this.lbgrado.Size = new System.Drawing.Size(122, 16);
            this.lbgrado.TabIndex = 644;
            this.lbgrado.Text = "Seleccionar grado:";
            this.lbgrado.Visible = false;
            // 
            // Imp_eti_tar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(978, 301);
            this.Controls.Add(this.cbgrado);
            this.Controls.Add(this.lbgrado);
            this.Controls.Add(this.btnsaveout);
            this.Controls.Add(this.btnvalidareti);
            this.Controls.Add(this.DGV);
            this.Controls.Add(this.cbselprod);
            this.Controls.Add(this.lbselprod);
            this.Controls.Add(this.txtrecibo);
            this.Controls.Add(this.lbrecibo);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Imp_eti_tar";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Imprimir etiqueta por tarima";
            this.Load += new System.EventHandler(this.Imp_eti_tar_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGV)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnsaveout;
        private System.Windows.Forms.Button btnvalidareti;
        private System.Windows.Forms.DataGridView DGV;
        private System.Windows.Forms.ComboBox cbselprod;
        private System.Windows.Forms.Label lbselprod;
        private System.Windows.Forms.TextBox txtrecibo;
        private System.Windows.Forms.Label lbrecibo;
        private System.Drawing.Printing.PrintDocument eti_blanca;
        private System.Windows.Forms.ComboBox cbgrado;
        private System.Windows.Forms.Label lbgrado;
        private System.Windows.Forms.DataGridViewTextBoxColumn clave_prod;
        private System.Windows.Forms.DataGridViewTextBoxColumn nom_prod;
        private System.Windows.Forms.DataGridViewTextBoxColumn num_tar;
        private System.Windows.Forms.DataGridViewTextBoxColumn eti;
        private System.Windows.Forms.DataGridViewTextBoxColumn num_eti;
        private System.Windows.Forms.DataGridViewTextBoxColumn eti_fin;
        private System.Windows.Forms.DataGridViewCheckBoxColumn imp;
        private System.Windows.Forms.DataGridViewTextBoxColumn gra;
        private System.Windows.Forms.DataGridViewTextBoxColumn country;
    }
}