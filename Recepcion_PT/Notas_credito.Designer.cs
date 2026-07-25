namespace Recepcion_PT
{
    partial class Notas_credito
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Notas_credito));
            this.DGV6 = new System.Windows.Forms.DataGridView();
            this.rpt_recibo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fecha = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clave = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.nombre = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.linea = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.motivo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.rpt_tipo_recepcion = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fcn_folio = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.linclave = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DGV7 = new System.Windows.Forms.DataGridView();
            this.prodclave = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.prodnombre = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.canti = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.afect = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tiporecep = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lin = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.linnom = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.envnom = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.not_cre = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.en_pes = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.DGV6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV7)).BeginInit();
            this.SuspendLayout();
            // 
            // DGV6
            // 
            this.DGV6.AllowUserToAddRows = false;
            this.DGV6.AllowUserToDeleteRows = false;
            this.DGV6.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV6.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.rpt_recibo,
            this.fecha,
            this.clave,
            this.nombre,
            this.linea,
            this.motivo,
            this.rpt_tipo_recepcion,
            this.fcn_folio,
            this.linclave});
            this.DGV6.Location = new System.Drawing.Point(12, 26);
            this.DGV6.Name = "DGV6";
            this.DGV6.ReadOnly = true;
            this.DGV6.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.DGV6.Size = new System.Drawing.Size(860, 297);
            this.DGV6.TabIndex = 0;
            this.DGV6.RowEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.DGV6_RowEnter);
            this.DGV6.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.DGV6_KeyPress);
            // 
            // rpt_recibo
            // 
            this.rpt_recibo.HeaderText = "N. Crédito";
            this.rpt_recibo.Name = "rpt_recibo";
            this.rpt_recibo.ReadOnly = true;
            // 
            // fecha
            // 
            this.fecha.HeaderText = "Fecha";
            this.fecha.Name = "fecha";
            this.fecha.ReadOnly = true;
            // 
            // clave
            // 
            this.clave.HeaderText = "Clave";
            this.clave.Name = "clave";
            this.clave.ReadOnly = true;
            // 
            // nombre
            // 
            this.nombre.HeaderText = "Nombre";
            this.nombre.Name = "nombre";
            this.nombre.ReadOnly = true;
            // 
            // linea
            // 
            this.linea.HeaderText = "Línea";
            this.linea.Name = "linea";
            this.linea.ReadOnly = true;
            // 
            // motivo
            // 
            this.motivo.HeaderText = "Motivo";
            this.motivo.Name = "motivo";
            this.motivo.ReadOnly = true;
            // 
            // rpt_tipo_recepcion
            // 
            this.rpt_tipo_recepcion.HeaderText = "Tipo recepción";
            this.rpt_tipo_recepcion.Name = "rpt_tipo_recepcion";
            this.rpt_tipo_recepcion.ReadOnly = true;
            // 
            // fcn_folio
            // 
            this.fcn_folio.HeaderText = "fcn folio";
            this.fcn_folio.Name = "fcn_folio";
            this.fcn_folio.ReadOnly = true;
            // 
            // linclave
            // 
            this.linclave.HeaderText = "linclave";
            this.linclave.Name = "linclave";
            this.linclave.ReadOnly = true;
            this.linclave.Visible = false;
            // 
            // DGV7
            // 
            this.DGV7.AllowUserToAddRows = false;
            this.DGV7.AllowUserToDeleteRows = false;
            this.DGV7.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV7.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.prodclave,
            this.prodnombre,
            this.canti,
            this.afect,
            this.tiporecep,
            this.lin,
            this.linnom,
            this.envnom,
            this.not_cre,
            this.en_pes});
            this.DGV7.Location = new System.Drawing.Point(79, 344);
            this.DGV7.Name = "DGV7";
            this.DGV7.ReadOnly = true;
            this.DGV7.Size = new System.Drawing.Size(698, 183);
            this.DGV7.TabIndex = 1;
            this.DGV7.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DGV7_CellDoubleClick);
            // 
            // prodclave
            // 
            this.prodclave.HeaderText = "Clave";
            this.prodclave.Name = "prodclave";
            this.prodclave.ReadOnly = true;
            // 
            // prodnombre
            // 
            this.prodnombre.HeaderText = "Nombre";
            this.prodnombre.Name = "prodnombre";
            this.prodnombre.ReadOnly = true;
            // 
            // canti
            // 
            this.canti.HeaderText = "Cantidad";
            this.canti.Name = "canti";
            this.canti.ReadOnly = true;
            // 
            // afect
            // 
            this.afect.HeaderText = "Afectado";
            this.afect.Name = "afect";
            this.afect.ReadOnly = true;
            // 
            // tiporecep
            // 
            this.tiporecep.HeaderText = "mot";
            this.tiporecep.Name = "tiporecep";
            this.tiporecep.ReadOnly = true;
            // 
            // lin
            // 
            this.lin.HeaderText = "Linea";
            this.lin.Name = "lin";
            this.lin.ReadOnly = true;
            this.lin.Visible = false;
            // 
            // linnom
            // 
            this.linnom.HeaderText = "nombre linea";
            this.linnom.Name = "linnom";
            this.linnom.ReadOnly = true;
            this.linnom.Visible = false;
            // 
            // envnom
            // 
            this.envnom.HeaderText = "envase nombre";
            this.envnom.Name = "envnom";
            this.envnom.ReadOnly = true;
            this.envnom.Visible = false;
            // 
            // not_cre
            // 
            this.not_cre.HeaderText = "N. Credito";
            this.not_cre.Name = "not_cre";
            this.not_cre.ReadOnly = true;
            this.not_cre.Visible = false;
            // 
            // en_pes
            // 
            this.en_pes.HeaderText = "envase peso";
            this.en_pes.Name = "en_pes";
            this.en_pes.ReadOnly = true;
            this.en_pes.Visible = false;
            // 
            // Notas_credito
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(884, 545);
            this.Controls.Add(this.DGV7);
            this.Controls.Add(this.DGV6);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Notas_credito";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Notas de crédito pendientes por afectar inventarios";
            this.Load += new System.EventHandler(this.Notas_credito_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGV6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV7)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView DGV6;
        private System.Windows.Forms.DataGridView DGV7;
        private System.Windows.Forms.DataGridViewTextBoxColumn rpt_recibo;
        private System.Windows.Forms.DataGridViewTextBoxColumn fecha;
        private System.Windows.Forms.DataGridViewTextBoxColumn clave;
        private System.Windows.Forms.DataGridViewTextBoxColumn nombre;
        private System.Windows.Forms.DataGridViewTextBoxColumn linea;
        private System.Windows.Forms.DataGridViewTextBoxColumn motivo;
        private System.Windows.Forms.DataGridViewTextBoxColumn rpt_tipo_recepcion;
        private System.Windows.Forms.DataGridViewTextBoxColumn fcn_folio;
        private System.Windows.Forms.DataGridViewTextBoxColumn linclave;
        private System.Windows.Forms.DataGridViewTextBoxColumn prodclave;
        private System.Windows.Forms.DataGridViewTextBoxColumn prodnombre;
        private System.Windows.Forms.DataGridViewTextBoxColumn canti;
        private System.Windows.Forms.DataGridViewTextBoxColumn afect;
        private System.Windows.Forms.DataGridViewTextBoxColumn tiporecep;
        private System.Windows.Forms.DataGridViewTextBoxColumn lin;
        private System.Windows.Forms.DataGridViewTextBoxColumn linnom;
        private System.Windows.Forms.DataGridViewTextBoxColumn envnom;
        private System.Windows.Forms.DataGridViewTextBoxColumn not_cre;
        private System.Windows.Forms.DataGridViewTextBoxColumn en_pes;
    }
}