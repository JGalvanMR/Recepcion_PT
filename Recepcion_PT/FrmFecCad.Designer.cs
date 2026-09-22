namespace Recepcion_PT
{
    partial class FrmFecCad
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmFecCad));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.TxtRecibo = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.DtFC = new System.Windows.Forms.DateTimePicker();
            this.DgDatos = new System.Windows.Forms.DataGridView();
            this.BtnSave = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.BtnGene = new System.Windows.Forms.Button();
            this.LblFeEla = new System.Windows.Forms.Label();
            this.NOMBRE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TARIMA = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CAJAS = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FECHACAD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ASIGNAR = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.DgDatos)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(222, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(243, 18);
            this.label1.TabIndex = 0;
            this.label1.Text = "ASIGNAR FECHA DE CADUCIDAD";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(24, 83);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 18);
            this.label2.TabIndex = 1;
            this.label2.Text = "FOLIO:";
            // 
            // TxtRecibo
            // 
            this.TxtRecibo.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtRecibo.Location = new System.Drawing.Point(86, 81);
            this.TxtRecibo.Name = "TxtRecibo";
            this.TxtRecibo.Size = new System.Drawing.Size(100, 24);
            this.TxtRecibo.TabIndex = 2;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(24, 126);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(178, 18);
            this.label3.TabIndex = 3;
            this.label3.Text = "FECHA DE CADUCIDAD:";
            // 
            // DtFC
            // 
            this.DtFC.Enabled = false;
            this.DtFC.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DtFC.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DtFC.Location = new System.Drawing.Point(208, 120);
            this.DtFC.Name = "DtFC";
            this.DtFC.Size = new System.Drawing.Size(111, 26);
            this.DtFC.TabIndex = 4;
            // 
            // DgDatos
            // 
            this.DgDatos.AllowUserToAddRows = false;
            this.DgDatos.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.DgDatos.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.DgDatos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgDatos.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.NOMBRE,
            this.TARIMA,
            this.CAJAS,
            this.FECHACAD,
            this.ASIGNAR});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.DgDatos.DefaultCellStyle = dataGridViewCellStyle2;
            this.DgDatos.Location = new System.Drawing.Point(27, 188);
            this.DgDatos.Name = "DgDatos";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.DgDatos.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.DgDatos.Size = new System.Drawing.Size(662, 260);
            this.DgDatos.TabIndex = 5;
            // 
            // BtnSave
            // 
            this.BtnSave.Enabled = false;
            this.BtnSave.Image = global::Recepcion_PT.Properties.Resources.BtnSave;
            this.BtnSave.Location = new System.Drawing.Point(614, 468);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(61, 59);
            this.BtnSave.TabIndex = 6;
            this.toolTip1.SetToolTip(this.BtnSave, "Guardar Información");
            this.BtnSave.UseVisualStyleBackColor = true;
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // BtnGene
            // 
            this.BtnGene.Image = global::Recepcion_PT.Properties.Resources.Buscar;
            this.BtnGene.Location = new System.Drawing.Point(195, 61);
            this.BtnGene.Name = "BtnGene";
            this.BtnGene.Size = new System.Drawing.Size(52, 42);
            this.BtnGene.TabIndex = 7;
            this.toolTip1.SetToolTip(this.BtnGene, "Consutla Folio");
            this.BtnGene.UseVisualStyleBackColor = true;
            this.BtnGene.Click += new System.EventHandler(this.BtnGene_Click);
            // 
            // LblFeEla
            // 
            this.LblFeEla.AutoSize = true;
            this.LblFeEla.BackColor = System.Drawing.Color.Transparent;
            this.LblFeEla.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblFeEla.ForeColor = System.Drawing.Color.White;
            this.LblFeEla.Location = new System.Drawing.Point(263, 87);
            this.LblFeEla.Name = "LblFeEla";
            this.LblFeEla.Size = new System.Drawing.Size(20, 18);
            this.LblFeEla.TabIndex = 8;
            this.LblFeEla.Text = "...";
            this.LblFeEla.Visible = false;
            // 
            // NOMBRE
            // 
            this.NOMBRE.HeaderText = "NOMBRE";
            this.NOMBRE.Name = "NOMBRE";
            this.NOMBRE.ReadOnly = true;
            this.NOMBRE.Width = 300;
            // 
            // TARIMA
            // 
            this.TARIMA.HeaderText = "TARIMA";
            this.TARIMA.Name = "TARIMA";
            this.TARIMA.ReadOnly = true;
            this.TARIMA.Width = 65;
            // 
            // CAJAS
            // 
            this.CAJAS.HeaderText = "CAJAS";
            this.CAJAS.Name = "CAJAS";
            this.CAJAS.ReadOnly = true;
            this.CAJAS.Width = 50;
            // 
            // FECHACAD
            // 
            this.FECHACAD.HeaderText = "FECHA CAD";
            this.FECHACAD.Name = "FECHACAD";
            this.FECHACAD.Width = 90;
            // 
            // ASIGNAR
            // 
            this.ASIGNAR.HeaderText = "ASIGNAR";
            this.ASIGNAR.Name = "ASIGNAR";
            // 
            // FrmFecCad
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DarkGray;
            this.ClientSize = new System.Drawing.Size(756, 534);
            this.Controls.Add(this.LblFeEla);
            this.Controls.Add(this.BtnGene);
            this.Controls.Add(this.BtnSave);
            this.Controls.Add(this.DgDatos);
            this.Controls.Add(this.DtFC);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.TxtRecibo);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmFecCad";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "2.3 ASIGNAR FECHA DE CADUCIDAD";
            this.Load += new System.EventHandler(this.FrmFecCad_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DgDatos)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TxtRecibo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker DtFC;
        private System.Windows.Forms.DataGridView DgDatos;
        private System.Windows.Forms.Button BtnSave;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button BtnGene;
        private System.Windows.Forms.Label LblFeEla;
        private System.Windows.Forms.DataGridViewTextBoxColumn NOMBRE;
        private System.Windows.Forms.DataGridViewTextBoxColumn TARIMA;
        private System.Windows.Forms.DataGridViewTextBoxColumn CAJAS;
        private System.Windows.Forms.DataGridViewTextBoxColumn FECHACAD;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ASIGNAR;
    }
}