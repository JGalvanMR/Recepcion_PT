namespace Recepcion_PT
{
    partial class etiqueta_blanca
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(etiqueta_blanca));
            this.btnSalir = new System.Windows.Forms.Button();
            this.btnImp = new System.Windows.Forms.Button();
            this.eti_blanca = new System.Drawing.Printing.PrintDocument();
            this.SuspendLayout();
            // 
            // btnSalir
            // 
            this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnSalir.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSalir.ForeColor = System.Drawing.Color.Black;
            this.btnSalir.Location = new System.Drawing.Point(96, 80);
            this.btnSalir.Name = "btnSalir";
            this.btnSalir.Size = new System.Drawing.Size(137, 27);
            this.btnSalir.TabIndex = 111;
            this.btnSalir.Text = "SALIR";
            this.btnSalir.UseVisualStyleBackColor = false;
            this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
            // 
            // btnImp
            // 
            this.btnImp.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnImp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnImp.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnImp.ForeColor = System.Drawing.Color.Black;
            this.btnImp.Location = new System.Drawing.Point(97, 32);
            this.btnImp.Name = "btnImp";
            this.btnImp.Size = new System.Drawing.Size(137, 27);
            this.btnImp.TabIndex = 110;
            this.btnImp.Text = "IMPRIMIR";
            this.btnImp.UseVisualStyleBackColor = false;
            this.btnImp.Click += new System.EventHandler(this.btnImp_Click);
            // 
            // eti_blanca
            // 
            this.eti_blanca.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.eti_blanca_PrintPage);
            // 
            // etiqueta_blanca
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(331, 138);
            this.Controls.Add(this.btnSalir);
            this.Controls.Add(this.btnImp);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "etiqueta_blanca";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Impresión de etiquetas";
            this.Load += new System.EventHandler(this.etiqueta_blanca_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnSalir;
        private System.Windows.Forms.Button btnImp;
        private System.Drawing.Printing.PrintDocument eti_blanca;
    }
}