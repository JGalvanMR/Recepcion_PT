namespace Recepcion_PT
{
    partial class peso_x_tarimas
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(peso_x_tarimas));
            this.DGV = new System.Windows.Forms.DataGridView();
            this.btnguardar = new System.Windows.Forms.Button();
            this.btnexit = new System.Windows.Forms.Button();
            this.product = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.peso_tar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.num_box = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.enva = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.peso_env = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.peso_total = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.peso_uni = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.DGV)).BeginInit();
            this.SuspendLayout();
            // 
            // DGV
            // 
            this.DGV.AllowUserToAddRows = false;
            this.DGV.AllowUserToDeleteRows = false;
            this.DGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.product,
            this.name,
            this.peso_tar,
            this.num_box,
            this.enva,
            this.peso_env,
            this.peso_total,
            this.peso_uni});
            this.DGV.Location = new System.Drawing.Point(0, 22);
            this.DGV.Name = "DGV";
            this.DGV.RowHeadersVisible = false;
            this.DGV.Size = new System.Drawing.Size(943, 150);
            this.DGV.TabIndex = 0;
            this.DGV.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.DGV_CellEndEdit);
            // 
            // btnguardar
            // 
            this.btnguardar.BackgroundImage = global::Recepcion_PT.Properties.Resources.save;
            this.btnguardar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnguardar.Location = new System.Drawing.Point(12, 178);
            this.btnguardar.Name = "btnguardar";
            this.btnguardar.Size = new System.Drawing.Size(56, 49);
            this.btnguardar.TabIndex = 1;
            this.btnguardar.UseVisualStyleBackColor = true;
            this.btnguardar.Click += new System.EventHandler(this.btnguardar_Click);
            // 
            // btnexit
            // 
            this.btnexit.BackgroundImage = global::Recepcion_PT.Properties.Resources.salir;
            this.btnexit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnexit.Location = new System.Drawing.Point(876, 178);
            this.btnexit.Name = "btnexit";
            this.btnexit.Size = new System.Drawing.Size(56, 49);
            this.btnexit.TabIndex = 2;
            this.btnexit.UseVisualStyleBackColor = true;
            this.btnexit.Click += new System.EventHandler(this.btnexit_Click);
            // 
            // product
            // 
            this.product.HeaderText = "Producto";
            this.product.Name = "product";
            this.product.ReadOnly = true;
            // 
            // name
            // 
            this.name.HeaderText = "Nombre";
            this.name.Name = "name";
            this.name.ReadOnly = true;
            // 
            // peso_tar
            // 
            this.peso_tar.HeaderText = "Peso tarima";
            this.peso_tar.Name = "peso_tar";
            // 
            // num_box
            // 
            this.num_box.HeaderText = "# cajas x tarimas";
            this.num_box.Name = "num_box";
            // 
            // enva
            // 
            this.enva.HeaderText = "Envase";
            this.enva.Name = "enva";
            this.enva.ReadOnly = true;
            // 
            // peso_env
            // 
            this.peso_env.HeaderText = "Peso envase";
            this.peso_env.Name = "peso_env";
            this.peso_env.ReadOnly = true;
            // 
            // peso_total
            // 
            this.peso_total.HeaderText = "Peso total";
            this.peso_total.Name = "peso_total";
            this.peso_total.ReadOnly = true;
            // 
            // peso_uni
            // 
            this.peso_uni.HeaderText = "Peso unitario";
            this.peso_uni.Name = "peso_uni";
            this.peso_uni.ReadOnly = true;
            // 
            // peso_x_tarimas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(944, 241);
            this.ControlBox = false;
            this.Controls.Add(this.btnexit);
            this.Controls.Add(this.btnguardar);
            this.Controls.Add(this.DGV);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "peso_x_tarimas";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Peso por tarima del producto";
            this.Load += new System.EventHandler(this.peso_x_tarimas_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGV)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView DGV;
        private System.Windows.Forms.Button btnguardar;
        private System.Windows.Forms.Button btnexit;
        private System.Windows.Forms.DataGridViewTextBoxColumn product;
        private System.Windows.Forms.DataGridViewTextBoxColumn name;
        private System.Windows.Forms.DataGridViewTextBoxColumn peso_tar;
        private System.Windows.Forms.DataGridViewTextBoxColumn num_box;
        private System.Windows.Forms.DataGridViewTextBoxColumn enva;
        private System.Windows.Forms.DataGridViewTextBoxColumn peso_env;
        private System.Windows.Forms.DataGridViewTextBoxColumn peso_total;
        private System.Windows.Forms.DataGridViewTextBoxColumn peso_uni;
    }
}