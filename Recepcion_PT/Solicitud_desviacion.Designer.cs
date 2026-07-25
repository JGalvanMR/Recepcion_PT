namespace Recepcion_PT
{
    partial class Solicitud_desviacion
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Solicitud_desviacion));
            this.txtfolio = new System.Windows.Forms.TextBox();
            this.lbfolio = new System.Windows.Forms.Label();
            this.DTPFec = new System.Windows.Forms.DateTimePicker();
            this.lbfecha = new System.Windows.Forms.Label();
            this.lbnumnc = new System.Windows.Forms.Label();
            this.txtnoconformidad = new System.Windows.Forms.TextBox();
            this.cbturno = new System.Windows.Forms.ComboBox();
            this.lbturno = new System.Windows.Forms.Label();
            this.lbprod = new System.Windows.Forms.Label();
            this.txtprod = new System.Windows.Forms.TextBox();
            this.NUDCanti = new System.Windows.Forms.NumericUpDown();
            this.lbcant = new System.Windows.Forms.Label();
            this.txtesp = new System.Windows.Forms.TextBox();
            this.lbesp = new System.Windows.Forms.Label();
            this.txtnc = new System.Windows.Forms.TextBox();
            this.lbnc = new System.Windows.Forms.Label();
            this.lbds = new System.Windows.Forms.Label();
            this.txtds = new System.Windows.Forms.TextBox();
            this.lbmotivo = new System.Windows.Forms.Label();
            this.txtmot = new System.Windows.Forms.TextBox();
            this.RBsi1 = new System.Windows.Forms.RadioButton();
            this.RB1no = new System.Windows.Forms.RadioButton();
            this.GBAprovcnte = new System.Windows.Forms.GroupBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.RBrechazado = new System.Windows.Forms.RadioButton();
            this.RBaceptado = new System.Windows.Forms.RadioButton();
            this.lbrepcnte = new System.Windows.Forms.Label();
            this.txtreprecnte = new System.Windows.Forms.TextBox();
            this.DTPcnte = new System.Windows.Forms.DateTimePicker();
            this.lbfeccnte = new System.Windows.Forms.Label();
            this.txtevidencia = new System.Windows.Forms.TextBox();
            this.lbevi = new System.Windows.Forms.Label();
            this.txtcom = new System.Windows.Forms.TextBox();
            this.lbcom = new System.Windows.Forms.Label();
            this.RBrechazada = new System.Windows.Forms.RadioButton();
            this.RBaceptada = new System.Windows.Forms.RadioButton();
            this.btnsalir = new System.Windows.Forms.Button();
            this.btnguardar = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.NUDCanti)).BeginInit();
            this.GBAprovcnte.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtfolio
            // 
            this.txtfolio.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtfolio.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtfolio.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtfolio.Location = new System.Drawing.Point(61, 6);
            this.txtfolio.MaxLength = 10;
            this.txtfolio.Name = "txtfolio";
            this.txtfolio.ReadOnly = true;
            this.txtfolio.Size = new System.Drawing.Size(149, 27);
            this.txtfolio.TabIndex = 459;
            // 
            // lbfolio
            // 
            this.lbfolio.AutoSize = true;
            this.lbfolio.BackColor = System.Drawing.Color.Transparent;
            this.lbfolio.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbfolio.ForeColor = System.Drawing.Color.White;
            this.lbfolio.Location = new System.Drawing.Point(12, 9);
            this.lbfolio.Name = "lbfolio";
            this.lbfolio.Size = new System.Drawing.Size(43, 16);
            this.lbfolio.TabIndex = 442;
            this.lbfolio.Text = "Folio:";
            // 
            // DTPFec
            // 
            this.DTPFec.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DTPFec.Location = new System.Drawing.Point(293, 9);
            this.DTPFec.Name = "DTPFec";
            this.DTPFec.Size = new System.Drawing.Size(110, 24);
            this.DTPFec.TabIndex = 445;
            // 
            // lbfecha
            // 
            this.lbfecha.AutoSize = true;
            this.lbfecha.BackColor = System.Drawing.Color.Transparent;
            this.lbfecha.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbfecha.ForeColor = System.Drawing.Color.White;
            this.lbfecha.Location = new System.Drawing.Point(239, 9);
            this.lbfecha.Name = "lbfecha";
            this.lbfecha.Size = new System.Drawing.Size(48, 16);
            this.lbfecha.TabIndex = 444;
            this.lbfecha.Text = "Fecha:";
            // 
            // lbnumnc
            // 
            this.lbnumnc.AutoSize = true;
            this.lbnumnc.BackColor = System.Drawing.Color.Transparent;
            this.lbnumnc.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbnumnc.ForeColor = System.Drawing.Color.White;
            this.lbnumnc.Location = new System.Drawing.Point(412, 9);
            this.lbnumnc.Name = "lbnumnc";
            this.lbnumnc.Size = new System.Drawing.Size(126, 16);
            this.lbnumnc.TabIndex = 446;
            this.lbnumnc.Text = "# No conformidad:";
            // 
            // txtnoconformidad
            // 
            this.txtnoconformidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtnoconformidad.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtnoconformidad.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtnoconformidad.Location = new System.Drawing.Point(544, 6);
            this.txtnoconformidad.MaxLength = 10;
            this.txtnoconformidad.Name = "txtnoconformidad";
            this.txtnoconformidad.ReadOnly = true;
            this.txtnoconformidad.Size = new System.Drawing.Size(149, 27);
            this.txtnoconformidad.TabIndex = 447;
            // 
            // cbturno
            // 
            this.cbturno.FormattingEnabled = true;
            this.cbturno.Items.AddRange(new object[] {
            "1",
            "2"});
            this.cbturno.Location = new System.Drawing.Point(766, 7);
            this.cbturno.Name = "cbturno";
            this.cbturno.Size = new System.Drawing.Size(37, 23);
            this.cbturno.TabIndex = 449;
            // 
            // lbturno
            // 
            this.lbturno.AutoSize = true;
            this.lbturno.BackColor = System.Drawing.Color.Transparent;
            this.lbturno.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbturno.ForeColor = System.Drawing.Color.White;
            this.lbturno.Location = new System.Drawing.Point(711, 9);
            this.lbturno.Name = "lbturno";
            this.lbturno.Size = new System.Drawing.Size(49, 16);
            this.lbturno.TabIndex = 448;
            this.lbturno.Text = "Turno:";
            // 
            // lbprod
            // 
            this.lbprod.AutoSize = true;
            this.lbprod.BackColor = System.Drawing.Color.Transparent;
            this.lbprod.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbprod.ForeColor = System.Drawing.Color.White;
            this.lbprod.Location = new System.Drawing.Point(12, 46);
            this.lbprod.Name = "lbprod";
            this.lbprod.Size = new System.Drawing.Size(83, 16);
            this.lbprod.TabIndex = 450;
            this.lbprod.Text = "Producto(s):";
            // 
            // txtprod
            // 
            this.txtprod.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtprod.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtprod.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtprod.Location = new System.Drawing.Point(101, 43);
            this.txtprod.MaxLength = 10;
            this.txtprod.Name = "txtprod";
            this.txtprod.ReadOnly = true;
            this.txtprod.Size = new System.Drawing.Size(302, 27);
            this.txtprod.TabIndex = 451;
            // 
            // NUDCanti
            // 
            this.NUDCanti.Location = new System.Drawing.Point(86, 78);
            this.NUDCanti.Name = "NUDCanti";
            this.NUDCanti.Size = new System.Drawing.Size(102, 24);
            this.NUDCanti.TabIndex = 453;
            // 
            // lbcant
            // 
            this.lbcant.AutoSize = true;
            this.lbcant.BackColor = System.Drawing.Color.Transparent;
            this.lbcant.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbcant.ForeColor = System.Drawing.Color.White;
            this.lbcant.Location = new System.Drawing.Point(12, 78);
            this.lbcant.Name = "lbcant";
            this.lbcant.Size = new System.Drawing.Size(68, 16);
            this.lbcant.TabIndex = 452;
            this.lbcant.Text = "Cantidad:";
            // 
            // txtesp
            // 
            this.txtesp.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtesp.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtesp.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtesp.Location = new System.Drawing.Point(509, 43);
            this.txtesp.MaxLength = 50;
            this.txtesp.Multiline = true;
            this.txtesp.Name = "txtesp";
            this.txtesp.Size = new System.Drawing.Size(294, 61);
            this.txtesp.TabIndex = 455;
            // 
            // lbesp
            // 
            this.lbesp.AutoSize = true;
            this.lbesp.BackColor = System.Drawing.Color.Transparent;
            this.lbesp.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbesp.ForeColor = System.Drawing.Color.White;
            this.lbesp.Location = new System.Drawing.Point(412, 46);
            this.lbesp.Name = "lbesp";
            this.lbesp.Size = new System.Drawing.Size(91, 16);
            this.lbesp.TabIndex = 454;
            this.lbesp.Text = "Especificado:";
            // 
            // txtnc
            // 
            this.txtnc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtnc.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtnc.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtnc.Location = new System.Drawing.Point(132, 108);
            this.txtnc.MaxLength = 50;
            this.txtnc.Multiline = true;
            this.txtnc.Name = "txtnc";
            this.txtnc.Size = new System.Drawing.Size(271, 61);
            this.txtnc.TabIndex = 443;
            // 
            // lbnc
            // 
            this.lbnc.AutoSize = true;
            this.lbnc.BackColor = System.Drawing.Color.Transparent;
            this.lbnc.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbnc.ForeColor = System.Drawing.Color.White;
            this.lbnc.Location = new System.Drawing.Point(12, 111);
            this.lbnc.Name = "lbnc";
            this.lbnc.Size = new System.Drawing.Size(114, 16);
            this.lbnc.TabIndex = 456;
            this.lbnc.Text = "No conformidad:";
            // 
            // lbds
            // 
            this.lbds.AutoSize = true;
            this.lbds.BackColor = System.Drawing.Color.Transparent;
            this.lbds.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbds.ForeColor = System.Drawing.Color.White;
            this.lbds.Location = new System.Drawing.Point(413, 119);
            this.lbds.Name = "lbds";
            this.lbds.Size = new System.Drawing.Size(146, 16);
            this.lbds.TabIndex = 458;
            this.lbds.Text = "Desviación solicitada:";
            // 
            // txtds
            // 
            this.txtds.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtds.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtds.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtds.Location = new System.Drawing.Point(565, 111);
            this.txtds.MaxLength = 50;
            this.txtds.Multiline = true;
            this.txtds.Name = "txtds";
            this.txtds.Size = new System.Drawing.Size(239, 61);
            this.txtds.TabIndex = 444;
            // 
            // lbmotivo
            // 
            this.lbmotivo.AutoSize = true;
            this.lbmotivo.BackColor = System.Drawing.Color.Transparent;
            this.lbmotivo.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbmotivo.ForeColor = System.Drawing.Color.Black;
            this.lbmotivo.Location = new System.Drawing.Point(12, 176);
            this.lbmotivo.Name = "lbmotivo";
            this.lbmotivo.Size = new System.Drawing.Size(227, 16);
            this.lbmotivo.TabIndex = 460;
            this.lbmotivo.Text = "Motivo para solicitar la desviación:";
            // 
            // txtmot
            // 
            this.txtmot.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtmot.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtmot.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtmot.Location = new System.Drawing.Point(15, 195);
            this.txtmot.MaxLength = 50;
            this.txtmot.Multiline = true;
            this.txtmot.Name = "txtmot";
            this.txtmot.Size = new System.Drawing.Size(388, 61);
            this.txtmot.TabIndex = 445;
            // 
            // RBsi1
            // 
            this.RBsi1.AutoSize = true;
            this.RBsi1.Location = new System.Drawing.Point(8, 20);
            this.RBsi1.Name = "RBsi1";
            this.RBsi1.Size = new System.Drawing.Size(34, 19);
            this.RBsi1.TabIndex = 463;
            this.RBsi1.TabStop = true;
            this.RBsi1.Text = "SI";
            this.RBsi1.UseVisualStyleBackColor = true;
            // 
            // RB1no
            // 
            this.RB1no.AutoSize = true;
            this.RB1no.Location = new System.Drawing.Point(147, 20);
            this.RB1no.Name = "RB1no";
            this.RB1no.Size = new System.Drawing.Size(42, 19);
            this.RB1no.TabIndex = 464;
            this.RB1no.TabStop = true;
            this.RB1no.Text = "NO";
            this.RB1no.UseVisualStyleBackColor = true;
            // 
            // GBAprovcnte
            // 
            this.GBAprovcnte.BackColor = System.Drawing.Color.Transparent;
            this.GBAprovcnte.Controls.Add(this.RB1no);
            this.GBAprovcnte.Controls.Add(this.RBsi1);
            this.GBAprovcnte.Location = new System.Drawing.Point(408, 195);
            this.GBAprovcnte.Name = "GBAprovcnte";
            this.GBAprovcnte.Size = new System.Drawing.Size(195, 45);
            this.GBAprovcnte.TabIndex = 446;
            this.GBAprovcnte.TabStop = false;
            this.GBAprovcnte.Text = "Requiere aprobación del cliente:";
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.Transparent;
            this.groupBox1.Controls.Add(this.RBrechazado);
            this.groupBox1.Controls.Add(this.RBaceptado);
            this.groupBox1.Location = new System.Drawing.Point(609, 195);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(195, 45);
            this.groupBox1.TabIndex = 447;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Decisión del cliente:";
            // 
            // RBrechazado
            // 
            this.RBrechazado.AutoSize = true;
            this.RBrechazado.Location = new System.Drawing.Point(101, 20);
            this.RBrechazado.Name = "RBrechazado";
            this.RBrechazado.Size = new System.Drawing.Size(83, 19);
            this.RBrechazado.TabIndex = 464;
            this.RBrechazado.TabStop = true;
            this.RBrechazado.Text = "Rechazado";
            this.RBrechazado.UseVisualStyleBackColor = true;
            // 
            // RBaceptado
            // 
            this.RBaceptado.AutoSize = true;
            this.RBaceptado.Location = new System.Drawing.Point(8, 20);
            this.RBaceptado.Name = "RBaceptado";
            this.RBaceptado.Size = new System.Drawing.Size(76, 19);
            this.RBaceptado.TabIndex = 463;
            this.RBaceptado.TabStop = true;
            this.RBaceptado.Text = "Aceptado";
            this.RBaceptado.UseVisualStyleBackColor = true;
            // 
            // lbrepcnte
            // 
            this.lbrepcnte.AutoSize = true;
            this.lbrepcnte.BackColor = System.Drawing.Color.Transparent;
            this.lbrepcnte.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbrepcnte.ForeColor = System.Drawing.Color.Black;
            this.lbrepcnte.Location = new System.Drawing.Point(12, 270);
            this.lbrepcnte.Name = "lbrepcnte";
            this.lbrepcnte.Size = new System.Drawing.Size(170, 16);
            this.lbrepcnte.TabIndex = 466;
            this.lbrepcnte.Text = "Representante del cliente:";
            // 
            // txtreprecnte
            // 
            this.txtreprecnte.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtreprecnte.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtreprecnte.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtreprecnte.Location = new System.Drawing.Point(12, 289);
            this.txtreprecnte.MaxLength = 50;
            this.txtreprecnte.Multiline = true;
            this.txtreprecnte.Name = "txtreprecnte";
            this.txtreprecnte.Size = new System.Drawing.Size(388, 22);
            this.txtreprecnte.TabIndex = 448;
            // 
            // DTPcnte
            // 
            this.DTPcnte.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DTPcnte.Location = new System.Drawing.Point(477, 286);
            this.DTPcnte.Name = "DTPcnte";
            this.DTPcnte.Size = new System.Drawing.Size(110, 24);
            this.DTPcnte.TabIndex = 449;
            // 
            // lbfeccnte
            // 
            this.lbfeccnte.AutoSize = true;
            this.lbfeccnte.BackColor = System.Drawing.Color.Transparent;
            this.lbfeccnte.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbfeccnte.ForeColor = System.Drawing.Color.Black;
            this.lbfeccnte.Location = new System.Drawing.Point(417, 286);
            this.lbfeccnte.Name = "lbfeccnte";
            this.lbfeccnte.Size = new System.Drawing.Size(48, 16);
            this.lbfeccnte.TabIndex = 468;
            this.lbfeccnte.Text = "Fecha:";
            // 
            // txtevidencia
            // 
            this.txtevidencia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtevidencia.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtevidencia.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtevidencia.Location = new System.Drawing.Point(15, 340);
            this.txtevidencia.MaxLength = 50;
            this.txtevidencia.Multiline = true;
            this.txtevidencia.Name = "txtevidencia";
            this.txtevidencia.Size = new System.Drawing.Size(388, 22);
            this.txtevidencia.TabIndex = 450;
            // 
            // lbevi
            // 
            this.lbevi.AutoSize = true;
            this.lbevi.BackColor = System.Drawing.Color.Transparent;
            this.lbevi.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbevi.ForeColor = System.Drawing.Color.Black;
            this.lbevi.Location = new System.Drawing.Point(15, 321);
            this.lbevi.Name = "lbevi";
            this.lbevi.Size = new System.Drawing.Size(71, 16);
            this.lbevi.TabIndex = 470;
            this.lbevi.Text = "Evidencia:";
            // 
            // txtcom
            // 
            this.txtcom.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtcom.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtcom.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtcom.Location = new System.Drawing.Point(420, 340);
            this.txtcom.MaxLength = 50;
            this.txtcom.Multiline = true;
            this.txtcom.Name = "txtcom";
            this.txtcom.Size = new System.Drawing.Size(388, 61);
            this.txtcom.TabIndex = 451;
            // 
            // lbcom
            // 
            this.lbcom.AutoSize = true;
            this.lbcom.BackColor = System.Drawing.Color.Transparent;
            this.lbcom.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbcom.ForeColor = System.Drawing.Color.Black;
            this.lbcom.Location = new System.Drawing.Point(417, 321);
            this.lbcom.Name = "lbcom";
            this.lbcom.Size = new System.Drawing.Size(92, 16);
            this.lbcom.TabIndex = 472;
            this.lbcom.Text = "Comentarios:";
            // 
            // RBrechazada
            // 
            this.RBrechazada.AutoSize = true;
            this.RBrechazada.BackColor = System.Drawing.Color.Transparent;
            this.RBrechazada.Location = new System.Drawing.Point(111, 382);
            this.RBrechazada.Name = "RBrechazada";
            this.RBrechazada.Size = new System.Drawing.Size(93, 19);
            this.RBrechazada.TabIndex = 466;
            this.RBrechazada.TabStop = true;
            this.RBrechazada.Text = "RECHAZADA";
            this.RBrechazada.UseVisualStyleBackColor = false;
            // 
            // RBaceptada
            // 
            this.RBaceptada.AutoSize = true;
            this.RBaceptada.BackColor = System.Drawing.Color.Transparent;
            this.RBaceptada.Location = new System.Drawing.Point(18, 382);
            this.RBaceptada.Name = "RBaceptada";
            this.RBaceptada.Size = new System.Drawing.Size(84, 19);
            this.RBaceptada.TabIndex = 452;
            this.RBaceptada.TabStop = true;
            this.RBaceptada.Text = "ACEPTADA";
            this.RBaceptada.UseVisualStyleBackColor = false;
            // 
            // btnsalir
            // 
            this.btnsalir.BackgroundImage = global::Recepcion_PT.Properties.Resources.salir;
            this.btnsalir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnsalir.Location = new System.Drawing.Point(293, 371);
            this.btnsalir.Name = "btnsalir";
            this.btnsalir.Size = new System.Drawing.Size(48, 41);
            this.btnsalir.TabIndex = 475;
            this.btnsalir.UseVisualStyleBackColor = true;
            this.btnsalir.Click += new System.EventHandler(this.btnsalir_Click);
            // 
            // btnguardar
            // 
            this.btnguardar.BackgroundImage = global::Recepcion_PT.Properties.Resources.save;
            this.btnguardar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnguardar.Location = new System.Drawing.Point(239, 371);
            this.btnguardar.Name = "btnguardar";
            this.btnguardar.Size = new System.Drawing.Size(48, 41);
            this.btnguardar.TabIndex = 453;
            this.btnguardar.UseVisualStyleBackColor = true;
            this.btnguardar.Click += new System.EventHandler(this.btnguardar_Click);
            // 
            // Solicitud_desviacion
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(829, 430);
            this.ControlBox = false;
            this.Controls.Add(this.btnsalir);
            this.Controls.Add(this.btnguardar);
            this.Controls.Add(this.RBrechazada);
            this.Controls.Add(this.txtcom);
            this.Controls.Add(this.RBaceptada);
            this.Controls.Add(this.lbcom);
            this.Controls.Add(this.txtevidencia);
            this.Controls.Add(this.lbevi);
            this.Controls.Add(this.DTPcnte);
            this.Controls.Add(this.lbfeccnte);
            this.Controls.Add(this.txtreprecnte);
            this.Controls.Add(this.lbrepcnte);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.GBAprovcnte);
            this.Controls.Add(this.txtmot);
            this.Controls.Add(this.lbmotivo);
            this.Controls.Add(this.txtds);
            this.Controls.Add(this.lbds);
            this.Controls.Add(this.txtnc);
            this.Controls.Add(this.lbnc);
            this.Controls.Add(this.txtesp);
            this.Controls.Add(this.lbesp);
            this.Controls.Add(this.NUDCanti);
            this.Controls.Add(this.lbcant);
            this.Controls.Add(this.txtprod);
            this.Controls.Add(this.lbprod);
            this.Controls.Add(this.cbturno);
            this.Controls.Add(this.lbturno);
            this.Controls.Add(this.txtnoconformidad);
            this.Controls.Add(this.lbnumnc);
            this.Controls.Add(this.DTPFec);
            this.Controls.Add(this.lbfecha);
            this.Controls.Add(this.txtfolio);
            this.Controls.Add(this.lbfolio);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Solicitud_desviacion";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Solicitud de desviación";
            this.Load += new System.EventHandler(this.Solicitud_desviacion_Load);
            ((System.ComponentModel.ISupportInitialize)(this.NUDCanti)).EndInit();
            this.GBAprovcnte.ResumeLayout(false);
            this.GBAprovcnte.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtfolio;
        private System.Windows.Forms.Label lbfolio;
        private System.Windows.Forms.DateTimePicker DTPFec;
        private System.Windows.Forms.Label lbfecha;
        private System.Windows.Forms.Label lbnumnc;
        private System.Windows.Forms.TextBox txtnoconformidad;
        private System.Windows.Forms.ComboBox cbturno;
        private System.Windows.Forms.Label lbturno;
        private System.Windows.Forms.Label lbprod;
        private System.Windows.Forms.TextBox txtprod;
        private System.Windows.Forms.NumericUpDown NUDCanti;
        private System.Windows.Forms.Label lbcant;
        private System.Windows.Forms.TextBox txtesp;
        private System.Windows.Forms.Label lbesp;
        private System.Windows.Forms.TextBox txtnc;
        private System.Windows.Forms.Label lbnc;
        private System.Windows.Forms.Label lbds;
        private System.Windows.Forms.TextBox txtds;
        private System.Windows.Forms.Label lbmotivo;
        private System.Windows.Forms.TextBox txtmot;
        private System.Windows.Forms.RadioButton RBsi1;
        private System.Windows.Forms.RadioButton RB1no;
        private System.Windows.Forms.GroupBox GBAprovcnte;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton RBrechazado;
        private System.Windows.Forms.RadioButton RBaceptado;
        private System.Windows.Forms.Label lbrepcnte;
        private System.Windows.Forms.TextBox txtreprecnte;
        private System.Windows.Forms.DateTimePicker DTPcnte;
        private System.Windows.Forms.Label lbfeccnte;
        private System.Windows.Forms.TextBox txtevidencia;
        private System.Windows.Forms.Label lbevi;
        private System.Windows.Forms.TextBox txtcom;
        private System.Windows.Forms.Label lbcom;
        private System.Windows.Forms.RadioButton RBrechazada;
        private System.Windows.Forms.RadioButton RBaceptada;
        private System.Windows.Forms.Button btnsalir;
        private System.Windows.Forms.Button btnguardar;
    }
}