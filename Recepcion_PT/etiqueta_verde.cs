using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Net;
using System.Net.Sockets;



namespace Recepcion_PT
{

    public partial class etiqueta_verde : Form
    {
        public etiqueta_verde()
        {
            InitializeComponent();
            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);

        }



        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlCommand cmnd1 = new SqlCommand();
        SqlCommand cmnd11 = new SqlCommand();
        SqlDataReader reader1;

        SqlConnection thisConnection2 = new SqlConnection("Data Source= GABIRA1\\SQL2005;Initial Catalog=DBGAB;Connect Timeout=130;User ID=sa;  MultipleActiveResultSets=True");
        SqlCommand cmnd2 = new SqlCommand();
        //SqlDataReader reader2;

        //OleDbConnection MyConnection = new OleDbConnection(@"Provider= VFPOLEDB.1;Data Source=c:\mr_lucky\base_de_datos;Collating Sequence=general;");
        OleDbConnection MyConnection = new OleDbConnection(Utilerias.Class1.ConnectionStringFox);
        OleDbCommand cmd1 = new OleDbCommand();
        //OleDbDataReader read1;

        public DataTable dt1 = new DataTable();
        public DataTable dtoriginal = new DataTable();
        public DataTable tmp_etiquetas = new DataTable();
        public DataTable tmp_trazabilidad = new DataTable();
        public int i = 0, posicion = 0;
        public static string mes = "", anio = "", etilote = "", fecad = "", dia = "", linea = "", imprimir = "", producto = "", grado = "", fila = "", num = "", pais_origen = "", pti_famous = "";
        public static string impr = "", impr2 = "", imprverde = "", codigote = "", date = "", lote = "", fec = "", imp100x50 = "";
        public DateTime FechaEla;
        public string proveedorclave = "";
        public string existetrazabilidad = "";


        private void etiqueta_verde_Load(object sender, EventArgs e)
        {
            var file_process = "C:\\Reportes\\" + txtrecibo.Text + "_1.txt";
            if (!File.Exists(file_process))
            {
                var fileStream = File.Create(file_process);
            }
            existetrazabilidad = "";
            txtclave.Text = Recepcion_PT.cletiqueta;
            txtrecibo.Text = Recepcion_PT.recibo.ToString();
            txtviaje.Text = Recepcion_PT.num_viaje.ToString();
            txtprov.Text = Recepcion_PT.proveedor;
            txtrchtbl.Text = Recepcion_PT.rancho_nom;
            txthora.Text = Recepcion_PT.hora;
            LblHrCap.Text = Recepcion_PT.HrCaptura;
            proveedorclave = Recepcion_PT.provee;

            etilote = "";
            thisConnection.Open();
            String Cadena = "SELECT * FROM tb_det_trazabilidad WHERE recibo = '" + txtrecibo.Text.Trim() + "'";
            SqlDataAdapter da = new SqlDataAdapter(Cadena, thisConnection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Trazabilidad");
            DataTable Trazabilidad = ds.Tables["Trazabilidad"];
            //foreach (DataRow row in Trazabilidad.Rows)
            //{
            //    existetrazabilidad = "S";
            //}
            if (Trazabilidad.Rows.Count > 0)
                existetrazabilidad = "S";
            thisConnection.Close();


            FechaEla = Recepcion_PT.FechaEla;
            foreach (DataRow rw in dt1.Rows)
            {
                DGV4.Rows.Add(rw[0].ToString(), rw[1].ToString(), rw[2].ToString(), rw[3].ToString(), rw[4].ToString(), true, "", "", Convert.ToString(rw[5].ToString()), rw["items"]);
            }
            DGV4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        private void btnactu_Click(object sender, EventArgs e)
        {
            DGV5.Rows.Clear();
            for (int j = 0; j < DGV4.Rows.Count; j++)
            {
                imprimir = Convert.ToString(DGV4.Rows[j].Cells["imp"].Value).Trim();
                if (imprimir == "True")
                {
                    if ((Convert.ToString(DGV4.Rows[j].Cells["prod_clave"].Value).Trim() == "18007JI56V") ||
                        (Convert.ToString(DGV4.Rows[j].Cells["prod_clave"].Value).Trim() == "18JIBOML66") ||
                        (Convert.ToString(DGV4.Rows[j].Cells["prod_clave"].Value).Trim() == "16001TO561") ||
                        (Convert.ToString(DGV4.Rows[j].Cells["prod_clave"].Value).Trim() == "18007JI55M") ||
                        (Convert.ToString(DGV4.Rows[j].Cells["prod_clave"].Value).Trim() == "18007JVM56") ||
                        (Convert.ToString(DGV4.Rows[j].Cells["prod_clave"].Value).Trim() == "18007JVM55") ||
                        (Convert.ToString(DGV4.Rows[j].Cells["prod_clave"].Value).Trim() == "18007JVM66") ||
                        (Convert.ToString(DGV4.Rows[j].Cells["prod_clave"].Value).Trim() == "18007JI56M") ||
                        (Convert.ToString(DGV4.Rows[j].Cells["prod_clave"].Value).Trim() == "18JIML6610"))
                    {
                        if (cbgrado.SelectedIndex == -1)
                        {
                            MessageBox.Show("favor de seleccionar el tipo de grado", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            cbgrado.Visible = true;
                            lbgrado.Visible = true;
                            cbgrado.Focus();
                            fila = Convert.ToString(DGV4.Rows[j].Cells["prod_clave"].Value);
                            DGV5.Rows.Clear();
                            return;
                        }
                    }

                    int res = 0;
                    foreach (DataRow rw in dtoriginal.Select("prod_clave = '" + DGV4.Rows[j].Cells[0].Value.ToString() + "'"))
                    {
                        rw["cajas_x_tarima"] = Convert.ToString(rw["cajas_x_tarima"].ToString()).Replace(",", "");
                        res = Convert.ToInt32(rw["cajas_x_tarima"].ToString()) % Convert.ToInt32(DGV4.Rows[j].Cells[3].Value.ToString());
                    }
                    producto = Convert.ToString(DGV4.Rows[j].Cells[0].Value).Trim();
                    int total_tarimas = Convert.ToInt32(DGV4.Rows[j].Cells[2].Value);
                    for (int i = 1; i <= total_tarimas; i++)
                    {

                        thisConnection.Open();
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = "IF NOT EXISTS(SELECT ID_PALLET FROM Tb_Mstr_Pallets_SSCC WHERE pall_Folio = '" + txtrecibo.Text + "' AND Tipo_Prod = 'PTC' AND  Prod_Clave = '" + producto + "' AND  Tarima = '" + i.ToString() + "') INSERT INTO  Tb_Mstr_Pallets_SSCC(pall_Folio, pall_fecha, Tipo_Prod, Prod_Clave, Tarima) " +
                            "VALUES('" + txtrecibo.Text + "',getdate(),'PTC','" + producto + "','" + i.ToString() + "')"; ;
                        cmnd2.ExecuteReader();

                        string Cadena = "SELECT ID_PALLET FROM Tb_Mstr_Pallets_SSCC WHERE pall_Folio = '" + txtrecibo.Text + "' AND Tipo_Prod = 'PTC' AND  Prod_Clave = '" + producto + "' AND  Tarima = '" + i.ToString() + "'";
                        SqlCommand cmd = new SqlCommand(Cadena, thisConnection);
                        int ID_SSCC = Convert.ToInt32(cmd.ExecuteScalar());
                        thisConnection.Close();



                        DGV5.Rows.Add(DGV4.Rows[j].Cells[0].Value.ToString(), DGV4.Rows[j].Cells[1].Value.ToString(), DGV4.Rows[j].Cells[3].Value.ToString(), i.ToString(), DGV4.Rows[j].Cells[2].Value.ToString(), Convert.ToString(DGV4.Rows[j].Cells[4].Value), Convert.ToString(DGV4.Rows[j].Cells[6].Value), true, ID_SSCC, DGV4.Rows[j].Cells["Items"].Value);
                        if (i == total_tarimas)
                        {
                            if (res != 0)
                            {
                                DGV5.Rows[DGV5.Rows.Count - 1].Cells[2].Value = res;
                            }
                        }
                    }
                }
            }
            btnvalidareti.Enabled = true;
            DGV5.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        private void btnvalidareti_Click(object sender, EventArgs e)
        {

            int aux_tar = 0, aux_caj = 0, a = 0, aux7 = 0;
            bool ind = true;
            DataTable indicador = new DataTable();
            if (indicador.Columns.Count == 0)
            {
                DataColumn col1 = new DataColumn("c0");
                col1.DataType = System.Type.GetType("System.String");
                col1.DefaultValue = "";
                indicador.Columns.Add(col1);

                DataColumn col2 = new DataColumn("c1");
                col2.DataType = System.Type.GetType("System.String");
                col2.DefaultValue = "";
                indicador.Columns.Add(col2);

                DataColumn col3 = new DataColumn("c2");
                col3.DataType = System.Type.GetType("System.String");
                col3.DefaultValue = "";
                indicador.Columns.Add(col3);

                DataColumn col4 = new DataColumn("i_tarimas");
                col4.DataType = System.Type.GetType("System.Decimal");
                col4.DefaultValue = 0;
                indicador.Columns.Add(col4);

                DataColumn col5 = new DataColumn("trm_cap");
                col5.DataType = System.Type.GetType("System.Decimal");
                col5.DefaultValue = 0;
                indicador.Columns.Add(col5);

                DataColumn col6 = new DataColumn("i_cajas");
                col6.DataType = System.Type.GetType("System.Decimal");
                col6.DefaultValue = 0;
                indicador.Columns.Add(col6);

                DataColumn col7 = new DataColumn("cja_cap");
                col7.DataType = System.Type.GetType("System.Decimal");
                col7.DefaultValue = 0;
                indicador.Columns.Add(col7);

                DataColumn col8 = new DataColumn("sm_tar");
                col8.DataType = System.Type.GetType("System.Decimal");
                col8.DefaultValue = 0;
                indicador.Columns.Add(col8);

                DataColumn col9 = new DataColumn("sm_caj");
                col9.DataType = System.Type.GetType("System.Decimal");
                col9.DefaultValue = 0;
                indicador.Columns.Add(col9);
            }
            DGV4.Columns[2].ReadOnly = false;

            for (int x = 0; x < DGV4.Rows.Count; x++)
            {
                imprimir = Convert.ToString(DGV4.Rows[x].Cells["imp"].Value).Trim();
                if (imprimir != "True")
                {
                    //posicion = x;                    
                    foreach (DataRow row in dtoriginal.Select("prod_clave = '" + Convert.ToString(DGV4.Rows[x].Cells["prod_clave"].Value).Trim() + "'"))
                    {
                        row.Delete();
                    }
                }
                dtoriginal.AcceptChanges();
            }
            //dtoriginal.Rows.RemoveAt(x);

            for (int i = 0; i < dtoriginal.Rows.Count; i++)
            {
                aux_tar = 0;
                aux_caj = 0;
                DataRow dr = indicador.NewRow();
                dr["c1"] = dtoriginal.Rows[i][0].ToString();
                dr["c2"] = dtoriginal.Rows[i][1].ToString();
                dr["i_tarimas"] = Convert.ToDecimal(dtoriginal.Rows[i]["numero_tarimas"].ToString());
                dr["i_cajas"] = Convert.ToDecimal(dtoriginal.Rows[i]["cajas_x_tarima"].ToString());
                indicador.Rows.Add(dr);
                a = 0;

                for (int j = 0; j < DGV5.Rows.Count; j++)
                {
                    if (Convert.ToString(DGV5.Rows[j].Cells[0].Value).Trim() == dtoriginal.Rows[i][0].ToString().Trim())
                    //foreach (DataRow row in dtoriginal.Select("prod_clave = '" + Convert.ToString(DGV5.Rows[j].Cells[0].Value).Trim() + "'"))
                    {
                        aux7 = 0;
                        aux7 = Convert.ToInt32(DGV5.Rows[j].Cells[2].Value);
                        a = 1;
                        aux_tar = aux_tar + 1;
                        aux_caj = aux_caj + aux7;
                    }
                }//for dgv5
                if (a == 1)
                {
                    if (Convert.ToInt32(dtoriginal.Rows[i]["cajas_x_tarima"].ToString().Replace(",", "")) != aux_caj)
                        ind = false;

                    if (Convert.ToInt32(dtoriginal.Rows[i]["numero_tarimas"].ToString()) != aux_tar)
                        ind = false;
                    indicador.Rows[i]["trm_cap"] = aux_tar;
                    indicador.Rows[i]["cja_cap"] = aux_caj;
                }
                else
                {
                    //borra el primer registro de indicador
                    indicador.Rows.RemoveAt(0);
                }
            }//for dtoriginal
            if (ind == true)
            {
                MessageBox.Show("Proceso de validación correcto", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnimpeti.Enabled = true;
                btnimpeti.Focus();
            }
            else
            {
                MessageBox.Show("Proceso incorreco de validación de datos", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                for (int i = 0; i < indicador.Rows.Count; i++)
                {
                    indicador.Rows[i]["c0"] = "B";
                    indicador.Rows[i]["sm_tar"] = Convert.ToInt32(indicador.Rows[i]["i_tarimas"].ToString()) - Convert.ToInt32(indicador.Rows[i]["trm_cap"].ToString());
                    indicador.Rows[i]["sm_caj"] = Convert.ToInt32(indicador.Rows[i]["i_cajas"].ToString()) - Convert.ToInt32(indicador.Rows[i]["cja_cap"].ToString());
                    indicador.Rows[i]["c0"] = "M";
                    MessageBox.Show("c0:" + indicador.Rows[i]["c0"].ToString() + "\r\nc2: " + indicador.Rows[i]["c2"].ToString() + "\r\n i_tarimas :" + indicador.Rows[i]["i_tarimas"].ToString() + "\n\rtrm_cap :" + indicador.Rows[i]["trm_cap"].ToString() + "\r\ni_cajas: " + indicador.Rows[i]["i_cajas"].ToString() + "\r\ncja_cap: " + indicador.Rows[i]["cja_cap"].ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void CorrerProceso()
        {
            PaperSize ps;
            PrintDocument pd1 = new System.Drawing.Printing.PrintDocument();
            ps = new PaperSize("etiqueta_verde", 400, 600);
            pd1.PrintPage += new PrintPageEventHandler(this.printDocument2_PrintPage);
            //PrintPreviewDialog VistaPrevia1 = new PrintPreviewDialog();

            //VistaPrevia1.Document = pd1;
            pd1.DefaultPageSettings.PaperSize = ps;
            //VistaPrevia1.ShowDialog();                                
            pd1.Print();
        }


        private void btnimpeti_Click(object sender, EventArgs e)
        {
            if (tmp_etiquetas.Columns.Count == 0 && tmp_trazabilidad.Columns.Count == 0)
            {
                tmp_etiquetas.Columns.Add("etq_clave", Type.GetType("System.String"), "");
                tmp_etiquetas.Columns.Add("etq_cajas_por_tarima", Type.GetType("System.Decimal"));
                tmp_etiquetas.Columns.Add("etq_tarimas", Type.GetType("System.Decimal"));
                tmp_etiquetas.Columns.Add("etq_proveedor", Type.GetType("System.String"), "");
                tmp_etiquetas.Columns.Add("etq_rch_tabla", Type.GetType("System.String"), "");
                tmp_etiquetas.Columns.Add("etq_codigo_barras", Type.GetType("System.String"), "");
                tmp_etiquetas.Columns.Add("etq_recibo", Type.GetType("System.String"), "");
                tmp_etiquetas.Columns.Add("etq_viaje", Type.GetType("System.Int32"));
                tmp_etiquetas.Columns.Add("etq_hora", Type.GetType("System.String"), "");
                tmp_etiquetas.Columns.Add("etq_ini", Type.GetType("System.String"), "");
                tmp_etiquetas.Columns.Add("etq_final", Type.GetType("System.String"), "");
                tmp_etiquetas.Columns.Add("etq_codigo", Type.GetType("System.String"), "");
                tmp_etiquetas.Columns.Add("etq_nota", Type.GetType("System.String"), "");
                tmp_etiquetas.Columns.Add("etq_descrip", Type.GetType("System.String"), "");
                tmp_etiquetas.Columns.Add("etq_prueba1", Type.GetType("System.String"), "");
                tmp_etiquetas.Columns.Add("etq_prueba2", Type.GetType("System.String"), "");
                tmp_etiquetas.Columns.Add("etq_id_SSCC", Type.GetType("System.String"), "");
                tmp_etiquetas.Columns.Add("etq_pti_famous", Type.GetType("System.String"), "");

                tmp_trazabilidad.Columns.Add("trz_clave", Type.GetType("System.String"));
                tmp_trazabilidad.Columns.Add("trz_cajas_por_tarima", Type.GetType("System.Int32"));
                tmp_trazabilidad.Columns.Add("trz_proveedor", Type.GetType("System.String"));
                tmp_trazabilidad.Columns.Add("trz_rch_tabla", Type.GetType("System.String"));
                tmp_trazabilidad.Columns.Add("trz_codigo_barras", Type.GetType("System.String"));
                tmp_trazabilidad.Columns.Add("trz_recibo", Type.GetType("System.String"));
                tmp_trazabilidad.Columns.Add("trz_viaje", Type.GetType("System.Int32"));
                tmp_trazabilidad.Columns.Add("trz_hora", Type.GetType("System.String"));
                tmp_trazabilidad.Columns.Add("trz_ini", Type.GetType("System.String"));
                tmp_trazabilidad.Columns.Add("trz_final", Type.GetType("System.String"));
                tmp_trazabilidad.Columns.Add("trz_codigo", Type.GetType("System.String"));
                tmp_trazabilidad.Columns.Add("trz_nota", Type.GetType("System.String"));
                tmp_trazabilidad.Columns.Add("trz_descrip", Type.GetType("System.String"));
                tmp_trazabilidad.Columns.Add("trz_prueba1", Type.GetType("System.String"));
                tmp_trazabilidad.Columns.Add("trz_prueba2", Type.GetType("System.String"));
                tmp_trazabilidad.Columns.Add("feccad", Type.GetType("System.String"));
                tmp_trazabilidad.Columns.Add("imprimi", Type.GetType("System.String"));
                tmp_trazabilidad.Columns.Add("trz_id_SSCC", Type.GetType("System.String"), "");
                tmp_trazabilidad.Columns.Add("trz_pti_famous", Type.GetType("System.String"), "");
            }
            tmp_etiquetas.Rows.Clear();
            tmp_trazabilidad.Rows.Clear();
            for (int i = 0; i < DGV5.Rows.Count; i++)
            {
                imprimir = Convert.ToString(DGV5.Rows[i].Cells["imp_tar"].Value).Trim();
                if (imprimir == "True")
                {
                    string ax1 = "0" + Convert.ToString(DGV5.Rows[i].Cells[3].Value);
                    string ax2 = "0" + Convert.ToString(DGV5.Rows[i].Cells[4].Value);
                    string lon = ax1;
                    if (lon.Length == 3)
                        lon = lon.Substring(1, 2);
                    string lon2 = ax2;
                    if (lon2.Length == 3)
                        lon2 = lon2.Substring(1, 2);

                    if (Convert.ToString(DGV5.Rows[i].Cells[4].Value).Length == 3)
                    {
                        ax1 = Convert.ToString(DGV5.Rows[i].Cells[3].Value);
                        ax1 = ax1.PadLeft(3, '0');
                        lon = ax1;
                        ax2 = Convert.ToString(DGV5.Rows[i].Cells[4].Value);
                        lon2 = ax2;
                    }



                    /*string aux_cod_bar = "";
                    if (Convert.ToString(DGV5.Rows[i].Cells[1].Value).Contains("TOMATE"))
                    {
                        aux_cod_bar = string.Format("{0:000000}", Convert.ToInt32(txtrecibo.Text)) + Convert.ToString(DGV5.Rows[i].Cells[0].Value).Trim() + "  "+ Convert.ToInt16(lon).ToString().Trim();
                    }else{
                        aux_cod_bar = txtrecibo.Text + Convert.ToString(DGV5.Rows[i].Cells[0].Value).Trim() + lon + lon2;   
                    }*/

                    string aux_cod_bar = "";

                    if (proveedorclave == "03" && (Convert.ToString(DGV5.Rows[i].Cells[1].Value).Contains("TOMATE") || Convert.ToString(DGV5.Rows[i].Cells[1].Value).Contains("ELOTE")))
                    {
                        aux_cod_bar = txtrecibo.Text.PadLeft(6, '0') + Convert.ToString(DGV5.Rows[i].Cells[0].Value).Trim() + lon.PadLeft(3);
                    }
                    else
                    {
                        aux_cod_bar = txtrecibo.Text + Convert.ToString(DGV5.Rows[i].Cells[0].Value).Trim() + lon + lon2;
                    }
                    if (DGV5.Rows[i].Cells["Item"].Value.ToString().Trim().Length > 0)
                        pti_famous = txtrecibo.Text.PadLeft(6, '0') + Convert.ToString(DGV5.Rows[i].Cells["item"].Value).Trim().Substring(0, 4) + lon.PadLeft(2);

                    DataRow rwtrz = tmp_trazabilidad.NewRow();
                    rwtrz["trz_clave"] = txtclave.Text;
                    rwtrz["trz_cajas_por_tarima"] = Convert.ToInt32(DGV5.Rows[i].Cells[2].Value);
                    rwtrz["trz_proveedor"] = txtprov.Text;
                    rwtrz["trz_rch_tabla"] = txtrchtbl.Text;
                    rwtrz["trz_codigo_barras"] = aux_cod_bar;
                    rwtrz["trz_recibo"] = txtrecibo.Text;
                    rwtrz["trz_viaje"] = Convert.ToInt32(txtviaje.Text);
                    rwtrz["trz_hora"] = LblHrCap.Text;  //txthora.Text;
                    rwtrz["trz_ini"] = lon;
                    rwtrz["trz_final"] = lon2;
                    rwtrz["trz_codigo"] = Convert.ToString(DGV5.Rows[i].Cells[0].Value);
                    rwtrz["trz_descrip"] = Convert.ToString(DGV5.Rows[i].Cells[1].Value);
                    rwtrz["trz_nota"] = Convert.ToString(DGV5.Rows[i].Cells[5].Value);
                    rwtrz["trz_prueba1"] = Convert.ToString(DGV5.Rows[i].Cells[4].Value);
                    rwtrz["trz_prueba2"] = "123ABC5567890";
                    rwtrz["feccad"] = Convert.ToString(DGV5.Rows[i].Cells[6].Value);
                    rwtrz["imprimi"] = "S";
                    rwtrz["trz_id_SSCC"] = Convert.ToString(DGV5.Rows[i].Cells[8].Value);
                    rwtrz["trz_pti_famous"] = pti_famous;
                    tmp_trazabilidad.Rows.Add(rwtrz);

                    DataRow rweti = tmp_etiquetas.NewRow();
                    rweti["etq_clave"] = txtclave.Text;
                    rweti["etq_cajas_por_tarima"] = Convert.ToDecimal(DGV5.Rows[i].Cells[2].Value);
                    rweti["etq_proveedor"] = txtprov.Text;
                    rweti["etq_rch_tabla"] = txtrchtbl.Text;
                    //rweti["etq_codigo_barras"] = "*" + aux_cod_bar + "*";
                    rweti["etq_codigo_barras"] = aux_cod_bar;
                    rweti["etq_recibo"] = txtrecibo.Text;
                    rweti["etq_viaje"] = Convert.ToInt32(txtviaje.Text);
                    rweti["etq_hora"] = LblHrCap.Text;  //txthora.Text;
                    rweti["etq_ini"] = lon;
                    rweti["etq_final"] = lon2;
                    rweti["etq_codigo"] = Convert.ToString(DGV5.Rows[i].Cells[0].Value);
                    rweti["etq_descrip"] = Convert.ToString(DGV5.Rows[i].Cells[1].Value);
                    rweti["etq_nota"] = Convert.ToString(DGV5.Rows[i].Cells[5].Value);
                    rweti["etq_prueba1"] = "*123ABC5567890*";
                    rweti["etq_prueba2"] = "123ABC5567890";
                    //etq_id_SSCC
                    rweti["etq_id_SSCC"] = Convert.ToString(DGV5.Rows[i].Cells[8].Value);
                    rweti["etq_pti_famous"] = pti_famous;
                    tmp_etiquetas.Rows.Add(rweti);
                }
            }

            #region EtiVerde
            if (MessageBox.Show("¿Desea imprimir las etiquetas verdes?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                //if (i <= tmp_etiquetas.Rows.Count - 1)
                imprverde = "";

                string sscc = "";
                for (int i = 0; i <= tmp_etiquetas.Rows.Count - 1; i++)
                {
                    thisConnection.Open();
                    string Cadena = "SELECT prod_NoEti FROM tb_cat_producto WHERE Prod_Clave = '" + tmp_etiquetas.Rows[i][11].ToString().Trim() + "'";
                    SqlCommand cmd = new SqlCommand(Cadena, thisConnection);
                    string prod_NoEti = cmd.ExecuteScalar().ToString().Trim();
                    thisConnection.Close();

                    sscc = "0000796631" + tmp_etiquetas.Rows[i][16].ToString().Trim().PadLeft(9, '0');
                    imprverde += "^XA\n";
                    imprverde += "~SD20"; //~SD20
                    imprverde += "^PW832";
                    imprverde += "^FO100,25,^A0N,48,35,^FDCLAVE: " + txtclave.Text + "^FS\n";
                    imprverde += "^FO564,10,^GB150,100,7^FS\n";
                    imprverde += "^FO580,50,^A2N,28,28^FD" + Convert.ToString(tmp_etiquetas.Rows[i][1].ToString().Trim()) + "^FS\n";
                    imprverde += "^FO30,100,^A0N,48,35^FDSource^FS\n";
                    imprverde += "^FO30,150,^A0N,48,35^FD" + txtprov.Text + "^FS\n";
                    imprverde += "^FO30,230,^A0N,48,35^FDField^FS\n";
                    imprverde += "^FO030,280^A0N,48,35,^FD" + txtrchtbl.Text + "^FS\n";
                    imprverde += "^FO30,360,^A0N,48,35^FDProduct^FS\n";
                    if (Convert.ToString(tmp_etiquetas.Rows[i]["etq_descrip"].ToString()).Length >= 25)
                    {
                        imprverde += "^FO30,410,^A0N,48,35^FD" + Convert.ToString(tmp_etiquetas.Rows[i]["etq_descrip"].ToString()).Substring(0, 25) + "^FS\n";
                        imprverde += "^FO30,450,^A0N,48,35^FD" + Convert.ToString(tmp_etiquetas.Rows[i]["etq_descrip"].ToString()).Substring(25) + "^FS\n";
                    }
                    else
                        imprverde += "^FO30,410,^A0N,48,35^FD" + Convert.ToString(tmp_etiquetas.Rows[i]["etq_descrip"].ToString()).Trim() + "^FS\n";

                    imprverde += "^FO250,500,^A0N,48,35,^FDTrazability Lot #: " + Convert.ToString(tmp_etiquetas.Rows[i]["etq_recibo"].ToString()).Trim() + "^FS";
                    imprverde += "^FO250,550,^A0N,48,35,^FDLoad #: " + Convert.ToString(tmp_etiquetas.Rows[i]["etq_viaje"].ToString()).Trim() + "^FS";
                    imprverde += "^FO250,600,^A0N,48,35,^FDTime in " + Convert.ToString(tmp_etiquetas.Rows[i]["etq_hora"].ToString()).Trim() + "^FS";
                    imprverde += "^FO180,650,^GB500,100,7^FS\n";

                    imprverde += "^FO15,950^ABN,25,17^FDSSCC^FS\n"; //O15,1000
                    imprverde += "^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>8" + sscc + "^FS\n"; // CODIGO SSCC GS1 128 TARIMA UNICA

                    imprverde += "^FO700,1050^A0N,20,18^FD" + "F-200-37" + "^FS\n"; //FO650,980
                    imprverde += "^FO700,1075^A0N,20,18^FD" + "REV: 04" + "^FS\n"; //O650,1005

                    imprverde += "^FO196,690,^A2N,33,23^FD" + Convert.ToString(tmp_etiquetas.Rows[i]["etq_ini"].ToString().Trim()) + " of " + Convert.ToString(tmp_etiquetas.Rows[i]["etq_final"].ToString().Trim()) + " Pallets^FS\n";
                    //imprverde += "^FO120,780,^BY2,^BCN,120,N,N,N^FD" + Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString().Trim()) + "^FS\n";
                    // esta es la ubicacion actual del codigo QR RCC 11 jun 2024 
                    //imprverde += "^FO340,720^BQN,2,7^FDLA," + Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString().Trim()) + "^FS\n"; // CODIGO QR
                    // esta es la nueva propuesta de ubicacion del codigo QR RCC 11 de junio 2024
                    imprverde += "^FO590,180^BQN,2,7^FDLA," + Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString().Trim()) + "^FS\n"; // CODIGO QR
                    //imprverde += "^FO140,910,^A2N,15,15,^FD" + Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString().Trim()) + "^FS\n";
                    imprverde += "^FO30,880,^A0N,40,30^FD" + DateTime.Now.ToString("HH:mm:ss") + "^FS\n";
                    imprverde += "^FO580,880,^A0N,40,30^FD" + Convert.ToString(tmp_etiquetas.Rows[i]["etq_nota"].ToString().Trim()) + "^FS\n";

                    if (tmp_etiquetas.Rows[i]["etq_pti_famous"].ToString().Trim().Length > 0)
                    {
                        imprverde += "^FO120,1060,^BY3,^BCN,90,N,N,N^FD" + Convert.ToString(tmp_etiquetas.Rows[i]["etq_pti_famous"].ToString().Trim()) + "^FS\n";
                        imprverde += "^FO230,1160,^A2N,20,20,^FD" + Convert.ToString(tmp_etiquetas.Rows[i]["etq_pti_famous"].ToString().Trim()) + "^FS\n";
                    }
                    else
                    {
                        imprverde += "^FO120,1060,^BY2,^BCN,120,N,N,N^FD" + Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString().Trim()) + "^FS\n";
                        imprverde += "^FO140,1190,^A2N,35,20,^FD" + Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString().Trim()) + "^FS\n";
                    }
                    if (prod_NoEti != "")
                    {
                        imprverde += "^PQ2";
                    }
                    imprverde += "^XZ\n";

                    //codigote += "^XA\n";
                    ////codigote += "^FO210,150,^BY4,^BCR,500,N,N,N^FD" + Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString().Trim()) + "^FS\n";
                    //codigote += "^FO200,100^BQN,4,4^FDLA," + Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString().Trim()) + "^FS\n";
                    //codigote += "^XZ\n";
                }
                //RawPrinterHelper.SendStringToPrinter("etiqueta_verde", imprverde);
                if (System.Environment.MachineName.ToString() == "RICARDOGAB")
                    RawPrinterHelper.SendStringToPrinter("etiqueta_blanca", imprverde);
                RawPrinterHelper.SendStringToPrinter("etiqueta_verde", imprverde);


                /*IPAddress addr = IPAddress.Parse("192.168.123.231");
                EndPoint ep = new IPEndPoint(addr, 9100);
                Socket sock = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                sock.Connect(ep);
                NetworkStream ns = new NetworkStream(sock);
                String transferType = "^MTT\n";// Sets the type to thermal transfer
                String ZPLString = "^LH5,5\n" + transferType +
                                   "^BY2" + "^MNM\n" + imprverde;
                byte[] toSend = Encoding.ASCII.GetBytes(ZPLString);
                ns.Write(toSend, 0, toSend.Length);
                */

                //RawPrinterHelper.SendStringToPrinter("ZDesigner ZT410-203dpi ZPL", imprverde); 
                //RawPrinterHelper.SendStringToPrinter("ZDesigner ZT410-203dpi ZPL", imprverde);
                //RawPrinterHelper.SendStringToPrinter("etiqueta_verde", codigote);
                //RawPrinterHelper.SendStringToPrinter("ZDesigner Z4Mplus 203DPI (Copiar 1)", imprverde);
                //ThreadStart delegado = new ThreadStart(CorrerProceso);
                //Thread hilo = new Thread(delegado);
                //hilo.Start();
                //Thread.Sleep(1000);
                //PaperSize ps;
                //PrintDocument pd1 = new System.Drawing.Printing.PrintDocument();
                //ps = new PaperSize("etiqueta_verde", 400, 600);
                //pd1.PrintPage += new PrintPageEventHandler(this.printDocument2_PrintPage);
                ////PrintPreviewDialog VistaPrevia1 = new PrintPreviewDialog();

                ////VistaPrevia1.Document = pd1;
                //pd1.DefaultPageSettings.PaperSize = ps;
                ////VistaPrevia1.ShowDialog();                                
                //pd1.Print();
            }
            #endregion EtiVerde

            etiqueta_blanca eti_blanca = new etiqueta_blanca();
            eti_blanca.recibo = txtrecibo.Text;
            if (!(eti_blanca.trazabilidad.Columns.Contains("pti_clave")))
            {
                eti_blanca.trazabilidad.Columns.Add("pti_clave");
                eti_blanca.trazabilidad.Columns.Add("prod_nom_ingles");
                eti_blanca.trazabilidad.Columns.Add("gtin_clave");
                eti_blanca.trazabilidad.Columns.Add("prod_clave");
                eti_blanca.trazabilidad.Columns.Add("prod_nombre");
                eti_blanca.trazabilidad.Columns.Add("etiqueta");
                eti_blanca.trazabilidad.Columns.Add("fecha_cad");
                eti_blanca.trazabilidad.Columns.Add("tarima");
                eti_blanca.trazabilidad.Columns.Add("recibo");
                eti_blanca.trazabilidad.Columns.Add("pais");
            }
            string aux1 = "", aux2 = "", aux3 = "", aux3_texto = "", aux5 = "", mfeccad = "", clave_gab = "";
            string auxc = "", auxd = "", aux4 = "", auxnum1 = "", auxnum2 = "", auxd_1 = "", aux4_1 = "", auxblanca2paso = "";
            string nutar1 = "", totar1 = "", nutar2 = "", totar2 = "";
            decimal auxn = 0, tot_cajas = 0;
            int y = 0, lon_cad1 = 0, lon_cad2 = 0;
            //string eti_grande = "";
            impr = string.Empty;
            impr2 = string.Empty;
            imp100x50 = "";
            //if (MessageBox.Show("¿Desea imrprimir en etiqueta grande?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)            
            //    eti_grande = "S";

            //179228 
            //using (StreamWriter sw = new StreamWriter((@"C:\Reportes\" + txtrecibo.Text + "_1.txt")))
            //{
            thisConnection.Open();
            //for (int x = 0; x < DGV4.Rows.Count; x++)
            string CveProd = "", ConLechuga = "N"; ;
            for (int x = 0; x < DGV5.Rows.Count; x++)
            {
                if (CveProd != DGV5.Rows[x].Cells["clave_producto"].Value.ToString())
                {
                    String Cadena = "SELECT compp_clave FROM tb_mstr_comp_prod WHERE prod_clave = '" + DGV5.Rows[x].Cells["clave_producto"].Value.ToString() + "'";
                    SqlDataAdapter da = new SqlDataAdapter(Cadena, thisConnection);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "MatPri");
                    DataTable MatPri = ds.Tables["MatPri"];
                    foreach (DataRow row in MatPri.Rows)
                    {
                        if (row["compp_clave"].ToString().Trim() == "16004" || row["compp_clave"].ToString().Trim() == "09001" || row["compp_clave"].ToString().Trim() == "05004" || row["compp_clave"].ToString().Trim() == "05015" || row["compp_clave"].ToString().Trim() == "05005")
                        {
                            ConLechuga = "S";
                            break;
                        }
                    }
                    CveProd = DGV5.Rows[x].Cells["clave_producto"].Value.ToString();
                }
                imprimir = Convert.ToString(DGV5.Rows[x].Cells["imp_tar"].Value).Trim();

                if (imprimir.Trim() == "True")
                {
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select a.pti_clave, a.prod_nom_ingles, a.gtin_clave, a.prod_clave, a.prod_nombre, a.fecha_cad, a.etiqueta, b.prod_paisorigen, " +
                                        "a.lote, a.pti_fecha, b.prod_codebar from tb_det_trazabilidad a, tb_cat_producto b where " +
                                        "a.recibo = '" + txtrecibo.Text + "' and a.prod_clave = '" + Convert.ToString(DGV5.Rows[x].Cells["clave_producto"].Value).Trim() + "' " +
                                        "and a.tarima = '" + Convert.ToString(DGV5.Rows[x].Cells["tar"].Value) + "' and b.prod_clave = a.prod_clave " +
                                        "order by  a.recibo, a.prod_clave, a.tarima";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        DataRow row = eti_blanca.trazabilidad.NewRow();
                        row["pti_clave"] = reader1.GetValue(0).ToString();
                        row["prod_nom_ingles"] = reader1.GetValue(1).ToString();
                        row["gtin_clave"] = reader1.GetValue(2).ToString();
                        row["prod_clave"] = reader1.GetValue(3).ToString();
                        row["prod_nombre"] = reader1.GetValue(4).ToString();
                        row["fecha_cad"] = reader1.GetValue(5).ToString();
                        row["etiqueta"] = reader1.GetValue(6).ToString();
                        row["tarima"] = Convert.ToString(DGV5.Rows[x].Cells["tar"].Value);
                        row["recibo"] = txtrecibo.Text;
                        row["pais"] = reader1.GetValue(7).ToString().Trim();
                        pais_origen = reader1.GetValue(7).ToString().Trim();
                        lote = reader1.GetValue(8).ToString().Trim();
                        eti_blanca.trazabilidad.Rows.Add(row);
                        aux1 = reader1.GetValue(0).ToString().Trim(); //pti_clave
                        aux2 = reader1.GetValue(4).ToString();//prod_nombre
                        aux3 = reader1.GetValue(2).ToString().Trim();//gtin_clave
                        clave_gab = reader1.GetValue(3).ToString().Trim();//prod_clave
                        aux5 = reader1.GetValue(1).ToString();//prod_nombre_ingles
                        aux5 = aux5.Replace("''", "");
                        aux5 = aux5.Replace("'", "");
                        String CodeBar = reader1.GetValue(10).ToString().Trim();
                        fec = reader1.GetValue(9).ToString().Trim();
                        //mfeccad = reader1.GetValue(5).ToString().Trim();//fecha caducidad
                        auxn = Convert.ToDecimal(reader1.GetValue(6).ToString().Trim());//etiqueta
                        fecad = reader1.GetValue(5).ToString().Trim();//fecha caducidad
                                                                      //if ((fecad.Trim() != ""))
                        string lin = clave_gab.Substring(0, 2);//saber la linea
                        string CodigoBar = "";
                        if (reader1["prod_codebar"].ToString().Trim().Length > 0)
                            CodigoBar = reader1["prod_codebar"].ToString().Trim().Substring(0, reader1["prod_codebar"].ToString().Trim().Length - 1);
                        if (((clave_gab == "05003ML3P") || (clave_gab == "05005ML2P") || (lin == "19") || clave_gab.Trim() == "03001ML09" || clave_gab.Trim() == "03002ML12") && fecad.Trim().Length > 0)
                        {
                            if (lin == "19")
                                etilote = "LOTE: " + TrarItem(clave_gab).Trim() + " O2";
                            else
                                etilote = "LOTE: " + txtrecibo.Text;
                            anio = new string(fecad.Reverse().Take(2).Reverse().ToArray());
                            dia = fecad.Substring(0, 2);
                            fecad = fecad.Substring(3, 2);
                            obtenerNombreMesNumero(Convert.ToInt32(fecad));
                            mfeccad = "CAD " + dia + mes.ToUpper() + anio;
                        }
                        if (aux3 == "")
                        {
                            aux3 = clave_gab;
                            aux3_texto = "(01)" + aux3;
                        }
                        else
                            aux3_texto = "(01)" + aux3;

                        tot_cajas = tot_cajas + auxn;

                        for (int i = 1; i <= auxn; i++)
                        {
                            //decimal resto = (auxn % 2);
                            //if (resto == 0)

                            auxc = i.ToString();
                            if (auxc.Length == 1)
                                auxd = "00" + auxc;
                            if (auxc.Length == 2)
                                auxd = "0" + auxc;

                            aux4 = aux1 + i.ToString().Trim().PadLeft(3, '0'); //auxd;
                            auxnum1 = i.ToString();

                            y = i + 1;
                            auxnum2 = y.ToString();
                            auxc = y.ToString();
                            auxd = auxc;
                            auxd_1 = auxd;
                            if (auxd_1.Length == 1)
                                auxd_1 = "00" + auxd;
                            if (auxc.Length == 2)
                                auxd_1 = "0" + auxd;

                            aux4_1 = aux1 + y.ToString().Trim().PadLeft(3, '0'); //auxd_1;
                            lon_cad1 = aux4.Length + 1;
                            lon_cad2 = aux4_1.Length + 1;

                            /*nutar1 = aux4.Substring(lon_cad1 - 8, 2);
                            totar1 = aux4.Substring(lon_cad1 - 6, 2);
                            nutar2 = aux4_1.Substring(lon_cad1 - 8, 2);
                            totar2 = aux4_1.Substring(lon_cad1 - 6, 2);*/
                            nutar1 = Convert.ToInt32(DGV5.Rows[x].Cells["tar"].Value).ToString();
                            totar1 = Convert.ToInt32(DGV5.Rows[x].Cells["tarimas_total"].Value).ToString();

                            nutar2 = Convert.ToInt32(DGV5.Rows[x].Cells["tar"].Value).ToString();
                            totar2 = Convert.ToInt32(DGV5.Rows[x].Cells["tarimas_total"].Value).ToString();

                            string Church = "N";
                            if (aux2.Contains("CHURCH"))
                            {
                                Church = "S";
                            }

                            if (pais_origen.Trim() != "")
                            {
                                #region etiqueta nueva
                                //sw.WriteLine("^XA");
                                //sw.WriteLine("^FO645,30^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS");
                                //sw.WriteLine("^FO15,25^A2N,23,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS");
                                //sw.WriteLine("^FO15,75^A0N,48,24^FD" + aux2 + "^FS");
                                //sw.WriteLine("^FO15,140^A0N,48,24^FD" + aux5 + "^FS");
                                //sw.WriteLine("^FO15,213^A0N,25,25^FDPRODUCT OF " + pais_origen + "^FS");
                                //sw.WriteLine("^FO730,235^A0B,35,28,^FD" + mfeccad + "^FS");
                                //sw.WriteLine("^FO740,193^A0N,21,15^FDC:" + i.ToString() + "^FS");
                                //sw.WriteLine("^^FO80,270,^BY2,^BCN,90,N,N,N^FD01" + aux3 + "10" + txtrecibo.Text + "^FS");
                                //sw.WriteLine("^^FO120,370^A2N,17,17,^FD" + aux3_texto + "(10)" + txtrecibo.Text + "^FS");
                                //sw.WriteLine("^XZ");

                                //impr2 = impr2 + "^XA\n";
                                //impr2 = impr2 + "^FWB,0\n";
                                //impr2 = impr2 + "^FO15,250^A20,23,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";
                                //impr2 = impr2 + "^FO50,30^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                //impr2 = impr2 + "^FO65,320^A0,48,24^FD" + aux2 + "^FS\n";
                                //impr2 = impr2 + "^FO115,320^A0,48,24^FD" + aux5 + "^FS\n";
                                //impr2 = impr2 + "^FO193,80^A0,21,15^FDC:" + i.ToString() + "^FS\n";
                                //impr2 = impr2 + "^FO170,600^A0,25,25^FDPRODUCT OF " + pais_origen + "^FS\n";
                                //impr2 = impr2 + "^FO170,300^A0,25,25^FDHEB US#1^FS\n";
                                ////impr2 = impr2 + "^FO730,235^A0B,35,28,^FD" + mfeccad + "^FS\n";                                    
                                //impr2 = impr2 + "^FO210,150,^BY2,3,^BCR,90,N,N,N^FD01" + aux3 + "^FS\n";
                                ////impr2 = impr2 + "^FO310,300^A20,17,17,^FD" + aux3_texto + "(10)" + txtrecibo.Text + "^FS\n";
                                //impr2 = impr2 + "^FO310,300^A20,17,17,^FD" + aux3_texto + "^FS\n";
                                //impr2 = impr2 + "^FO350,280^A0,25,25^FDPACKED BY: COMERCIALIZADORA GAB, S.A. DE C.V.^FS\n";
                                //impr2 = impr2 + "^FO370,375^A0,25,25^FDCARRETERA PANAMERICANA KM. 291 - 1^FS\n";
                                //impr2 = impr2 + "^FO390,310^A0,25,25^FDCOL. LA FORTALEZA CORTAZAR, GTO. C.P. 38300^FS\n";
                                //impr2 = impr2 + "^FO410,460^A0,25,25^FDMEXICO R.F.C. CGA-960614-2C5^FS\n";
                                //impr2 = impr2 + "^XZ\n";    


                                //COMENTADO 08/12/2016
                                //impr2 += "^XA\n";
                                //impr2 += "^FO05,55^A2N,23,17,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";
                                //impr2 += "^FO15,100^A0N,21,15^FDC:" + i.ToString() + "^FS\n";
                                //impr2 += "^FO200,100^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                //impr2 += "^FO15,250^A0N,48,24^FD" + aux2 + "^FS\n";
                                //impr2 += "^FO15,300^A0N,48,24^FD" + aux5 + "^FS\n";
                                //impr2 += "^FO15,350^A0N,25,25^FDPRODUCT OF " + pais_origen + "^FS\n";
                                //impr2 += "^FO250,350^A0N,25,25^FDHEB US#1^FS\n";
                                //impr2 += "^FO50,380,^BY2,3,^BCN,90,N,N,N^FD01" + aux3 + "^FS\n";
                                //impr2 += "^FO150,480^A0N,25,25^FD" + aux3_texto + "^FS\n";
                                //impr2 += "^FO15,520^A0N,25,25^FDPACKED BY: COMERCIALIZADORA GAB, S.A. DE C.V.^FS\n";
                                //impr2 += "^FO15,550^A0N,25,25^FDCARRETERA PANAMERICANA KM. 291 - 1^FS\n";
                                //impr2 += "^FO15,580^A0N,25,25^FDCOL. LA FORTALEZA CORTAZAR, GTO. C.P. 38300^FS\n";
                                //impr2 += "^FO15,610^A0N,25,25^FDMEXICO R.F.C. CGA-960614-2C5^FS\n";
                                //impr2 += "^XZ\n";

                                impr2 += "^XA\n";
                                impr2 += "~SD12\n"; //~SD15
                                impr2 += "^PW832\n";
                                impr2 += "^FO20,15^A2N,23,17,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";
                                impr2 += "^FO622,200^A0N,21,15^FDC:" + i.ToString() + "^FS\n";
                                impr2 += "^FO622,30^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                impr2 += "^FO20,50^A0N,48,24^FD" + aux2 + "^FS\n";
                                impr2 += "^FO20,100^A0N,48,24^FD" + aux5 + "^FS\n";
                                impr2 += "^FO20,150^A0N,25,25^FDPRODUCT OF " + pais_origen + "^FS\n";
                                //impr2 += "^FO300,150^A0N,25,25^FDHEB US#1^FS\n";//AQUI
                                impr2 += "^FO100,180,^BY2,3,^BCN,90,N,N,N^FD01" + aux3 + "^FS\n";
                                impr2 += "^FO150,275^A0N,25,25^FD" + aux3_texto + "^FS\n";
                                impr2 += "^FO20,305^A0N,20,20^FDPACKED BY: COMERCIALIZADORA GAB, S.A. DE C.V.^FS\n";
                                impr2 += "^FO20,325^A0N,20,20^FDCARRETERA PANAMERICANA KM. 291 - 1^FS\n";
                                impr2 += "^FO20,345^A0N,20,20^FDCOL. LA FORTALEZA CORTAZAR, GTO. C.P. 38300^FS\n";
                                impr2 += "^FO20,365^A0N,20,20^FDMEXICO R.F.C. CGA-960614-2C5^FS\n";
                                impr2 += "^XZ\n";
                                #endregion
                            }
                            else
                            {
                                #region etiqueta 4" x 2"  
                                if ((ConLechuga == "S" || Church == "S" || clave_gab.Trim() == "05006MLNA2") && clave_gab.Trim() != "05003ML3P")
                                {
                                    if (clave_gab.Trim() == "05006MLNA2")
                                    {
                                        //for (int ii = 1; ii <= auxn; ii++)
                                        //{
                                        imp100x50 += "^XA\n";
                                        imp100x50 += "~SD12\n"; //~SD15
                                        imp100x50 += "^PW832\n";
                                        //imp100x50 += "^FO635,30^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                        //imp100x50 += "^FO15,25^A2N,23,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";
                                        //imp100x50 += "^FO15,75^A0N,48,32^FD" + aux5 + "^FS\n";
                                        //imp100x50 += "^FO15,135^A0N,32,24^FD" + "Lot: " + txtrecibo.Text + "    Pack Date: " + Convert.ToDateTime(fec).ToString("MMM dd", CultureInfo.CreateSpecificCulture("en-US")).ToUpper().Replace(".", "") + "^FS\n";
                                        //imp100x50 += "^FO15,235^A0N,25,25^FDPRODUCT OF MEXICO^FS\n";
                                        //imp100x50 += "^FO650,191^A0N,25,25,^FD" + mfeccad + "^FS\n";
                                        //imp100x50 += "^FO740,193^A0N,21,15^FDC:" + i.ToString() + "^FS\n";
                                        //imp100x50 += "^FO400,140,^BY2,^BUN,50^FD" + CodeBar + "^FS\n";
                                        //imp100x50 += "^FO60,275,^BY3,^BCN,80,N,N,N,D^FD01" + aux3 + "13" + Convert.ToDateTime(fec).ToString("yyMMdd") + "10" + txtrecibo.Text + "^FS\n";
                                        //imp100x50 += "^FO140,368^A2N,14,12,^FD" + aux3_texto + "(13)" + Convert.ToDateTime(fec).ToString("yyMMdd") + "(10)" + txtrecibo.Text + "^FS\n";
                                        //imp100x50 += "^XZ\n";
                                        imp100x50 += "^FO60,15,^BY3,^BCN,80,N,N,N,D^FD01" + aux3 + "13" + Convert.ToDateTime(fec).ToString("yyMMdd") + "10" + txtrecibo.Text + "^FS\n";
                                        imp100x50 += "^FO140,108^A2N,14,12,^FD" + aux3_texto + "(13)" + Convert.ToDateTime(fec).ToString("yyMMdd") + "(10)" + txtrecibo.Text + "^FS\n";
                                        imp100x50 += "^FO15,145^A0N,48,32^FD" + aux5 + "^FS\n";
                                        imp100x50 += "^FO580,240^GB195,80,3^FS\n";
                                        imp100x50 += "^FO590,200^A0N,40,30,^FDPack Date: ^FS\n";
                                        imp100x50 += "^FO590,247^A0N,80,60,^FD" + Convert.ToDateTime(fec).ToString("MMM dd", CultureInfo.CreateSpecificCulture("en-US")).ToUpper().Replace(".", "") + "^FS\n";
                                        imp100x50 += "^FO15,225^A0N,30,30^FDPRODUCT OF MEXICO^FS\n";
                                        imp100x50 += "^FO590,340^A0N,35,28^FD" + "Lot: " + txtrecibo.Text + "^FS\n";
                                        imp100x50 += "^FO50,260^BQN,3,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                        imp100x50 += "^FO50,380^A0N,22,18^FDTar: " + nutar1 + "  C:" + i.ToString() + "^FS\n";
                                        imp100x50 += "^FO300,300,^BY2,^BUN,50^FD" + CodeBar + "^FS\n";
                                        imp100x50 += "^XZ\n";
                                        //}
                                        //continue;
                                    }
                                    else
                                    {
                                        if (clave_gab.Trim() != "05003ML3P" && clave_gab.Trim() != "03001ML09" && clave_gab.Trim() != "03002ML12")
                                        //23 may 23 RCC if (clave_gab.Trim() != "05003ML3P")
                                        {
                                            imp100x50 += "^XA\n";
                                            imp100x50 += "~SD12\n"; //~SD15
                                            imp100x50 += "^PW832\n";
                                            imp100x50 += "^FO635,30^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                            imp100x50 += "^FO15,25^A2N,23,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";
                                            imp100x50 += "^FO15,75^A0N,48,24^FD" + aux2 + "^FS\n";
                                            imp100x50 += "^FO15,135^A0N,48,24^FD" + aux5 + "^FS\n";
                                            imp100x50 += "^FO15,191^A0N,25,25^FDPRODUCT OF MEXICO^FS\n";
                                            imp100x50 += "^FO650,191^A0N,25,25,^FD" + mfeccad + "^FS\n";
                                            imp100x50 += "^FO740,193^A0N,21,15^FDC:" + i.ToString() + "^FS\n";
                                            imp100x50 += "^FO80,220,^BY2,^BCN,90,N,N,N^FD01" + aux3 + "10" + txtrecibo.Text + "^FS\n";
                                            imp100x50 += "^FO120,328^A2N,17,17,^FD" + aux3_texto + "(10)" + txtrecibo.Text + "^FS\n";
                                            if (ConLechuga == "S")
                                                imp100x50 += "^FO20,350^A0N,20,20^FDRomaine Grown in / Origen :  Guanajuato - " + txtrchtbl.Text.Trim() + " ^FS\n";
                                            imp100x50 += "^FO20,370^A0N,20,20^FDHarvested On / Cosechado : " + FechaEla.ToString("dd MMM yy", CultureInfo.CreateSpecificCulture("en-US")).ToUpper() + "^FS\n";
                                            imp100x50 += "^XZ\n";
                                        }
                                    }
                                }
                                if (clave_gab.Trim() == "05003ML3P" || clave_gab.Trim() == "03001ML09" || clave_gab.Trim() == "03002ML12") // ETIQUETA DE COSTCO LECHUGA MR.LUCKY 3 PZAS.
                                                                                                                                           //if (clave_gab.Trim() == "05003ML3P") // ETIQUETA DE COSTCO LECHUGA MR.LUCKY 3 PZAS.
                                {
                                    imp100x50 += "^XA\n";
                                    imp100x50 += "~SD12\n"; //~SD15
                                    imp100x50 += "^PW832\n";
                                    if (clave_gab.Trim().Length > 25)
                                    {
                                        imp100x50 += "^FO10,20^A0N,80,60^FD" + aux2 + "^FS\n";
                                        imp100x50 += "^FO10,95^A0N,80,60^FD" + aux5 + "^FS\n";
                                    }
                                    else
                                        imp100x50 += "^FO10,20^A0N,80,60^FD" + aux2 + "^FS\n";
                                    imp100x50 += "^FO55,160^BQN,4,4^FDLA," + "http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + txtrecibo.Text + clave_gab + nutar1.Trim().PadLeft(2, '0') + totar1.Trim().PadLeft(2, '0') + i.ToString().Trim().PadLeft(3, '0') + "^FS\n"; // GENERA EL CODIGO QR
                                    if (CodigoBar.Trim().Length > 0)
                                        imp100x50 += "^FO540,330,^BY2,^BUN,50,Y,N,Y^FD" + CodigoBar + "^FS\n";
                                    imp100x50 += "^FO300,180^A0N,30,30^FDLOTE:^FS\n";
                                    imp100x50 += "^FO300,210^A0N,40,40^FD" + txtrecibo.Text.Trim() + "^FS\n";
                                    imp100x50 += "^FO300,250^A0N,20,20^FD" + "T: " + nutar1.Trim().PadLeft(3, '0') + " / " + i.ToString().Trim().PadLeft(3, '0') + "^FS\n";
                                    //imp100x50 +="^FO500,170^A0N,60,60^FDCons. Pref.^FS\n";
                                    //imp100x50 +="^FO500,230^GB280,90,90^FS\n";
                                    string tmp = "", FCad = "";
                                    if (mfeccad.Trim().Length > 0)
                                    //if (fecad.Trim().Length > 0)
                                    {
                                        tmp = mfeccad.Substring(mfeccad.Trim().Length - 7);
                                        FCad = tmp.Substring(2, 3) + " " + tmp.Substring(0, 2);
                                        imp100x50 += "^FO500,170^A0N,60,60^FDCons. Pref.^FS\n";
                                        imp100x50 += "^FO500,230^GB280,90,90^FS\n";
                                        imp100x50 += "^FO520,240^A0N,90,80^FR^FD" + FCad + "^FS\n";
                                    }
                                    //imp100x50 +="^FO520,240^A0N,90,80^FR^FD" + FCad + "^FS\n";
                                    imp100x50 += "^FO15,310^A0N,25,25^FDProducto de Mexico^FS\n";
                                    imp100x50 += "^FO15,340^A0N,20,20^FDDistribuido por: Comercializadora GAB SA de CV^FS\n";
                                    imp100x50 += "^FO15,360^A0N,20,20^FDCarretera Panamericana Km 291-1^FS\n";
                                    imp100x50 += "^FO15,380^A0N,20,20^FDCortazar, Guanajuato., Mexico^FS\n";
                                    imp100x50 += "^XZ\n";
                                    //if (ConLechuga == "S")
                                    ConLechuga = "S";
                                    //EtiCosto = "S";
                                }

                                #endregion

                                #region etiqueta para Japón ed 4"x2" 
                                if (clave_gab.Trim() == "02BROMLJAP")
                                {
                                    impr2 += "^XA\n";
                                    impr2 += "^FO645,30^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                    impr2 += "^FO15,25^A2N,23,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";
                                    impr2 += "^FO15,75^A0N,48,24^FD" + aux2 + "^FS\n";
                                    impr2 += "^FO15,140^A0N,48,24^FD" + aux5 + "^FS\n";
                                    impr2 += "^FO15,213^A0N,25,25^FDPRODUCT OF MEXICO^FS\n";
                                    impr2 += "^FO750,235^A0B,35,28,^FD" + mfeccad + "^FS\n";
                                    impr2 += "^FO740,193^A0N,21,15^FDC:" + i.ToString() + "^FS\n";
                                    impr2 += "^FO80,270,^BY2,^BCN,90,N,N,N^FD01" + aux3 + "10" + txtrecibo.Text + "^FS\n";
                                    impr2 += "^FO700,235^A0B,35,28,^FD" + lote + "^FS\n";
                                    impr2 += "^FO120,370^A2N,17,17,^FD" + aux3_texto + "(10)" + txtrecibo.Text + "^FS\n";
                                    impr2 += "^XZ\n";
                                }
                                #endregion

                                #region etiqueta dos al paso
                                else
                                {
                                    if (ConLechuga != "S" && Church == "N")
                                    {
                                        if (i < auxn)
                                        {
                                            impr = impr + "^XA\n";
                                            //impr += "~SD12\n"; //~SD15
                                            impr += "^PW832\n";
                                            impr = impr + "^FO270,25^BQN,2,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n"; //CODIGO 2D BQ DE CUADRO - PTI"
                                            impr = impr + "^FO680,25^BQN,2,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4_1 + "^FS\n";  // CODIGO 2D BQ DE CUADRO - PTI"
                                                                                                                                                                      //impr = impr + "^FO-15,20^A1N,22,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";//etiqueta 1
                                                                                                                                                                      //impr = impr + "^FO430,20^A1N,22,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar2 + "/" + totar2 + "^FS\n";//etiqueta 2
                                            impr = impr + "^FO-15,15^A2N,22,8,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";//etiqueta 1
                                            impr = impr + "^FO430,15^A2N,22,8,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar2 + "/" + totar2 + "^FS\n";//etiqueta 2
                                                                                                                                                                //nombre en ingles etiqueta 1
                                            if (aux2.Length > 25)
                                            {
                                                impr = impr + "^FO-15,55^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                impr = impr + "^FO-15,80^A0N,22,18^FD" + aux2.Substring(25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                            }
                                            else
                                                impr = impr + "^FO-15,55^A0N,22,18^FD" + aux2 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1

                                            //nombre en español etiqueta 1
                                            if (aux5.Length > 25)
                                            {
                                                impr = impr + "^FO-15,120^A0N,22,18^FD" + aux5.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                impr = impr + "^FO-15,143^A0N,22,18^FD" + aux5.Substring(25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                            }
                                            else
                                                impr = impr + "^FO-15,120^A0N,22,18^FD" + aux5 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1


                                            //nombre en ingles etiqueta 2
                                            if (aux2.Length > 25)
                                            {
                                                impr = impr + "^FO430,55^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2
                                                impr = impr + "^FO430,80^A0N,22,18^FD" + aux2.Substring(25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2
                                            }
                                            else
                                                impr = impr + "^FO430,55^A0N,22,18^FD" + aux2 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2

                                            //nombre en español etiqueta 2
                                            if (aux5.Length > 25)
                                            {
                                                impr = impr + "^FO430,120^A0N,22,18^FD" + aux5.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2
                                                impr = impr + "^FO430,143^A0N,22,18^FD" + aux5.Substring(25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2
                                            }
                                            else
                                                impr = impr + "^FO430,120^A0N,22,18^FD" + aux5 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2


                                            impr = impr + "^FO320,170^A0B,25,25,^FD" + lote + "^FS\n";//LOTE ETIQUETA 1
                                            impr = impr + "^FO740,170^A0B,25,25,^FD" + lote + "^FS\n";//LOTE ETIQUETA 2

                                            if (txtrchtbl.Text.Trim() == "PILAR-2" && clave_gab == "05005LEORN" || clave_gab == "05005LSANA")
                                            {
                                                impr = impr + "^FO210,150^A0N,22,18^FD" + "18MR4001 ^FS\n";
                                                impr = impr + "^FO630,150^A0N,22,18^FD" + "18MR4001 ^FS\n";
                                            }
                                            if (txtrchtbl.Text.Trim() == "PILAR-4" && clave_gab == "05005LEORN" || clave_gab == "05005LSANA")
                                            {
                                                impr = impr + "^FO210,150^A0N,22,18^FD" + "18MR4002 ^FS\n";
                                                impr = impr + "^FO630,150^A0N,22,18^FD" + "18MR4002 ^FS\n";

                                            }

                                            impr = impr + "^FO320,140^A0N,22,18^FD" + "C: " + auxnum1 + "^FS\n"; // CAJA ETIQUETA 1
                                            impr = impr + "^FO740,140^A0N,22,18^FD" + "C: " + auxnum2 + "^FS\n"; // CAJA ETIQUETA 2

                                            if ((clave_gab.Trim() == "18007JI56V") || (clave_gab.Trim() == "18JIBOML66") || (clave_gab.Trim() == "16001TO561") ||
                                               (clave_gab.Trim() == "18007JI55M") || (clave_gab == "18007JVM55") || (clave_gab == "18007JVM66") ||
                                               (clave_gab.Trim() == "18007JI56M") || (clave_gab == "18007JVM56") || (clave_gab == "18JIML6610"))
                                            {
                                                impr = impr + "^FO-15,185^A0B,30,25,^FD" + grado + num + "^FS\n"; // SE IMPRIME TIPO GRADO 1
                                                impr = impr + "^FO430,185^A0B,30,25,^FD" + grado + num + "^FS\n"; // SE IMPRIME TIPO GRADO 2
                                            }

                                            if ((clave_gab.Trim() == "05006MLNA2"))
                                            {
                                                impr = impr + "^FO-15,170^A0B,13,15,^FDPACK DATE:^FS\n"; // fecha etiqueta 1    
                                                impr = impr + "^FO-30,180^A0B,17,18,^FD" + Convert.ToDateTime(fec).ToString("MM/dd/yy") + "^FS\n"; // fecha etiqueta 1     
                                                impr = impr + "^FO430,170^A0B,13,15,^FDPACK DATE:^FS\n"; // fecha etiqueta 2 
                                                impr = impr + "^FO445,180^A0B,17,18,^FD" + Convert.ToDateTime(fec).ToString("MM/dd/yy") + "^FS\n"; // fecha etiqueta 2
                                                                                                                                                   //impr = impr + "^FO-15,175^A0B,17,18,^FD" + Convert.ToDateTime(fec).ToString("MM/dd/yy") + "^FS\n"; // fecha etiqueta 1                                                
                                                                                                                                                   //impr = impr + "^FO430,175^A0B,17,18,^FD" + Convert.ToDateTime(fec).ToString("MM/dd/yy") + "^FS\n"; // fecha etiqueta 2                                                 
                                            }

                                            impr = impr + "^FO70,170,^BY1,^BCN,60,N,N,N^FD" + "01" + aux3 + "^FS\n"; // CODIGO GTIN COD DE BARRAS 128 1
                                            impr = impr + "^FO495,170,^BY1,^BCN,60,N,N,N^FD" + "01" + aux3 + "^FS\n"; // CODIGO GTIN COD DE BARRAS 128 2
                                            impr = impr + "^FO110,240^A0N,17,15,^FD" + aux3_texto + "^FS\n"; // DESCRIPCION DEL PRODUCTO EN EL CODIGO BQ LINEA 1 
                                            if (etilote.Length > 0)
                                            {
                                                impr = impr + "^FO290,145^A0B,17,18,^FD" + etilote + "^FS\n"; // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 1
                                                impr = impr + "^FO710,145^A0B,17,18,^FD" + etilote + "^FS\n";
                                            }
                                            impr = impr + "^FO350,165^A0B,17,18,^FD" + mfeccad + "^FS\n"; // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 1
                                            impr = impr + "^FO770,165^A0B,17,18,^FD" + mfeccad + "^FS\n"; // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 2    
                                            impr = impr + "^FO510,240^A0N,17,17,^FD" + aux3_texto + "^FS\n"; // DESCRIPCION DEL PRODUCTO EN EL CODIGO BQ LINEA 2
                                            impr = impr + "^XZ\n";

                                            #region
                                            /*sw.WriteLine("^XA");
                                        sw.WriteLine("^FO270,20^BQN,2,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS"); //CODIGO 2D BQ DE CUADRO - PTI
                                        sw.WriteLine("^FO680,20^BQN,2,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4_1 + "^FS");  // CODIGO 2D BQ DE CUADRO - PTI
                                        sw.WriteLine("^FO10,15^A0N,22,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS");//etiqueta 1
                                        sw.WriteLine("^FO430,15^A0N,22,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar2 + "/" + totar2 + "^FS");//etiqueta 2

                                        //nombre en ingles etiqueta 1
                                        if (aux2.Length > 25)
                                        {
                                            sw.WriteLine("^FO10,45^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                            sw.WriteLine("^FO10,70^A0N,22,18^FD" + aux2.Substring(26, 24) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                        }
                                        else
                                            sw.WriteLine("^FO10,45^A0N,22,18^FD" + aux2 + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1

                                        //nombre en español etiqueta 1
                                        if (aux5.Length > 25)
                                        {
                                            sw.WriteLine("^FO10,110^A0N,22,18^FD" + aux5.Substring(0, 25) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                            sw.WriteLine("^FO10,133^A0N,22,18^FD" + aux5.Substring(26, 19) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                        }
                                        else
                                            sw.WriteLine("^FO10,110^A0N,22,18^FD" + aux5 + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1

                                        //nombre en ingles etiqueta 2
                                        if (aux2.Length > 25)
                                        {
                                            sw.WriteLine("^FO430,45^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 2
                                            sw.WriteLine("^FO430,70^A0N,22,18^FD" + aux2.Substring(26, 24) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 2
                                        }
                                        else
                                            sw.WriteLine("^FO430,45^A0N,22,18^FD" + aux2 + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 2

                                        //nombre en español etiqueta 2
                                        if (aux5.Length > 25)
                                        {
                                            sw.WriteLine("^FO430,110^A0N,22,18^FD" + aux5.Substring(0, 25) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 2
                                            sw.WriteLine("^FO430,133^A0N,22,18^FD" + aux5.Substring(26, 19) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 2
                                        }
                                        else
                                            sw.WriteLine("^FO430,110^A0N,22,18^FD" + aux5 + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 2

                                        sw.WriteLine("^FO320,140^A0N,22,18^FD" + "C: " + auxnum1 + "^FS"); // CAJA ETIQUETA 1
                                        sw.WriteLine("^FO740,140^A0N,22,18^FD" + "C: " + auxnum2 + "^FS"); // CAJA ETIQUETA 2
                                        sw.WriteLine("^FO90,160,^BY1,^BCN,60,N,N,N^MD5F^FD" + aux3 + "^FS"); // CODIGO GTIN COD DE BARRAS 128 1
                                        sw.WriteLine("^FO495,160,^BY1,^BCN,60,N,N,N^FD" + aux3 + "^FS"); // CODIGO GTIN COD DE BARRAS 128 2
                                        sw.WriteLine("^FO110,230^A0N,17,15,^FD" + aux3_texto + "^FS"); // DESCRIPCION DEL PRODUCTO EN EL CODIGO BQ LINEA 1                            
                                        sw.WriteLine("^FO350,165^A0B,17,18,^FD" + mfeccad + "^FS"); // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 1
                                        sw.WriteLine("^FO770,165^A0B,17,18,^FD" + mfeccad + "^FS"); // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 2    
                                        sw.WriteLine("^FO510,230^A0N,17,17,^FD" + aux3_texto + "^FS"); // DESCRIPCION DEL PRODUCTO EN EL CODIGO BQ LINEA 2
                                        sw.WriteLine("^XZ");*/
                                            i++;
                                            #endregion
                                        }
                                        else
                                        {
                                            impr = impr + "^XA\n";
                                            //impr += "~SD12\n"; //~SD15
                                            impr += "^PW832\n";
                                            impr = impr + "^FO270,25^BQN,2,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n"; //CODIGO 2D BQ DE CUADRO - PTI"
                                            impr = impr + "^FO-15,15^A2N,22,8,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";//etiqueta 1
                                                                                                                                                                //nombre en ingles etiqueta 1
                                            if (aux2.Length > 25)
                                            {
                                                impr = impr + "^FO-15,55^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                impr = impr + "^FO-15,80^A0N,22,18^FD" + aux2.Substring(25, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                            }
                                            else
                                                impr = impr + "^FO-15,55^A0N,22,18^FD" + aux2 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1

                                            //nombre en español etiqueta 1
                                            if (aux5.Length > 25)
                                            {
                                                impr = impr + "^FO-15,120^A0N,22,18^FD" + aux5.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                impr = impr + "^FO-15,143^A0N,22,18^FD" + aux5.Substring(25, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                            }
                                            else
                                                impr = impr + "^FO-15,120^A0N,22,18^FD" + aux5 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1


                                            impr = impr + "^FO320,140^A0N,22,18^FD" + "C: " + auxnum1 + "^FS\n"; // CAJA ETIQUETA 1
                                            impr = impr + "^FO320,170^A0B,25,25,^FD" + lote + "^FS\n";//LOTE ETIQUETA 1

                                            if (etilote.Length > 0)
                                            {
                                                impr = impr + "^FO290,145^A0B,17,18,^FD" + etilote + "^FS\n"; // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 1
                                            }
                                            if ((clave_gab.Trim() == "18007JI56V") || (clave_gab.Trim() == "18JIBOML66") || (clave_gab.Trim() == "16001TO561") ||
                                               (clave_gab.Trim() == "18007JI55M") || (clave_gab == "18007JVM55") || (clave_gab == "18007JVM66") ||
                                               (clave_gab.Trim() == "18007JI56M") || (clave_gab == "18007JVM56") || (clave_gab == "18JIML6610"))
                                            {
                                                impr = impr + "^FO-15,185^A0B,30,25,^FD" + grado + "^FS\n"; // SE IMPRIME TIPO GRADO 1                                                                                
                                            }

                                            if (clave_gab.Trim() == "05006MLNA2")
                                            {
                                                impr = impr + "^FO-15,175^A0B,17,18,^FD" + Convert.ToDateTime(fec).ToString("MM/dd/yy") + "^FS\n"; // fecha etiqueta 1                                                                                     
                                            }
                                            impr = impr + "^FO90,170,^BY1,^BCN,60,N,N,N^MD5F^FD" + "01" + aux3 + "^FS\n"; // CODIGO GTIN COD DE BARRAS 128 1                                    
                                            impr = impr + "^FO110,240^A0N,17,15,^FD" + aux3_texto + "^FS\n"; // DESCRIPCION DEL PRODUCTO EN EL CODIGO BQ LINEA 1                            
                                            impr = impr + "^FO350,165^A0B,17,18,^FD" + mfeccad + "^FS\n"; // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 1                                    
                                            impr = impr + "^XZ\n";

                                            #region
                                            /*sw.WriteLine("^XA");
                                        sw.WriteLine("^FO270,20^BQN,2,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS"); //CODIGO 2D BQ DE CUADRO - PTI                                    
                                        sw.WriteLine("^FO10,15^A0N,22,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS");//etiqueta 1
                                        sw.WriteLine("^FO10,45^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                        sw.WriteLine("^FO10,70^A0N,22,18^FD" + aux2.Substring(26, 24) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1

                                        //nombre en ingles etiqueta 1
                                        if (aux2.Length > 25)
                                        {
                                            sw.WriteLine("^FO10,45^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                            sw.WriteLine("^FO10,70^A0N,22,18^FD" + aux2.Substring(26, 24) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                        }
                                        else
                                            sw.WriteLine("^FO10,45^A0N,22,18^FD" + aux2 + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1

                                        //nombre en español etiqueta 1
                                        if (aux5.Length > 25)
                                        {
                                            sw.WriteLine("^FO10,110^A0N,22,18^FD" + aux5.Substring(0, 25) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                            sw.WriteLine("^FO10,133^A0N,22,18^FD" + aux5.Substring(26, 19) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                        }
                                        else
                                            sw.WriteLine("^FO10,110^A0N,22,18^FD" + aux5 + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1                     

                                        sw.WriteLine("^FO320,140^A0N,22,18^FD" + "C: " + auxnum1 + "^FS"); // caja etiqueta 1
                                        sw.WriteLine("^FO90,160,^BY1,^BCN,60,N,N,N^MD5F^FD" + aux3 + "^FS"); // CODIGO GTIN COD DE BARRAS 128 1
                                        sw.WriteLine("^FO110,230^A0N,17,15,^FD" + aux3_texto + "^FS"); // DESCRIPCION DEL PRODUCTO EN EL CODIGO BQ LINEA 1                                    
                                        sw.WriteLine("^FO350,165^A0B,17,18,^FD" + mfeccad + "^FS"); // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 1
                                        sw.WriteLine("^XZ");*/
                                            #endregion
                                        }
                                        //i++;
                                    }
                                }
                                #endregion
                            }
                            //sw.WriteLine("^XA");
                            //sw.WriteLine("^FO80,30^A0N,45,20^FD -     -      -    -     -   -^FS");// DESCRIPCION DEL PRODUCTO LINEA 2
                            //sw.WriteLine("^FO460,30^A0N,45,20^FD -     -      -    -     -   -^FS");// DESCRIPCION DEL PRODUCTO LINEA 1
                            //sw.WriteLine("^FO60,30,^BY3,^BCN,80,N,N,N^FD00000^XZ");// CODIGO GTIN COD DE BARRAS 128                                    
                        }
                        //if (eti_grande == "S")
                        //{
                        //    sw.WriteLine("^XA");
                        //    sw.WriteLine("^FO80,30^A0N,45,20^FD -     -      -    -     -   -^FS");// DESCRIPCION DEL PRODUCTO LINEA 2
                        //    //sw.WriteLine("^FO460,30^A0N,45,20^FD -     -      -    -     -   -^FS");// DESCRIPCION DEL PRODUCTO LINEA 1
                        //    sw.WriteLine("^FO60,30,^BY3,^BCN,80,N,N,N^FD00000^XZ");// CODIGO GTIN COD DE BARRAS 128                    
                        //}
                        //else
                        //{
                        //if (pais_origen.Trim() == "02BROMLJAP")
                        //{
                        //    impr = impr + "^XA";
                        //    impr = impr + "^FO80,30^A0N,45,20^FD -     -      -    -     -   -^FS";
                        //    impr = impr + "^FO460,30^A0N,45,20^FD -     -      -    -     -   -^FS";
                        //    impr = impr + "^FO60,30,^BY3,^BCN,80,N,N,N^FD00000^XZ";
                        //}
                        //else 
                        //if (clave_gab.Trim() == "")
                        if (impr.Trim() != "")
                        {
                            impr = impr + "^XA\n";
                            impr = impr + "^FO80,30^A0N,45,20^FD -     -      -    -     -   -^FS\n";
                            impr = impr + "^FO460,30^A0N,45,20^FD -     -      -    -     -   -^FS\n";
                            impr = impr + "^FO60,30,^BY3,^BCN,80,N,N,N^FD00000^XZ\n";
                        }
                        else
                        {
                            impr2 = impr2 + "^XA\n";
                            impr2 = impr2 + "^FO80,30^A0N,45,20^FD -     -      -    -     -   -^FS\n";
                            impr2 = impr2 + "^FO460,30^A0N,45,20^FD -     -      -    -     -   -^FS\n";
                            impr2 = impr2 + "^FO60,30,^BY3,^BCN,80,N,N,N^FD00000^XZ\n";
                        }
                        /*sw.WriteLine("^XA");
                        sw.WriteLine("^FO80,30^A0N,45,20^FD -     -      -    -     -   -^FS");
                        sw.WriteLine("^FO460,30^A0N,45,20^FD -     -      -    -     -   -^FS");
                        sw.WriteLine("^FO60,30,^BY3,^BCN,80,N,N,N^FD00000^XZ");*/
                        //}
                    }
                    if (reader1.HasRows == false) // no estan impresas las Etiquetas
                    {
                        string auxtr_1 = "", auxtr_2 = "", auxtr_3 = "", auxtr_4 = "", auxtr_5 = "", auxtr_6 = "", auxtr_7 = "", auxtr_8 = "", auxtr_9 = "", auxtr_10 = "", auxtr_11 = "", auxtr_12 = "", auxtr_13 = "", auxtr_14 = "", auxtr_15 = "", auxtr_16 = "", aux_ingles = "";
                        //for (int i = 0; i < tmp_trazabilidad.Rows.Count; i++)
                        //{
                        //    //if (Convert.ToString(DGV5.Rows[x].Cells["clave_producto"].Value).Trim() == Convert.ToString(tmp_trazabilidad.Rows[i]["trz_codigo"].ToString().Trim()))
                        if (Convert.ToString(DGV5.Rows[x].Cells["tar"].Value).Trim().Length == 1)
                            DGV5.Rows[x].Cells["tar"].Value = "0" + Convert.ToString(DGV5.Rows[x].Cells["tar"].Value);
                        if (Convert.ToString(DGV5.Rows[x].Cells["tarimas_total"].Value).Trim().Length == 3)
                        {
                            DGV5.Rows[x].Cells["tar"].Value = Convert.ToString(DGV5.Rows[x].Cells["tar"].Value).PadLeft(3, '0');
                        }
                        string CodigoBar = "";
                        foreach (DataRow rows in tmp_trazabilidad.Select("trz_recibo = '" + txtrecibo.Text + "' and trz_ini = '" + Convert.ToString(DGV5.Rows[x].Cells["tar"].Value) + "' and trz_codigo = '" + Convert.ToString(DGV5.Rows[x].Cells["clave_producto"].Value).Trim() + "'"))
                        {
                            #region SQL

                            auxtr_2 = Convert.ToString(rows["trz_codigo"].ToString());
                            auxtr_3 = Convert.ToString(rows["trz_ini"].ToString());
                            //auxtr_3 = Convert.ToString(rows["trz_codigo_barras"].ToString());
                            //auxtr_3 = auxtr_3.Substring(auxtr_3.Length - 5, 2);

                            auxtr_4 = Convert.ToString(rows["trz_cajas_por_tarima"].ToString());
                            auxtr_5 = Convert.ToString(rows["trz_codigo_barras"].ToString());
                            auxtr_5 = auxtr_5.Replace("*", "");
                            auxtr_6 = Convert.ToString(rows["trz_descrip"].ToString());
                            auxtr_6 = auxtr_6.Replace("'", "''");
                            auxtr_7 = Convert.ToString(rows["trz_recibo"].ToString());
                            auxtr_8 = "PTC";
                            auxtr_9 = Recepcion_PT.provee;
                            auxtr_11 = Recepcion_PT.rancho;
                            auxtr_13 = Recepcion_PT.tabla;
                            auxtr_15 = txtclave.Text;
                            auxtr_16 = Convert.ToString(rows["trz_id_SSCC"].ToString());
                            mfeccad = Convert.ToString(rows["trz_nota"].ToString());

                            if (rows["trz_codigo_barras"].ToString().Trim().Length > 0)
                                CodigoBar = rows["trz_codigo_barras"].ToString().Trim().Substring(0, rows["trz_codigo_barras"].ToString().Trim().Length - 1);
                            string PtiFamous = "";
                            if (rows["trz_pti_famous"].ToString().Trim().Length > 0)
                                PtiFamous = rows["trz_pti_famous"].ToString().Trim();

                            //auxtr_2 = Convert.ToString(tmp_trazabilidad.Rows[i]["trz_codigo"].ToString());
                            //auxtr_3 = Convert.ToString(tmp_trazabilidad.Rows[i]["trz_codigo_barras"].ToString());
                            //auxtr_3 = auxtr_3.Substring(auxtr_3.Length - 5, 2);

                            //auxtr_4 = Convert.ToString(tmp_trazabilidad.Rows[i]["trz_cajas_por_tarima"].ToString());
                            //auxtr_5 = Convert.ToString(tmp_trazabilidad.Rows[i]["trz_codigo_barras"].ToString());
                            //auxtr_5 = auxtr_5.Replace("*", "");
                            //auxtr_6 = Convert.ToString(tmp_trazabilidad.Rows[i]["trz_descrip"].ToString());
                            //auxtr_6 = auxtr_6.Replace("'", "''");
                            //auxtr_7 = Convert.ToString(tmp_trazabilidad.Rows[i]["trz_recibo"].ToString());
                            //auxtr_8 = "PTC";
                            //auxtr_9 = Recepcion_PT.provee;
                            //auxtr_11 = Recepcion_PT.rancho;
                            //auxtr_13 = Recepcion_PT.tabla;
                            //auxtr_15 = txtclave.Text;
                            //mfeccad = Convert.ToString(tmp_trazabilidad.Rows[i]["trz_nota"].ToString());
                            //auxtr_11 = "AL";
                            //thisConnection.Open();

                            //cmnd1 = thisConnection.CreateCommand();
                            //cmnd1.CommandText = "SELECT prod_tipoeti FROM TB_CAT_PRODUCTO WHERE PROD_CLAVE = '" + auxtr_2 + "'";
                            //string tipo_eti = Convert.ToString(cmnd1.ExecuteScalar()).Trim();
                            //if (tipo_eti != "2")
                            //{
                            //    thisConnection.Close();
                            //    MessageBox.Show("No se puede imprimir la etiqueta para el producto " + auxtr_2.ToString() + " " + auxtr_6.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            //    return;
                            //}


                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "SELECT prod_codegtin, prod_nomb_ingles, prod_nombre, prod_paisorigen FROM TB_CAT_PRODUCTO WHERE PROD_CLAVE = '" + auxtr_2 + "'";
                            reader1 = cmnd1.ExecuteReader();
                            while (reader1.Read())
                            {
                                auxtr_1 = reader1.GetValue(0).ToString().Trim();
                                if (auxtr_1.Trim() == "")
                                    auxtr_1 = auxtr_2;

                                aux_ingles = reader1.GetValue(1).ToString().Replace("'", " ");
                                if (aux_ingles.Length > 160)
                                    aux_ingles = aux_ingles.Substring(100, 50).Trim();
                                //if (auxtr_2 == "07ARML1217")
                                //    aux_ingles = aux_ingles.Substring(77, 22);
                                //if (auxtr_2 == "11001ML30E")
                                //    aux_ingles = aux_ingles.Substring(99,22);
                                //if (auxtr_2 == "05LENACVFA")
                                //    aux_ingles = aux_ingles.Substring(83, 32);


                                auxtr_6 = reader1.GetValue(2).ToString();
                                auxtr_6 = auxtr_6.Replace("'", " ");
                                pais_origen = reader1.GetValue(3).ToString().Trim();
                            }
                            reader1.Read();


                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "SELECT PROV_NOMBRE FROM TB_CAT_PROVEEDOR WHERE PROV_CLAVE = '" + auxtr_9 + "'";
                            reader1 = cmnd1.ExecuteReader();
                            while (reader1.Read())
                            {
                                auxtr_10 = reader1.GetValue(0).ToString().Trim();
                            }
                            reader1.Dispose();

                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "SELECT RCH_NOMBRE FROM TB_CAT_RANCHOS WHERE PROV_CLAVE = '" + auxtr_9 + "' and rch_clave ='" + auxtr_11 + "'";
                            reader1 = cmnd1.ExecuteReader();
                            while (reader1.Read())
                            {
                                auxtr_12 = reader1.GetValue(0).ToString().Trim();
                            }
                            reader1.Dispose();


                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "SELECT TBL_NOMBRE FROM TB_CAT_TABLAS WHERE PROV_CLAVE = '" + auxtr_9 + "' and rch_clave = '" + auxtr_11 + "' and tbl_clave = '" + auxtr_13 + "'";
                            reader1 = cmnd1.ExecuteReader();
                            while (reader1.Read())
                            {
                                auxtr_14 = reader1.GetValue(0).ToString().Trim();
                            }
                            reader1.Dispose();

                            fecad = mfeccad;//fecha caducidad
                            if (fecad.Trim() != "")
                            //string lin = auxtr_2.Substring(0, 2);//saber la linea
                            //if ((auxtr_2.Trim() == "05003ML3P") || (auxtr_2.Trim() == "05005ML2P") || (auxtr_2.Trim() == "19"))
                            {
                                dia = fecad.Substring(0, 2);
                                fecad = obtenerNombreMesNumero(Convert.ToInt32(fecad.Substring(3, 2))).ToUpper();
                                //fecad = fecad.Substring(3, 2);
                            }
                            else
                            {
                                dia = "0";
                                fecad = "";
                            }

                            //guarda en tb_hist_recepcion de DBGAB                            
                            decimal trfin = Convert.ToDecimal(tmp_trazabilidad.Rows[i]["trz_final"].ToString());
                            decimal trini = Convert.ToDecimal(auxtr_3);
                            decimal dy = Convert.ToDecimal(dia);
                            //thisConnection2.Open();
                            //cmnd2 = thisConnection2.CreateCommand();
                            //cmnd2.CommandText = "insert into tb_hist_recepcion (hrp_recibo, hrp_fecha, hrp_tipo_r, lin_clave, prod_clave, hrp_num_un, hrp_surtido, hrp_nomprod, " +
                            //                    "tarini, tarfin, HORA, OPCION, DIA, MES) values " +
                            //                    "('" + auxtr_7 + "', '" + DateTime.Now.ToShortDateString() + "', '" + auxtr_8 + "', '" + linea + "', '" + auxtr_2 + "', " +
                            //                    "" + Convert.ToDecimal(auxtr_4) + ", 0, '" + auxtr_6 + "', " + trini + ", " + trfin + ", '" + txthora.Text + "', " +
                            //                    "' ', " + dy + ", '" + fecad.ToUpper() + "')";
                            //reader2 = cmnd2.ExecuteReader();
                            //reader2.Dispose();
                            //thisConnection2.Close();

                            //guarda en tb_det_trazabilidad
                            string PesoXCaja = getPesoXCaja(auxtr_7, auxtr_2);
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "insert into tb_det_trazabilidad (gtin_clave, prod_clave, tarima, etiqueta, pti_clave, prod_nombre, recibo, tipo, prov_clave, prov_nombre, rch_clave, rch_nombre, tbl_clave, " +
                                              "tbl_nombre, lote, pti_fecha, prod_nom_ingles, fecha_cad, fecha_emb, fcn_folio, pti_hora, pti_estatus, surtido, pti_estatus_sur, id_Pallet, pti_famous, PesoXCaja) values ('" + auxtr_1 + "', '" + auxtr_2 + "', " +
                                              "'" + Convert.ToDecimal(auxtr_3) + "', '" + Convert.ToDecimal(auxtr_4) + "', '" + auxtr_5 + "', '" + auxtr_6 + "', '" + auxtr_7 + "', '" + auxtr_8 + "', '" + auxtr_9 + "', " +
                                              "'" + auxtr_10 + "', '" + auxtr_11 + "', '" + auxtr_12 + "', '" + auxtr_13 + "', '" + auxtr_14 + "', '" + auxtr_15 + "', '" + date + "', " +
                                              "'" + aux_ingles + "', '" + mfeccad + "', '', '', '', '', 0, '', '" + auxtr_16 + "','" + PtiFamous + "', '" + PesoXCaja + "')";
                            reader1 = cmnd1.ExecuteReader();
                            //thisConnection.Close();
                            reader1.Dispose();

                            #endregion

                            #region Fox
                            /*try
                            {
                                MyConnection.Open();
                                OleDbCommand dbCmdNull1 = MyConnection.CreateCommand();
                                dbCmdNull1.CommandText = "SET NULL OFF";
                                dbCmdNull1.ExecuteNonQuery();
                                dbCmdNull1.Dispose();

                                cmd1 = MyConnection.CreateCommand();
                                cmd1.CommandText = "insert into tb_det_trazabilidad ([gtin_clave], [prod_clave], [tarima], [etiqueta], [pti_clave], [prod_nombre], [recibo], [tipo], [prov_clave], [prov_nombre], [rch_clave], " +
                                                 "[rch_nombre], [tbl_clave], [tbl_nombre], [lote], [pti_fecha], [prod_nom_ingles], [fecha_cad]) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
                                cmd1.Parameters.Add("@Nombre1", OleDbType.Char).Value = auxtr_1;
                                cmd1.Parameters.Add("@Nombre2", OleDbType.Char).Value = auxtr_2;
                                cmd1.Parameters.Add("@Nombre3", OleDbType.Numeric).Value = float.Parse(auxtr_3);
                                cmd1.Parameters.Add("@Nombre4", OleDbType.Numeric).Value = float.Parse(auxtr_4);
                                cmd1.Parameters.Add("@Nombre5", OleDbType.Char).Value = auxtr_5;
                                cmd1.Parameters.Add("@Nombre6", OleDbType.Char).Value = auxtr_6;
                                cmd1.Parameters.Add("@Nombre7", OleDbType.Char).Value = auxtr_7;
                                cmd1.Parameters.Add("@Nombre8", OleDbType.Char).Value = auxtr_8;
                                cmd1.Parameters.Add("@Nombre9", OleDbType.Char).Value = auxtr_9;
                                cmd1.Parameters.Add("@Nombre10", OleDbType.Char).Value = auxtr_10;
                                cmd1.Parameters.Add("@Nombre11", OleDbType.Char).Value = auxtr_11;
                                cmd1.Parameters.Add("@Nombre12", OleDbType.Char).Value = auxtr_12;
                                cmd1.Parameters.Add("@Nombre13", OleDbType.Char).Value = auxtr_13;
                                cmd1.Parameters.Add("@Nombre14", OleDbType.Char).Value = auxtr_14;
                                cmd1.Parameters.Add("@Nombre15", OleDbType.Char).Value = auxtr_15;
                                cmd1.Parameters.Add("@Nombre16", OleDbType.Date).Value = DateTime.Now.ToShortDateString();
                                cmd1.Parameters.Add("@Nombre17", OleDbType.Char).Value = aux_ingles;
                                cmd1.Parameters.Add("@Nombre18", OleDbType.Char).Value = mfeccad;
                                read1 = cmd1.ExecuteReader();
                                read1.Dispose();
                                MyConnection.Close();
                            }
                            catch (OleDbException ex)
                            {
                                MyConnection.Close();
                                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.3", ex.ToString().Trim(), "PTFOX");
                                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString().Trim());
                                //Utilerias.Class1.registrar_movimiento(DateTime.Now, Utilerias.Class1.Nombre_equipo, Utilerias.Class1.Usuario, "errorSQL", Utilerias.Class1.Formulario, txtrecibo.Text, ex.ToString());
                                MessageBox.Show(ex.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }*/
                            #endregion

                            //DataRow row = eti_blanca.trazabilidad.NewRow();
                            //row["pti_clave"] = auxtr_5;
                            //row["prod_nom_ingles"] = aux_ingles;
                            //row["gtin_clave"] = auxtr_1;
                            //row["prod_clave"] = auxtr_2;
                            //row["prod_nombre"] = auxtr_6;
                            //row["fecha_cad"] = mfeccad;
                            //row["etiqueta"] = auxtr_4;
                            //row["tarima"] = auxtr_3;
                            //row["recibo"] = txtrecibo.Text;
                            //eti_blanca.trazabilidad.Rows.Add(row);

                            //fec = rows["trz_nota"].ToString().Trim();
                            aux1 = auxtr_5; //pti_clave
                            aux5 = aux_ingles;//prod_nombre_ingles
                            aux3 = auxtr_1;//gtin_clave
                            clave_gab = auxtr_2.Trim();//prod_clave
                            aux2 = auxtr_6;//prod_nombre
                                           //aux5 = aux5.Replace("'", "  ");                        
                            fecad = mfeccad;//fecha caducidad
                                            //if (fecad.Trim() != "")
                            string lin = auxtr_2.Substring(0, 2);//saber la linea
                            lote = txtclave.Text.Trim();
                            if ((auxtr_2.Trim() == "05003ML3P") || (auxtr_2.Trim() == "05005ML2P") || (lin.Trim() == "19"))
                            {
                                if (fecad.Trim() != "")
                                {
                                    //dia = fecad.Substring(0, 2);
                                    //fecad = fecad.Substring(3, 2);
                                    //obtenerNombreMesNumero(Convert.ToInt32(fecad));
                                    //mfeccad = "FC" + mes.ToUpper() + dia;
                                    if (lin.Trim() == "19")
                                        etilote = "LOTE:" + TrarItem(auxtr_2.Trim()).Trim() + " 02";
                                    else
                                        etilote = "LOTE: " + txtrecibo.Text;
                                    anio = new string(fecad.Trim().Reverse().Take(2).Reverse().ToArray());
                                    dia = fecad.Substring(0, 2);
                                    fecad = fecad.Substring(3, 2);
                                    obtenerNombreMesNumero(Convert.ToInt32(fecad));
                                    mfeccad = "CAD " + dia + mes.ToUpper() + anio;
                                }
                            }
                            else
                            {
                                mfeccad = "";
                            }
                            auxn = Convert.ToDecimal(auxtr_4);//etiqueta

                            if (aux3 == "")
                            {
                                aux3 = clave_gab;
                                aux3_texto = "(01)" + aux3;
                            }
                            else
                                aux3_texto = "(01)" + aux3;

                            tot_cajas = tot_cajas + auxn;

                            for (int j = 1; j <= auxn; j++)
                            {
                                auxc = j.ToString();
                                if (auxc.Length == 1)
                                    auxd = "00" + auxc;
                                if (auxc.Length == 2)
                                    auxd = "0" + auxc;

                                //aux4 = aux1 + auxd;
                                auxnum1 = j.ToString();
                                aux4 = aux1 + auxnum1.ToString().Trim().PadLeft(3, '0');
                                y = j + 1;
                                auxnum2 = y.ToString();
                                auxc = y.ToString();
                                auxd = auxc;
                                auxd_1 = auxd;

                                if (auxd_1.Length == 1)
                                    auxd_1 = "00" + auxd;
                                if (auxc.Length == 2)
                                    auxd_1 = "0" + auxd;

                                //aux4_1 = aux1 + auxd_1;
                                aux4_1 = aux1 + auxnum2.ToString().Trim().PadLeft(3, '0'); //auxd_1;
                                lon_cad1 = aux4.Length + 1;
                                lon_cad2 = aux4_1.Length + 1;

                                /*nutar1 = aux4.Substring(lon_cad1 - 8, 2);
                                totar1 = aux4.Substring(lon_cad1 - 6, 2);
                                nutar2 = aux4_1.Substring(lon_cad1 - 8, 2);
                                totar2 = aux4_1.Substring(lon_cad1 - 6, 2);*/

                                nutar1 = Convert.ToInt32(rows["trz_ini"].ToString()).ToString();
                                totar1 = Convert.ToInt32(rows["trz_final"].ToString()).ToString();

                                nutar2 = Convert.ToInt32(rows["trz_ini"].ToString()).ToString();
                                totar2 = Convert.ToInt32(rows["trz_final"].ToString()).ToString();
                                //PaperSize ps;
                                //PrintDocument pd1 = new System.Drawing.Printing.PrintDocument();
                                //ps = new PaperSize("etiqueta_blanca", 400, 600);
                                //pd1.PrintPage += new PrintPageEventHandler(this.printDocument2_PrintPage);
                                ////PrintPreviewDialog VistaPrevia1 = new PrintPreviewDialog();

                                ////VistaPrevia1.Document = pd1;
                                //pd1.DefaultPageSettings.PaperSize = ps;
                                ////VistaPrevia1.ShowDialog();                                
                                //pd1.Print();

                                //if (eti_grande == "S")
                                //{
                                #region etiqueta nueva
                                /*
                            if (pais_origen.Trim() != "")
                            {
                                //COMENTADO 08/12/2016
                                //etiqueta de papel
                                impr2 += "^XA\n";
                                impr2 += "^FO05,55^A2N,23,17,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";
                                impr2 += "^FO15,100^A0N,21,15^FDC:" + i.ToString() + "^FS\n";
                                impr2 += "^FO200,100^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                impr2 += "^FO15,250^A0N,48,24^FD" + aux2 + "^FS\n";
                                impr2 += "^FO15,300^A0N,48,24^FD" + aux5 + "^FS\n";
                                impr2 += "^FO15,350^A0N,25,25^FDPRODUCT OF " + pais_origen + "^FS\n";
                                impr2 += "^FO250,350^A0N,25,25^FDHEB US#1^FS\n";
                                impr2 += "^FO50,380,^BY2,3,^BCN,90,N,N,N^FD01" + aux3 + "^FS\n";
                                impr2 += "^FO150,480^A0N,25,25^FD" + aux3_texto + "^FS\n";
                                impr2 += "^FO15,520^A0N,25,25^FDPACKED BY: COMERCIALIZADORA GAB, S.A. DE C.V.^FS\n";
                                impr2 += "^FO15,550^A0N,25,25^FDCARRETERA PANAMERICANA KM. 291 - 1^FS\n";
                                impr2 += "^FO15,580^A0N,25,25^FDCOL. LA FORTALEZA CORTAZAR, GTO. C.P. 38300^FS\n";
                                impr2 += "^FO15,610^A0N,25,25^FDMEXICO R.F.C. CGA-960614-2C5^FS\n";
                                impr2 += "^XZ\n";
                                //impr2 += "^XA\n";
                                //impr2 = impr2 + "^FWB,0\n";
                                //impr2 += "^FO20,15^A2N,23,17,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";
                                //impr2 += "^FO622,200^A0N,21,15^FDC:" + i.ToString() + "^FS\n";
                                //impr2 += "^FO622,30^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                //impr2 += "^FO20,50^A0N,48,24^FD" + aux2 + "^FS\n";
                                //impr2 += "^FO20,100^A0N,48,24^FD" + aux5 + "^FS\n";
                                //impr2 += "^FO20,150^A0N,25,25^FDPRODUCT OF " + pais_origen + "^FS\n";
                                //impr2 += "^FO300,150^A0N,25,25^FDHEB US#1^FS\n";//AQUI
                                //impr2 += "^FO100,180,^BY2,3,^BCN,90,N,N,N^FD01" + aux3 + "^FS\n";
                                //impr2 += "^FO150,275^A0N,25,25^FD" + aux3_texto + "^FS\n";
                                //impr2 += "^FO20,305^A0N,20,20^FDPACKED BY: COMERCIALIZADORA GAB, S.A. DE C.V.^FS\n";
                                //impr2 += "^FO20,325^A0N,20,20^FDCARRETERA PANAMERICANA KM. 291 - 1^FS\n";
                                //impr2 += "^FO20,345^A0N,20,20^FDCOL. LA FORTALEZA CORTAZAR, GTO. C.P. 38300^FS\n";
                                //impr2 += "^FO20,365^A0N,20,20^FDMEXICO R.F.C. CGA-960614-2C5^FS\n";
                                //impr2 += "^XZ\n";
                            }
                            else
                            {
                                //etiqueta de 4"x 2"
                                impr += "^XA";
                                impr += "^FO645,30^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS";
                                impr += "^FO15,25^A2N,23,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS";
                                impr += "^FO15,75^A0N,48,24^FD" + aux2 + "^FS";
                                impr += "^FO15,140^A0N,48,24^FD" + aux5 + "^FS";
                                impr += "^FO15,213^A0N,25,25^FDPRODUCT OF MEXICO^FS";
                                impr += "^FO730,235^A0B,35,28,^FD" + mfeccad + "^FS";
                                impr += "^FO740,193^A0N,21,15^FDC:" + j.ToString() + "^FS";
                                impr += "^FO80,270,^BY2,^BCN,90,N,N,N^FD01" + aux3 + "10" + txtrecibo.Text + "^FS";
                                impr += "^FO120,370^A2N,17,17,^FD" + aux3_texto + "(10)" + txtrecibo.Text + "^FS";
                                impr += "^XZ";
                            }*/
                                #endregion
                                //}
                                //else
                                //{            

                                #region etiqueta para Japón ed 4"x2"
                                if (clave_gab.Trim() == "02BROMLJAP")
                                {
                                    impr2 += "^XA\n";
                                    impr2 += "~SD12\n";  //~SD15
                                    impr2 += "^PW832\n";
                                    impr2 += "^FO645,30^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                    impr2 += "^FO15,25^A2N,23,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";
                                    impr2 += "^FO15,75^A0N,48,24^FD" + aux2 + "^FS\n";
                                    impr2 += "^FO15,140^A0N,48,24^FD" + aux5 + "^FS\n";
                                    impr2 += "^FO15,213^A0N,25,25^FDPRODUCT OF MEXICO^FS\n";
                                    impr2 += "^FO750,235^A0B,35,28,^FD" + mfeccad + "^FS\n";
                                    impr2 += "^FO740,193^A0N,21,15^FDC:" + i.ToString() + "^FS\n";
                                    impr2 += "^FO80,270,^BY2,^BCN,90,N,N,N^FD01" + aux3 + "10" + txtrecibo.Text + "^FS\n";
                                    impr2 += "^FO700,235^A0B,35,28,^FD" + lote + "^FS\n";
                                    impr2 += "^FO120,370^A2N,17,17,^FD" + aux3_texto + "(10)" + txtrecibo.Text + "^FS\n";
                                    impr2 += "^XZ\n";
                                }
                                #endregion
                                #region etiqueta vieja
                                else
                                {
                                    if (ConLechuga != "S")
                                    {
                                        if (j < auxn)
                                        {
                                            if (pais_origen.Trim() != "")
                                            {
                                                //COMENTADO 08/12/2016
                                                //impr2 += "^XA\n";
                                                //impr2 += "^FO05,55^A2N,23,17,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";
                                                //impr2 += "^FO15,100^A0N,21,15^FDC:" + i.ToString() + "^FS\n";
                                                //impr2 += "^FO200,100^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                                //impr2 += "^FO15,250^A0N,48,24^FD" + aux2 + "^FS\n";
                                                //impr2 += "^FO15,300^A0N,48,24^FD" + aux5 + "^FS\n";
                                                //impr2 += "^FO15,350^A0N,25,25^FDPRODUCT OF " + pais_origen + "^FS\n";
                                                //impr2 += "^FO250,350^A0N,25,25^FDHEB US#1^FS\n";
                                                //impr2 += "^FO50,380,^BY2,3,^BCN,90,N,N,N^FD01" + aux3 + "^FS\n";
                                                //impr2 += "^FO150,480^A0N,25,25^FD" + aux3_texto + "^FS\n";
                                                //impr2 += "^FO15,520^A0N,25,25^FDPACKED BY: COMERCIALIZADORA GAB, S.A. DE C.V.^FS\n";
                                                //impr2 += "^FO15,550^A0N,25,25^FDCARRETERA PANAMERICANA KM. 291 - 1^FS\n";
                                                //impr2 += "^FO15,580^A0N,25,25^FDCOL. LA FORTALEZA CORTAZAR, GTO. C.P. 38300^FS\n";
                                                //impr2 += "^FO15,610^A0N,25,25^FDMEXICO R.F.C. CGA-960614-2C5^FS\n";
                                                //impr2 += "^XZ\n";
                                                impr2 += "^XA\n";
                                                impr2 += "~SD12\n"; //~SD15
                                                impr2 += "^PW832\n";
                                                impr2 = impr2 + "^FWB,0\n";
                                                impr2 += "^FO20,15^A2N,23,17,^FD" + txtrecibo.Text + "-" + clave_gab + "-" + nutar1 + "/" + totar1 + "^FS\n";
                                                impr2 += "^FO622,200^A0N,21,15^FDC:" + i.ToString() + "^FS\n";
                                                impr2 += "^FO622,30^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                                impr2 += "^FO20,50^A0N,48,24^FD" + aux2 + "^FS\n";
                                                impr2 += "^FO20,100^A0N,48,24^FD" + aux5 + "^FS\n";
                                                impr2 += "^FO20,150^A0N,25,25^FDPRODUCT OF " + pais_origen + "^FS\n";
                                                //impr2 += "^FO300,150^A0N,25,25^FDHEB US#1^FS\n";//AQUI
                                                impr2 += "^FO100,180,^BY2,3,^BCN,90,N,N,N^FD01" + aux3 + "^FS\n";
                                                impr2 += "^FO150,275^A0N,25,25^FD" + aux3_texto + "^FS\n";
                                                impr2 += "^FO20,305^A0N,20,20^FDPACKED BY: COMERCIALIZADORA GAB, S.A. DE C.V.^FS\n";
                                                impr2 += "^FO20,325^A0N,20,20^FDCARRETERA PANAMERICANA KM. 291 - 1^FS\n";
                                                impr2 += "^FO20,345^A0N,20,20^FDCOL. LA FORTALEZA CORTAZAR, GTO. C.P. 38300^FS\n";
                                                impr2 += "^FO20,365^A0N,20,20^FDMEXICO R.F.C. CGA-960614-2C5^FS\n";
                                                impr2 += "^XZ\n";
                                            }
                                            else
                                            {
                                                impr = impr + "^XA\n";
                                                impr += "~SD12\n"; //~SD15
                                                impr += "^PW832\n";
                                                impr = impr + "^FO270,20^BQN,2,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n"; //CODIGO 2D BQ DE CUADRO - PTI"
                                                impr = impr + "^FO680,20^BQN,2,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4_1 + "^FS\n";  // CODIGO 2D BQ DE CUADRO - PTI"
                                                impr = impr + "^FO-15,20^A0N,22,18,^FD" + txtrecibo.Text + "-" + clave_gab + "-" + nutar1 + "/" + totar1 + "^FS\n";//etiqueta 1
                                                impr = impr + "^FO430,20^A0N,22,18,^FD" + txtrecibo.Text + "-" + clave_gab + "-" + nutar2 + "/" + totar2 + "^FS\n";//etiqueta 2
                                                                                                                                                                   //nombre en ingles etiqueta 1
                                                if (aux2.Length > 25)
                                                {
                                                    impr = impr + "^FO-15,55^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                    impr = impr + "^FO-15,80^A0N,22,18^FD" + aux2.Substring(26) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                }
                                                else
                                                    impr = impr + "^FO-15,55^A0N,22,18^FD" + aux2 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1

                                                //nombre en español etiqueta 1
                                                if (aux5.Length > 25)
                                                {
                                                    impr = impr + "^FO-15,120^A0N,22,18^FD" + aux5.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                    impr = impr + "^FO-15,143^A0N,22,18^FD" + aux5.Substring(26) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                }
                                                else
                                                    impr = impr + "^FO-15,120^A0N,22,18^FD" + aux5 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1


                                                //nombre en ingles etiqueta 2
                                                if (aux2.Length > 25)
                                                {
                                                    impr = impr + "^FO430,55^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2
                                                    impr = impr + "^FO430,80^A0N,22,18^FD" + aux2.Substring(26) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2
                                                }
                                                else
                                                    impr = impr + "^FO430,55^A0N,22,18^FD" + aux2 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2

                                                //nombre en español etiqueta 2
                                                if (aux5.Length > 25)
                                                {
                                                    impr = impr + "^FO430,120^A0N,22,18^FD" + aux5.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2
                                                    impr = impr + "^FO430,143^A0N,22,18^FD" + aux5.Substring(26) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2
                                                }
                                                else
                                                    impr = impr + "^FO430,120^A0N,22,18^FD" + aux5 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2


                                                if ((clave_gab.Trim() == "18007JI56V") || (clave_gab.Trim() == "18JIBOML66") || (clave_gab.Trim() == "16001TO561") ||
                                               (clave_gab.Trim() == "18007JI55M") || (clave_gab == "18007JVM55") || (clave_gab == "18007JVM66") ||
                                               (clave_gab.Trim() == "18007JI56M") || (clave_gab == "18007JVM56") || (clave_gab == "18JIML6610"))
                                                {
                                                    impr = impr + "^FO-15,175^A0B,17,18,^FD" + grado + "^FS\n"; // SE IMPRIME TIPO GRADO 1
                                                    impr = impr + "^FO30,195^A1B,25,20,^FD" + num + "^FS\n"; // SE IMPRIME TIPO GRADO 1
                                                    impr = impr + "^FO430,175^A0B,17,18,^FD" + grado + "^FS\n"; // SE IMPRIME TIPO GRADO 2
                                                    impr = impr + "^FO450,195^A1B,25,20,^FD" + num + "^FS\n"; // SE IMPRIME TIPO GRADO 1
                                                }
                                                if (clave_gab.Trim() == "05006MLNA2")
                                                {
                                                    impr = impr + "^FO-15,170^A0B,13,15,^FDPACK DATE:^FS\n"; // fecha etiqueta 1    
                                                    impr = impr + "^FO-30,180^A0B,17,18,^FD" + Convert.ToDateTime(date).ToString("MM/dd/yy") + "^FS\n"; // fecha etiqueta 1     
                                                    impr = impr + "^FO430,170^A0B,13,15,^FDPACK DATE:^FS\n"; // fecha etiqueta 2 
                                                    impr = impr + "^FO445,180^A0B,17,18,^FD" + Convert.ToDateTime(date).ToString("MM/dd/yy") + "^FS\n"; // fecha etiqueta 2
                                                }

                                                impr = impr + "^FO320,170^A0B,25,25,^FD" + lote + "^FS\n";//LOTE ETIQUETA 1
                                                impr = impr + "^FO740,170^A0B,25,25,^FD" + lote + "^FS\n";//LOTE ETIQUETA 2
                                                impr = impr + "^FO320,140^A0N,22,18^FD" + "C: " + auxnum1 + "^FS\n"; // CAJA ETIQUETA 1
                                                impr = impr + "^FO740,140^A0N,22,18^FD" + "C: " + auxnum2 + "^FS\n"; // CAJA ETIQUETA 2
                                                impr = impr + "^FO90,170,^BY1,^BCN,60,N,N,N^MD5F^FD" + aux3 + "^FS\n"; // CODIGO GTIN COD DE BARRAS 128 1
                                                impr = impr + "^FO495,170,^BY1,^BCN,60,N,N,N^FD" + aux3 + "^FS\n"; // CODIGO GTIN COD DE BARRAS 128 2
                                                impr = impr + "^FO110,240^A0N,17,15,^FD" + aux3_texto + "^FS\n"; // DESCRIPCION DEL PRODUCTO EN EL CODIGO BQ LINEA 1 
                                                if (etilote.Length > 0)
                                                {
                                                    impr = impr + "^FO290,145^A0B,17,18,^FD" + etilote + "^FS\n"; // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 1
                                                    impr = impr + "^FO710,145^A0B,17,18,^FD" + etilote + "^FS\n"; // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 2    
                                                }
                                                impr = impr + "^FO350,165^A0B,17,18,^FD" + mfeccad + "^FS\n"; // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 1
                                                impr = impr + "^FO770,165^A0B,17,18,^FD" + mfeccad + "^FS\n"; // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 2    
                                                impr = impr + "^FO510,240^A0N,17,17,^FD" + aux3_texto + "^FS\n"; // DESCRIPCION DEL PRODUCTO EN EL CODIGO BQ LINEA 2
                                                impr = impr + "^XZ\n";

                                            }
                                            #region
                                            /*sw.WriteLine("^XA");
                                    sw.WriteLine("^FO270,20^BQN,2,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS"); //CODIGO 2D BQ DE CUADRO - PTI
                                    sw.WriteLine("^FO680,20^BQN,2,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4_1 + "^FS");  // CODIGO 2D BQ DE CUADRO - PTI
                                    sw.WriteLine("^FO10,15^A0N,22,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS");//etiqueta 1
                                    sw.WriteLine("^FO430,15^A0N,22,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar2 + "/" + totar2 + "^FS");//etiqueta 2

                                    //nombre en ingles etiqueta 1
                                    if (aux2.Length > 25)
                                    {
                                        sw.WriteLine("^FO10,45^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                        sw.WriteLine("^FO10,70^A0N,22,18^FD" + aux2.Substring(26, 24) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                    }
                                    else
                                        sw.WriteLine("^FO10,45^A0N,22,18^FD" + aux2 + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1

                                    //nombre en español etiqueta 1
                                    if (aux5.Length > 25)
                                    {
                                        sw.WriteLine("^FO10,110^A0N,22,18^FD" + aux5.Substring(0, 25) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                        sw.WriteLine("^FO10,133^A0N,22,18^FD" + aux5.Substring(26, 19) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                    }
                                    else
                                        sw.WriteLine("^FO10,110^A0N,22,18^FD" + aux5 + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1

                                    //nombre en ingles etiqueta 2
                                    if (aux2.Length > 25)
                                    {
                                        sw.WriteLine("^FO430,45^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 2
                                        sw.WriteLine("^FO430,70^A0N,22,18^FD" + aux2.Substring(26, 24) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 2
                                    }
                                    else
                                        sw.WriteLine("^FO430,45^A0N,22,18^FD" + aux2 + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 2

                                    //nombre en español etiqueta 2
                                    if (aux5.Length > 25)
                                    {
                                        sw.WriteLine("^FO430,110^A0N,22,18^FD" + aux5.Substring(0, 25) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 2
                                        sw.WriteLine("^FO430,133^A0N,22,18^FD" + aux5.Substring(26, 19) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 2
                                    }
                                    else
                                        sw.WriteLine("^FO430,110^A0N,22,18^FD" + aux5 + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 2

                                    sw.WriteLine("^FO320,140^A0N,22,18^FD" + "C: " + auxnum1 + "^FS"); // CAJA ETIQUETA 1
                                    sw.WriteLine("^FO740,140^A0N,22,18^FD" + "C: " + auxnum2 + "^FS"); // CAJA ETIQUETA 2
                                    sw.WriteLine("^FO90,160,^BY1,^BCN,60,N,N,N^MD5F^FD" + aux3 + "^FS"); // CODIGO GTIN COD DE BARRAS 128 1
                                    sw.WriteLine("^FO495,160,^BY1,^BCN,60,N,N,N^FD" + aux3 + "^FS"); // CODIGO GTIN COD DE BARRAS 128 2
                                    sw.WriteLine("^FO110,230^A0N,17,15,^FD" + aux3_texto + "^FS"); // DESCRIPCION DEL PRODUCTO EN EL CODIGO BQ LINEA 1                            
                                    sw.WriteLine("^FO350,165^A0B,17,18,^FD" + mfeccad + "^FS"); // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 1
                                    sw.WriteLine("^FO770,165^A0B,17,18,^FD" + mfeccad + "^FS"); // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 2    
                                    sw.WriteLine("^FO510,230^A0N,17,17,^FD" + aux3_texto + "^FS"); // DESCRIPCION DEL PRODUCTO EN EL CODIGO BQ LINEA 2
                                    sw.WriteLine("^XZ");*/
                                            #endregion
                                        }
                                        else
                                        {
                                            if (pais_origen.Trim() != "")
                                            {
                                                //COMENTADO 08/12/2016
                                                //impr2 += "^FWB^\n^XA\n";
                                                //impr2 += "^FO05,55^A2N,23,17,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";
                                                //impr2 += "^FO15,100^A0N,21,15^FDC:" + i.ToString() + "^FS\n";
                                                //impr2 += "^FO200,100^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                                //impr2 += "^FO15,250^A0N,48,24^FD" + aux2 + "^FS\n";
                                                //impr2 += "^FO15,300^A0N,48,24^FD" + aux5 + "^FS\n";
                                                //impr2 += "^FO15,350^A0N,25,25^FDPRODUCT OF " + pais_origen + "^FS\n";
                                                //impr2 += "^FO250,350^A0N,25,25^FDHEB US#1^FS\n";
                                                //impr2 += "^FO50,380,^BY2,3,^BCN,90,N,N,N^FD01" + aux3 + "^FS\n";
                                                //impr2 += "^FO150,480^A0N,25,25^FD" + aux3_texto + "^FS\n";
                                                //impr2 += "^FO15,520^A0N,25,25^FDPACKED BY: COMERCIALIZADORA GAB, S.A. DE C.V.^FS\n";
                                                //impr2 += "^FO15,550^A0N,25,25^FDCARRETERA PANAMERICANA KM. 291 - 1^FS\n";
                                                //impr2 += "^FO15,580^A0N,25,25^FDCOL. LA FORTALEZA CORTAZAR, GTO. C.P. 38300^FS\n";
                                                //impr2 += "^FO15,610^A0N,25,25^FDMEXICO R.F.C. CGA-960614-2C5^FS\n";
                                                //impr2 += "^XZ\n";
                                                impr2 += "^XA\n";
                                                impr2 += "~SD12\n"; //~SD15
                                                impr2 += "^PW832\n";
                                                impr2 += "^FO20,15^A2N,23,17,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";
                                                impr2 += "^FO622,200^A0N,21,15^FDC:" + i.ToString() + "^FS\n";
                                                impr2 += "^FO622,30^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                                impr2 += "^FO20,50^A0N,48,24^FD" + aux2 + "^FS\n";
                                                impr2 += "^FO20,100^A0N,48,24^FD" + aux5 + "^FS\n";
                                                impr2 += "^FO20,150^A0N,25,25^FDPRODUCT OF " + pais_origen + "^FS\n";
                                                //impr2 += "^FO300,150^A0N,25,25^FDHEB US#1^FS\n";//AQUI
                                                impr2 += "^FO100,180,^BY2,3,^BCN,90,N,N,N^FD01" + aux3 + "^FS\n";
                                                impr2 += "^FO150,275^A0N,25,25^FD" + aux3_texto + "^FS\n";
                                                impr2 += "^FO20,305^A0N,20,20^FDPACKED BY: COMERCIALIZADORA GAB, S.A. DE C.V.^FS\n";
                                                impr2 += "^FO20,325^A0N,20,20^FDCARRETERA PANAMERICANA KM. 291 - 1^FS\n";
                                                impr2 += "^FO20,345^A0N,20,20^FDCOL. LA FORTALEZA CORTAZAR, GTO. C.P. 38300^FS\n";
                                                impr2 += "^FO20,365^A0N,20,20^FDMEXICO R.F.C. CGA-960614-2C5^FS\n";
                                                impr2 += "^XZ\n";
                                            }

                                            else
                                            {
                                                impr = impr + "^XA\n";
                                                impr += "~SD12\n"; //~SD15
                                                impr += "^PW832\n";
                                                impr = impr + "^FO270,20^BQN,2,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n"; //CODIGO 2D BQ DE CUADRO - PTI"                                            
                                                impr = impr + "^FO-15,20^A0N,22,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";//etiqueta 1

                                                //nombre en ingles etiqueta 1
                                                if (aux2.Length > 25)
                                                {
                                                    impr = impr + "^FO-15,55^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                    impr = impr + "^FO-15,80^A0N,22,18^FD" + aux2.Substring(25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                }
                                                else
                                                    impr = impr + "^FO-15,55^A0N,22,18^FD" + aux2 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1

                                                //nombre en español etiqueta 1
                                                if (aux5.Length > 25)
                                                {
                                                    impr = impr + "^FO-15,120^A0N,22,18^FD" + aux5.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                    impr = impr + "^FO-15,143^A0N,22,18^FD" + aux5.Substring(25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                }
                                                else
                                                    impr = impr + "^FO-15,120^A0N,22,18^FD" + aux5 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1

                                                if ((clave_gab.Trim() == "18007JI56V") || (clave_gab.Trim() == "18JIBOML66") || (clave_gab.Trim() == "16001TO561") ||
                                           (clave_gab.Trim() == "18007JI55M") || (clave_gab == "18007JVM55") || (clave_gab == "18007JVM66") ||
                                           (clave_gab.Trim() == "18007JI56M") || (clave_gab == "18007JVM56") || (clave_gab == "18JIML6610"))
                                                {
                                                    impr = impr + "^FO-15,175^A0B,17,18,^FD" + grado + "^FS\n"; // SE IMPRIME TIPO GRADO 1
                                                    impr = impr + "^FO30,195^A1B,25,20,^FD" + num + "^FS\n"; // SE IMPRIME TIPO GRADO 1                                             
                                                }

                                                if (clave_gab.Trim() == "05006MLNA2")
                                                {
                                                    impr = impr + "^FO-15,175^A0B,17,18,^FD" + Convert.ToDateTime(fec).ToString("MM/dd/yy") + "^FS\n"; // fecha etiqueta 1                                                                                         
                                                }

                                                impr = impr + "^FO320,140^A0N,22,18^FD" + "C: " + auxnum1 + "^FS\n"; // CAJA ETIQUETA 1                                            
                                                impr = impr + "^FO320,170^A0B,25,25,^FD" + lote + "^FS\n";//LOTE ETIQUETA 1
                                                impr = impr + "^FO90,170,^BY1,^BCN,60,N,N,N^MD5F^FD" + aux3 + "^FS\n"; // CODIGO GTIN COD DE BARRAS 128 1                                            
                                                impr = impr + "^FO110,240^A0N,17,15,^FD" + aux3_texto + "^FS\n"; // DESCRIPCION DEL PRODUCTO EN EL CODIGO BQ LINEA 1                            
                                                impr = impr + "^FO350,165^A0B,17,18,^FD" + mfeccad + "^FS\n"; // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 1                                            
                                                impr = impr + "^XZ\n";

                                                #region
                                                /*sw.WriteLine("^XA");
                                    sw.WriteLine("^FO270,20^BQN,2,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS"); //CODIGO 2D BQ DE CUADRO - PTI                                    
                                    sw.WriteLine("^FO10,15^A0N,22,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS");//etiqueta 1
                                    //sw.WriteLine("^FO10,45^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                    //sw.WriteLine("^FO10,70^A0N,22,18^FD" + aux2.Substring(26, 24) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1

                                    //nombre en ingles etiqueta 1
                                    if (aux2.Length > 25)
                                    {
                                        sw.WriteLine("^FO10,45^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                        sw.WriteLine("^FO10,70^A0N,22,18^FD" + aux2.Substring(26, 24) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                    }
                                    else
                                        sw.WriteLine("^FO10,45^A0N,22,18^FD" + aux2 + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1

                                    //nombre en español etiqueta 1
                                    if (aux5.Length > 25)
                                    {
                                        sw.WriteLine("^FO10,110^A0N,22,18^FD" + aux5.Substring(0, 25) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                        sw.WriteLine("^FO10,133^A0N,22,18^FD" + aux5.Substring(26, 19) + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1
                                    }
                                    else
                                        sw.WriteLine("^FO10,110^A0N,22,18^FD" + aux5 + "^FS"); // DESCRIPCION DEL PRODUCTO LINEA 1                     

                                    sw.WriteLine("^FO320,140^A0N,22,18^FD" + "C: " + auxnum1 + "^FS"); // caja etiqueta 1
                                    sw.WriteLine("^FO90,160,^BY1,^BCN,60,N,N,N^MD5F^FD" + aux3 + "^FS"); // CODIGO GTIN COD DE BARRAS 128 1
                                    sw.WriteLine("^FO110,230^A0N,17,15,^FD" + aux3_texto + "^FS"); // DESCRIPCION DEL PRODUCTO EN EL CODIGO BQ LINEA 1                                    
                                    sw.WriteLine("^FO350,165^A0B,17,18,^FD" + mfeccad + "^FS"); // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 1
                                    sw.WriteLine("^XZ");*/
                                                #endregion
                                            }
                                        }
                                        //#region etiqueta 4" x 2"
                                        //imp100x50 += "^XA\n";
                                        //imp100x50 += "^FO645,30^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                        //imp100x50 += "^FO15,25^A2N,23,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";
                                        //imp100x50 += "^FO15,75^A0N,48,24^FD" + aux2 + "^FS\n";
                                        //imp100x50 += "^FO15,135^A0N,48,24^FD" + aux5 + "^FS\n";
                                        //imp100x50 += "^FO15,193^A0N,25,25^FDPRODUCT OF MEXICO^FS\n";
                                        //imp100x50 += "^FO730,235^A0B,35,28,^FD" + mfeccad + "^FS\n";
                                        //imp100x50 += "^FO740,193^A0N,21,15^FDC:" + i.ToString() + "^FS\n";
                                        //imp100x50 += "^FO80,220,^BY2,^BCN,90,N,N,N^FD01" + aux3 + "10" + txtrecibo.Text + "^FS\n";
                                        //imp100x50 += "^FO120,330^A2N,17,17,^FD" + aux3_texto + "(10)" + txtrecibo.Text + "^FS\n";
                                        //imp100x50 += "^FO20,350^A0N,20,20^FDRomaine Grown in Guanajuato / " + txtrchtbl.Text.Trim() + " ^FS\n";
                                        //imp100x50 += "^FO20,370^A0N,20,20^FDHarvested On: " + FechaEla.ToString("dd MMM yy", CultureInfo.CreateSpecificCulture("en-US")).ToUpper() + "^FS\n";
                                        //imp100x50 += "^XZ\n";
                                        //#endregion
                                        j++;
                                    }
                                }
                                #endregion

                                #region etiqueta 4" x 2"
                                if (ConLechuga == "S" && clave_gab.Trim() != "05003ML3P" && clave_gab.Trim() == "03001ML09")
                                // 23 may 23 RCC if (ConLechuga == "S" && clave_gab.Trim() != "05003ML3P") // && clave_gab.Trim() == "03001ML09")
                                {
                                    imp100x50 += "^XA\n";
                                    imp100x50 += "~SD12\n"; //~SD15
                                    imp100x50 += "^PW832\n";
                                    imp100x50 += "^FO635,30^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                    imp100x50 += "^FO15,25^A2N,23,18,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";
                                    imp100x50 += "^FO15,75^A0N,48,24^FD" + aux2 + "^FS\n";
                                    imp100x50 += "^FO15,135^A0N,48,24^FD" + aux5 + "^FS\n";
                                    imp100x50 += "^FO15,191^A0N,25,25^FDPRODUCT OF MEXICO^FS\n";
                                    imp100x50 += "^FO650,191^A0N,25,25,^FD" + mfeccad + "^FS\n";
                                    imp100x50 += "^FO740,193^A0N,21,15^FDC:" + i.ToString() + "^FS\n";
                                    imp100x50 += "^FO80,220,^BY2,^BCN,90,N,N,N^FD01" + aux3 + "10" + txtrecibo.Text + "^FS\n";
                                    imp100x50 += "^FO120,328^A2N,17,17,^FD" + aux3_texto + "(10)" + txtrecibo.Text + "^FS\n";
                                    imp100x50 += "^FO20,350^A0N,20,20^FDRomaine Grown in / Origen Guanajuato - " + txtrchtbl.Text.Trim() + " ^FS\n";
                                    imp100x50 += "^FO20,370^A0N,20,20^FDHarvested On / Cosechado : " + FechaEla.ToString("dd MMM yy", CultureInfo.CreateSpecificCulture("en-US")).ToUpper() + "^FS\n";
                                    imp100x50 += "^XZ\n";
                                }
                                if (clave_gab.Trim() == "05003ML3P" || clave_gab.Trim() == "03001ML09") // ETIQUETA DE COSTCO LECHUGA MR.LUCKY 3 PZAS.
                                                                                                        // 23 may 23 rcc if (clave_gab.Trim() == "05003ML3P" ) // ETIQUETA DE COSTCO LECHUGA MR.LUCKY 3 PZAS.
                                {
                                    imp100x50 += "^XA\n";
                                    imp100x50 += "~SD12\n"; //~SD15
                                    imp100x50 += "^PW832\n";
                                    if (clave_gab.Trim().Length > 25)
                                    {
                                        imp100x50 += "^FO10,20^A0N,80,60^FD" + aux2 + "^FS\n";
                                        imp100x50 += "^FO10,95^A0N,80,60^FD" + aux5 + "^FS\n";
                                    }
                                    else
                                        imp100x50 += "^FO10,20^A0N,80,60^FD" + aux2 + "^FS\n";
                                    imp100x50 += "^FO55,170^BQN,3,3^FDLA," + "http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + txtrecibo.Text + clave_gab + nutar1.Trim().PadLeft(2, '0') + totar1.Trim().PadLeft(2, '0') + j.ToString().Trim().PadLeft(3, '0') + "^FS\n"; // GENERA EL CODIGO QR
                                                                                                                                                                                                                                                                                           //imp100x50 += "^FO55,170^BQN,3,3^FDLA," + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n"; // GENERA EL CODIGO QR
                                    if (CodigoBar.Trim().Length > 0)
                                        imp100x50 += "^FO540,330,^BY2,^BUN,50,Y,N,Y^FD" + CodigoBar + "^FS\n";
                                    imp100x50 += "^FO300,180^A0N,30,30^FDLOTE:^FS\n";
                                    imp100x50 += "^FO300,210^A0N,40,40^FD" + txtrecibo.Text.Trim() + "^FS\n";
                                    imp100x50 += "^FO300,250^A0N,15,15^FD" + "T: " + nutar1.Trim().PadLeft(3, '0') + "/" + j.ToString().Trim().PadLeft(3, '0') + "^FS\n";
                                    //imp100x50 += "^FO500,170^A0N,60,60^FDCons. Pref.^FS\n";
                                    //imp100x50 += "^FO500,230^GB280,90,90^FS\n";
                                    string tmp = "", FCad = "";
                                    if (mfeccad.Trim().Length > 0)
                                    {
                                        //tmp = mfeccad.Substring(mfeccad.Trim().Length - 5);
                                        //FCad = tmp.Substring(0, 3) + " " + tmp.Substring(3);
                                        tmp = mfeccad.Substring(mfeccad.Trim().Length - 7);
                                        FCad = tmp.Substring(2, 3) + " " + tmp.Substring(0, 2);
                                        imp100x50 += "^FO500,170^A0N,60,60^FDCons. Pref.^FS\n";
                                        imp100x50 += "^FO500,230^GB280,90,90^FS\n";
                                        imp100x50 += "^FO520,240^A0N,90,80^FR^FD" + FCad + "^FS\n";
                                    }
                                    //imp100x50 += "^FO520,240^A0N,90,80^FR^FD" + FCad + "^FS\n";
                                    imp100x50 += "^FO15,310^A0N,25,25^FDProducto de Mexico^FS\n";
                                    imp100x50 += "^FO15,340^A0N,20,20^FDDistribuido por: Comercializadora GAB SA de CV^FS\n";
                                    imp100x50 += "^FO15,360^A0N,20,20^FDCarretera Panamericana Km 291-1^FS\n";
                                    imp100x50 += "^FO15,380^A0N,20,20^FDCortazar, Guanajuato., Mexico^FS\n";
                                    imp100x50 += "^XZ\n";
                                    //if (ConLechuga == "S")
                                    ConLechuga = "S";
                                    //EtiCosto = "S";
                                }
                                #endregion
                                //}
                            }
                            //if (eti_grande == "S")
                            //{
                            //    sw.WriteLine("^XA");
                            //    sw.WriteLine("^FO80,30^A0N,45,20^FD -     -      -    -     -   -^FS");// DESCRIPCION DEL PRODUCTO LINEA 2
                            //    //sw.WriteLine("^FO460,30^A0N,45,20^FD -     -      -    -     -   -^FS");// DESCRIPCION DEL PRODUCTO LINEA 1
                            //    sw.WriteLine("^FO60,30,^BY3,^BCN,80,N,N,N^FD00000^XZ");// CODIGO GTIN COD DE BARRAS 128                    
                            //}
                            //else
                            //{
                            if (ConLechuga != "S")
                            {
                                if (impr.Trim() != "")
                                {
                                    impr = impr + "^XA\n";
                                    impr = impr + "^FO80,30^A0N,45,20^FD -     -      -    -     -   -^FS\n";
                                    impr = impr + "^FO460,30^A0N,45,20^FD -     -      -    -     -   -^FS\n";
                                    impr = impr + "^FO60,30,^BY3,^BCN,80,N,N,N^FD00000^XZ\n";
                                }
                                else
                                {
                                    impr2 = impr2 + "^XA\n";
                                    impr2 = impr2 + "^FO80,30^A0N,45,20^FD -     -      -    -     -   -^FS\n";
                                    impr2 = impr2 + "^FO460,30^A0N,45,20^FD -     -      -    -     -   -^FS\n";
                                    impr2 = impr2 + "^FO60,30,^BY3,^BCN,80,N,N,N^FD00000^XZ\n";
                                }
                            }
                            //if (pais_origen.Trim() == "")
                            //{
                            //    impr = impr + "^XA";
                            //    impr = impr + "^FO80,30^A0N,45,20^FD -     -      -    -     -   -^FS";
                            //    impr = impr + "^FO460,30^A0N,45,20^FD -     -      -    -     -   -^FS";
                            //    impr = impr + "^FO60,30,^BY3,^BCN,80,N,N,N^FD00000^XZ";
                            //}
                            //else
                            //{
                            //    impr2 = impr2 + "^XA\n";
                            //    impr2 = impr2 + "^FO80,30^A0N,45,20^FD -     -      -    -     -   -^FS\n";
                            //    impr2 = impr2 + "^FO460,30^A0N,45,20^FD -     -      -    -     -   -^FS\n";
                            //    impr2 = impr2 + "^FO60,30,^BY3,^BCN,80,N,N,N^FD00000^XZ\n";
                            //}
                            /*sw.WriteLine("^XA");
                            sw.WriteLine("^FO80,30^A0N,45,20^FD -     -      -    -     -   -^FS");// DESCRIPCION DEL PRODUCTO LINEA 2
                            sw.WriteLine("^FO460,30^A0N,45,20^FD -     -      -    -     -   -^FS");// DESCRIPCION DEL PRODUCTO LINEA 1
                            sw.WriteLine("^FO60,30,^BY3,^BCN,80,N,N,N^FD00000^XZ");// CODIGO GTIN COD DE BARRAS 128*/
                            //}
                        }//if
                         //x++;
                         //}//for
                         //sw.Close();
                    }//if reader1 = false
                }//if impr = true
                 //sw.Close();
            }//for dgv5       
            //    sw.Close();
            //}
            reader1.Dispose();
            Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "G", "2.3", txtrecibo.Text, "GENERAR ETIQUETA BLANCA RECIBO " + txtrecibo.Text, "SIPGAB");
            //etiqueta_blanca eti_blanca = new etiqueta_blanca();
            //eti_blanca.trazabilidad = tmp_etiquetas.Clone();
            //foreach (DataRow rw in tmp_etiquetas.Rows)
            //    eti_blanca.trazabilidad.ImportRow(rw);

            //eti_blanca.recibo = txtrecibo.Text;
            //eti_blanca.ShowDialog();
            this.Close();

            //}//for
            //DataSet ds = new DataSet();
            //string query = "select pti_clave, prod_nom_ingles, gtin_clave, prod_clave, prod_nombre, etiqueta, tarima, fecha_cad from tb_det_trazabilidad where recibo = '" + txtrecibo.Text + "'";
            //SqlDataAdapter da = new SqlDataAdapter(query, thisConnection);
            //da.Fill(ds, "trazabilidad");
            //etiqueta_blanca eti_blanca = new etiqueta_blanca();
            //eti_blanca.trazabilidad = ds.Tables["trazabilidad"];            
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "update tb_mstr_recepcion_pt set fin_captura = '" + DateTime.Now.ToString() + "', ind = 'R' where rpt_recibo = '" + txtrecibo.Text + "'";
            reader1 = cmnd1.ExecuteReader();
            thisConnection.Close();
            eti_blanca.rancho = txtrchtbl.Text.Trim();
            eti_blanca.ShowDialog();
        }
        Image newImage1 = null;
        private void printDocument2_PrintPage(object sender, PrintPageEventArgs e)
        {
            string nombre01 = "CLAVE";
            string nombre02 = "Source";
            string nombre03 = "Field";
            string nombre04 = "Product";
            string nombre05 = "Lot #:";
            string nombre06 = "Load #:";
            string nombre07 = "Time in";
            string nombre08 = "of";
            string nombre09 = "Pallets";


            if (i <= tmp_etiquetas.Rows.Count - 1)
            {

                //BarcodeLib.Barcode.Linear code39 = new BarcodeLib.Barcode.Linear();
                //code39.Type = BarcodeLib.Barcode.BarcodeType.CODE39;
                //code39.Data = "CODE39-"+ Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString());
                //code39.N = 3;
                //code39.AddCheckSum = true;
                //code39.UOM = BarcodeLib.Barcode.UnitOfMeasure.PIXEL;
                //code39.drawBarcode("C:/sisgabweb/code39.jpg");
                //newImage1 = Image.FromFile(@"C:/sisgabweb/code39.jpg");
                //MessageBox.Show(Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString()));
                //if (File.Exists(@"C:/sisgabweb/code128.jpg"))
                //    File.Delete(@"C:/sisgabweb/code128.jpg");

                BarcodeLib.Barcode.Linear code128 = new BarcodeLib.Barcode.Linear();
                code128.Type = BarcodeLib.Barcode.BarcodeType.CODE128;
                code128.Data = Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString());
                code128.AddCheckSum = true;
                code128.UOM = BarcodeLib.Barcode.UnitOfMeasure.PIXEL;
                code128.BarHeight = 60;
                code128.BarWidth = 1;
                code128.RightMargin = 10;
                code128.LeftMargin = 10;
                code128.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                code128.drawBarcode("C:/sisgabweb/code128.jpeg");
                newImage1 = Image.FromFile(@"C:/sisgabweb/code128.jpg");


                SolidBrush color = new SolidBrush(Color.Black);
                Font fuente_encabezados = new Font("Arial", 15);
                Font texto = new Font("Arial", 12);
                //Font cod_bar = new Font("PF Barcode 39", 19);//código de barras
                Font cod_bar = new System.Drawing.Font("PF Barcode 128", 35);
                Font ord_pro = new Font("Arial", 18);
                Font nota = new Font("Arial", 10);

                //clave
                Point clave = new Point(40, 15);
                e.Graphics.DrawString(nombre01, fuente_encabezados, color, clave);

                Point txt_clave = new Point(160, 15);
                e.Graphics.DrawString(txtclave.Text, fuente_encabezados, color, txt_clave);

                Pen lapiz = new Pen(Color.Black, 2);
                //Rectangle rect = new Rectangle(269, 10, 70,50);
                Rectangle rect = new Rectangle(264, 10, 73, 48);
                e.Graphics.DrawRectangle(lapiz, rect);

                Point box_pallet = new Point(280, 20);
                e.Graphics.DrawString(Convert.ToString(tmp_etiquetas.Rows[i][1].ToString()), fuente_encabezados, color, box_pallet);

                //source
                Point sr = new Point(30, 60);
                e.Graphics.DrawString(nombre02, fuente_encabezados, color, sr);

                Point prov = new Point(30, 86);
                e.Graphics.DrawString(txtprov.Text, texto, color, prov);

                //field
                Point rch = new Point(30, 116);
                e.Graphics.DrawString(nombre03, fuente_encabezados, color, rch);

                Point ran = new Point(30, 142);
                e.Graphics.DrawString(txtrchtbl.Text, texto, color, ran);

                //producto
                Point pr = new Point(30, 172);
                e.Graphics.DrawString(nombre04, fuente_encabezados, color, pr);

                Point nom_prod = new Point(30, 198);
                string tmp = tmp_etiquetas.Rows[i]["etq_descrip"].ToString();
                if (tmp.Length >= 25)
                {
                    int resta = tmp.Length - 25;
                    string pro = tmp.Substring(0, 25);
                    string pro2 = tmp.Substring(25, resta);
                    Point nom_prod2 = new Point(30, 215);
                    e.Graphics.DrawString(pro, texto, color, nom_prod);
                    e.Graphics.DrawString(pro2, texto, color, nom_prod2);
                }
                else
                    e.Graphics.DrawString(tmp, texto, color, nom_prod);


                //orden de producción = lot
                Point lote = new Point(105, 245);
                e.Graphics.DrawString(nombre05, fuente_encabezados, color, lote);

                Point lt = new Point(170, 245);
                e.Graphics.DrawString(Convert.ToString(tmp_etiquetas.Rows[i]["etq_recibo"].ToString()), ord_pro, color, lt);

                //viaje = load
                Point via = new Point(105, 275);
                e.Graphics.DrawString(nombre06, fuente_encabezados, color, via);

                Point tra = new Point(200, 275);
                e.Graphics.DrawString(Convert.ToString(tmp_etiquetas.Rows[i]["etq_viaje"].ToString()), texto, color, tra);

                //hora = time
                Point tm = new Point(105, 305);
                e.Graphics.DrawString(nombre07, fuente_encabezados, color, tm);

                Point time = new Point(200, 305);
                e.Graphics.DrawString(Convert.ToString(tmp_etiquetas.Rows[i]["etq_hora"].ToString()), texto, color, time);

                //rectangulo de tarimas                
                //Rectangle rect = new Rectangle(269, 10, 73,48);
                Rectangle recta = new Rectangle(80, 350, 250, 30);
                e.Graphics.DrawRectangle(lapiz, recta);

                //inicio etiqueta
                Point ini = new Point(95, 355);
                e.Graphics.DrawString(Convert.ToString(tmp_etiquetas.Rows[i]["etq_ini"].ToString()), texto, color, ini);

                Point of = new Point(150, 355);
                e.Graphics.DrawString(nombre08, fuente_encabezados, color, of);

                //final etiqueta
                Point fin = new Point(200, 355);
                e.Graphics.DrawString(Convert.ToString(tmp_etiquetas.Rows[i]["etq_final"].ToString()), texto, color, fin);

                //pallets = tarima
                Point tar = new Point(240, 355);
                e.Graphics.DrawString(nombre09, fuente_encabezados, color, tar);

                //primer codigo de barras lado verde             
                //int x = 0;
                //int y = 400;
                //int width = 250;
                //int height = 30;                
                //e.Graphics.DrawImage(newImage1, x, y, width, height);
                //Rectangle destRect2 = new Rectangle(0, 400, 250, 30);
                e.Graphics.DrawImage(newImage1, 0, 400);
                //Point cb = new Point(20, 400);
                //e.Graphics.DrawString(Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString()), cod_bar, color, cb);
                //Point cb2 = new Point(20, 420);
                //e.Graphics.DrawString(Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString()), cod_bar, color, cb2);

                //Point cb3 = new Point(10, 440);
                //e.Graphics.DrawString(Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString()), fuente_encabezados, color, cb3);

                //hora de impresion de la etiqueta
                Point hour = new Point(10, 500);
                e.Graphics.DrawString(DateTime.Now.ToString("HH:mm:ss"), nota, color, hour);

                //nota
                Point note = new Point(90, 500);
                e.Graphics.DrawString(tmp_etiquetas.Rows[i]["etq_nota"].ToString(), nota, color, note);

                //segundo código de barras
                e.Graphics.DrawImage(newImage1, 0, 520);
                //Point cb4 = new Point(20, 540);
                //e.Graphics.DrawString(Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString()), cod_bar, color, cb4);
                //Point cb5 = new Point(20, 560);
                //e.Graphics.DrawString(Convert.ToString(tmp_etiquetas.Rows[i]["etq_codigo_barras"].ToString()), cod_bar, color, cb5);
            }
            if (i == tmp_etiquetas.Rows.Count - 1)
            {
                e.HasMorePages = false;
                i = 0;
            }
            else
                e.HasMorePages = true;

            i++;
            newImage1.Dispose();
        }

        private string obtenerNombreMesNumero(int numeroMes)
        {

            DateTimeFormatInfo formatoFecha = CultureInfo.CurrentCulture.DateTimeFormat;
            string nombreMes = formatoFecha.GetMonthName(numeroMes);
            mes = nombreMes.Substring(0, 3);
            return mes;

        }


        //int j = 1; co = 0, renglon_selec = 0; 
        private void eti_blanca_PrintPage(object sender, PrintPageEventArgs e)
        {
            /*if (aux3 == "")
            {
                aux3 = clave_gab;
                aux3_texto = "(01)" + aux3;
            }
            else
                aux3_texto = "(01)" + aux3;

            tot_cajas = tot_cajas + auxn;

            auxc = j.ToString();
            if (auxc.Length == 1)
                auxd = "00" + auxc;
            if (auxc.Length == 2)
                auxd = "0" + auxc;

            aux4 = aux1 + auxd;
            auxnum1 = j.ToString();
            y = j + 1;
            auxnum2 = y.ToString();
            auxc = y.ToString();
            auxd = auxc;
            auxd_1 = auxd;
            if (auxd_1.Length == 1)
                auxd_1 = "00" + auxd;
            if (auxc.Length == 2)
                auxd_1 = "0" + auxd;

            aux4_1 = aux1 + auxd_1;
            lon_cad1 = aux4.Length + 1;
            lon_cad2 = aux4_1.Length + 1;

            nutar1 = aux4.Substring(lon_cad1 - 8, 2);
            totar1 = aux4.Substring(lon_cad1 - 6, 2);
            nutar2 = aux4_1.Substring(lon_cad1 - 8, 2);
            totar2 = aux4_1.Substring(lon_cad1 - 6, 2);

            if (co == 0)
            {
                if ((renglon_selec % 2) != 0)
                {
                    QrEncoder qrEncoder = new QrEncoder(ErrorCorrectionLevel.H);
                    QrCode qrCode = new QrCode();
                    qrEncoder.TryEncode("http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4, out qrCode); //qrcode etiqueta 1

                    GraphicsRenderer renderer = new GraphicsRenderer(new FixedModuleSize(5, QuietZoneModules.Two), Brushes.Black, Brushes.White);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        renderer.WriteToStream(qrCode.Matrix, ImageFormat.Png, ms);
                        var ima = new Bitmap(ms);
                        var imag = new Bitmap(ima);

                        Image newImage = Image.FromStream(ms);
                        e.Graphics.DrawImage(newImage, 125, 3, 60, 60);//qrcode etiqueta 1
                        ms.Close();
                        ms.Dispose();
                        ima.Dispose();
                        imag.Dispose();
                    }

                    e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar1 + "/" + totar1, fuente_encabezados, color, 0, 5);// primera linea etiqueta 1                                
                    if (aux2.Length > 25)
                    {
                        e.Graphics.DrawString(aux2.Substring(0, 24), nota, color, 3, 15);//nombre en ingles parte 1 etiqueta 1
                        e.Graphics.DrawString(aux2.Substring(25), nota, color, 3, 28);//nombre en ingles parte 2 etiqueta 1                                    
                    }
                    else
                    {
                        e.Graphics.DrawString(aux2, nota, color, 3, 15);//nombre en ingles completo etiqueta 1                                    
                    }
                    if (aux5.Length > 25)
                    {
                        e.Graphics.DrawString(aux5.Substring(0, 24), nota, color, 3, 41);//nombre en español parte 1 etiqueta 1
                        e.Graphics.DrawString(aux5.Substring(25), nota, color, 3, 54);//nombre en español parte 2 etiqueta 1                                    
                    }
                    else
                    {
                        e.Graphics.DrawString(aux5, nota, color, 3, 41);//nombre en español completo etiqueta 1                                    
                    }

                    BarcodeLib.Barcode.Linear code128 = new BarcodeLib.Barcode.Linear();
                    code128.Type = BarcodeLib.Barcode.BarcodeType.CODE128;
                    code128.Data = aux3;
                    code128.AddCheckSum = true;
                    code128.UOM = BarcodeLib.Barcode.UnitOfMeasure.PIXEL;
                    code128.BarHeight = 50;
                    code128.BarWidth = 2;
                    //code128.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                    //code128.drawBarcode("C:/sisgabweb/code128.jpg");
                    byte[] code821 = code128.drawBarcodeAsBytes();
                    Stream ms1 = new MemoryStream(code821);
                    //Image newImage1 = Image.FromFile(@"C:/sisgabweb/code128.jpg");
                    Image newImage1 = Image.FromStream(ms1);
                    e.Graphics.DrawImage(newImage1, 0, 80, 150, 30);// codebar 128 etiqueta 1                                

                    StringFormat drawFormat = new StringFormat();
                    drawFormat.FormatFlags = StringFormatFlags.DirectionVertical;
                    e.Graphics.DrawString(mfeccad, cadu, color, 170, 80, drawFormat);//fecha de caducidad etiqueta 1                                

                    e.Graphics.DrawString("C: " + auxnum1, box, color, 150, 67);//número de caja etiqueta 1                                

                    Pen whitepen = new Pen(Color.White, 5);
                    int fin = 105;
                    for (int m = 0; m < 10; m++)
                    {
                        Rectangle recta = new Rectangle(50, fin, 100, 10);
                        e.Graphics.DrawRectangle(whitepen, recta);
                        fin++;
                    }
                    e.Graphics.DrawString("(01)" + aux3, nota, color, 30, 105); //numero codigo de barras etiqueta 1
                    j++;
                }
                else
                {
                    QrEncoder qrEncoder = new QrEncoder(ErrorCorrectionLevel.H);
                    QrCode qrCode = new QrCode();
                    qrEncoder.TryEncode("http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4, out qrCode); //qrcode etiqueta 1

                    GraphicsRenderer renderer = new GraphicsRenderer(new FixedModuleSize(5, QuietZoneModules.Two), Brushes.Black, Brushes.White);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        renderer.WriteToStream(qrCode.Matrix, ImageFormat.Png, ms);
                        var ima = new Bitmap(ms);
                        var imag = new Bitmap(ima);

                        Image newImage = Image.FromStream(ms);
                        e.Graphics.DrawImage(newImage, 125, 5, 60, 60);//qrcode etiqueta 1
                        ms.Close();
                        imag.Dispose();
                    }

                    QrEncoder qrEncoder2 = new QrEncoder(ErrorCorrectionLevel.H);
                    QrCode qrCode2 = new QrCode();
                    qrEncoder2.TryEncode("http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4_1, out qrCode); //qrcode etiqueta 2

                    GraphicsRenderer renderer2 = new GraphicsRenderer(new FixedModuleSize(5, QuietZoneModules.Two), Brushes.Black, Brushes.White);
                    using (MemoryStream ms2 = new MemoryStream())
                    {
                        renderer2.WriteToStream(qrCode.Matrix, ImageFormat.Png, ms2);
                        var ima2 = new Bitmap(ms2);
                        var imag2 = new Bitmap(ima2);

                        Image newImage2 = Image.FromStream(ms2);
                        e.Graphics.DrawImage(newImage2, 330, 3, 60, 60);//qrcode etiqueta 2
                        ms2.Close();
                        imag2.Dispose();
                    }

                    e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar1 + "/" + totar1, fuente_encabezados, color, 3, 5);// primera linea etiqueta 1
                    e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar2 + "/" + totar2, fuente_encabezados, color, 205, 5);// primera linea etiqueta 2
                    if (aux2.Length > 25)
                    {
                        e.Graphics.DrawString(aux2.Substring(0, 24), nota, color, 3, 15);//nombre en ingles parte 1 etiqueta 1
                        e.Graphics.DrawString(aux2.Substring(25), nota, color, 3, 28);//nombre en ingles parte 2 etiqueta 1

                        e.Graphics.DrawString(aux2.Substring(0, 24), nota, color, 205, 15);//nombre en ingles parte 1 etiqueta 2
                        e.Graphics.DrawString(aux2.Substring(25), nota, color, 205, 28);//nombre en ingles parte 2 etiqueta 2
                    }
                    else
                    {
                        e.Graphics.DrawString(aux2, nota, color, 3, 15);//nombre en ingles completo etiqueta 1
                        e.Graphics.DrawString(aux2, nota, color, 205, 15);//nombre en ingles completo etiqueta 2
                    }
                    if (aux5.Length > 25)
                    {
                        e.Graphics.DrawString(aux5.Substring(0, 24), nota, color, 3, 41);//nombre en español parte 1 etiqueta 1
                        e.Graphics.DrawString(aux5.Substring(25), nota, color, 3, 54);//nombre en español parte 2 etiqueta 1

                        e.Graphics.DrawString(aux5.Substring(0, 24), nota, color, 205, 41);//nombre en español parte 1 etiqueta 2
                        e.Graphics.DrawString(aux5.Substring(25), nota, color, 205, 54);//nombre en español parte 2 etiqueta 2
                    }
                    else
                    {
                        e.Graphics.DrawString(aux5, nota, color, 3, 41);//nombre en español completo etiqueta 1
                        e.Graphics.DrawString(aux5, nota, color, 205, 41);//nombre en español completo etiqueta 2
                    }

                    //BarcodeLib.Barcode.Linear code128 = new BarcodeLib.Barcode.Linear();
                    //code128.Type = BarcodeLib.Barcode.BarcodeType.CODE128;
                    //code128.Data = aux3;
                    //code128.AddCheckSum = true;
                    //code128.UOM = BarcodeLib.Barcode.UnitOfMeasure.PIXEL;
                    //code128.BarHeight = 50;
                    //code128.BarWidth = 2;
                    //code128.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                    //code128.drawBarcode("C:/sisgabweb/code128.jpg");
                    //Image newImage1 = Image.FromFile(@"C:/sisgabweb/code128.jpg");
                    //e.Graphics.DrawImage(newImage1, 0, 80, 150, 30);// codebar 128 etiqueta 1
                    Image newImage1 = Image.FromStream(ms1);
                    e.Graphics.DrawImage(newImage1, 0, 80, 150, 30);

                    //BarcodeLib.Barcode.Linear code1282 = new BarcodeLib.Barcode.Linear();
                    //code1282.Type = BarcodeLib.Barcode.BarcodeType.CODE128;
                    //code1282.Data = aux3;
                    //code1282.AddCheckSum = true;
                    //code1282.UOM = BarcodeLib.Barcode.UnitOfMeasure.PIXEL;
                    //code1282.BarHeight = 50;
                    //code1282.BarWidth = 2;
                    //code1282.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                    //code1282.drawBarcode("C:/sisgabweb/code1282.jpg");
                    //Image newImage3 = Image.FromFile(@"C:/sisgabweb/code1282.jpg");
                    e.Graphics.DrawImage(newImage1, 205, 80, 150, 30);// codebar 128 etiqueta 2

                    StringFormat drawFormat = new StringFormat();
                    drawFormat.FormatFlags = StringFormatFlags.DirectionVertical;
                    e.Graphics.DrawString(mfeccad, cadu, color, 170, 80, drawFormat);//fecha de caducidad etiqueta 1

                    StringFormat drawFormat2 = new StringFormat();
                    drawFormat2.FormatFlags = StringFormatFlags.DirectionVertical;
                    e.Graphics.DrawString(mfeccad, cadu, color, 370, 80, drawFormat2);//fecha de caducidad etiqueta 2


                    e.Graphics.DrawString("C: " + auxnum1, box, color, 150, 67);//número de caja etiqueta 1
                    e.Graphics.DrawString("C: " + auxnum2, box, color, 350, 67);//número de caja etiqueta 2

                    Pen whitepen = new Pen(Color.White, 5);
                    int fin = 105;
                    for (int m = 0; m < 10; m++)
                    {
                        Rectangle recta = new Rectangle(50, fin, 100, 10);
                        e.Graphics.DrawRectangle(whitepen, recta);
                        fin++;
                    }
                    e.Graphics.DrawString("(01)" + aux3, nota, color, 30, 105); //numero codigo de barras etiqueta 1

                    //Pen black = new Pen(Color.Black, 5);
                    fin = 105;
                    for (int m = 0; m < 10; m++)
                    {
                        Rectangle recta = new Rectangle(230, fin, 100, 10);
                        e.Graphics.DrawRectangle(whitepen, recta);
                        fin++;
                    }
                    e.Graphics.DrawString("(01)" + aux3, nota, color, 230, 105); //numero codigo de barras etiqueta 2  
                    j += 2;
                }
                co++;
            }
            //else
            if (co == 1)
            {
                QrEncoder qrEncoder = new QrEncoder(ErrorCorrectionLevel.H);
                QrCode qrCode = new QrCode();
                qrEncoder.TryEncode("http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4, out qrCode); //qrcode etiqueta 1

                GraphicsRenderer renderer = new GraphicsRenderer(new FixedModuleSize(5, QuietZoneModules.Two), Brushes.Black, Brushes.White);
                using (MemoryStream ms = new MemoryStream())
                {
                    renderer.WriteToStream(qrCode.Matrix, ImageFormat.Png, ms);
                    var ima = new Bitmap(ms);
                    var imag = new Bitmap(ima);

                    Image newImage = Image.FromStream(ms);
                    e.Graphics.DrawImage(newImage, 125, 5, 60, 60);//qrcode etiqueta 1
                    ms.Close();
                    imag.Dispose();
                }

                QrEncoder qrEncoder2 = new QrEncoder(ErrorCorrectionLevel.H);
                QrCode qrCode2 = new QrCode();
                qrEncoder2.TryEncode("http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4_1, out qrCode); //qrcode etiqueta 2

                GraphicsRenderer renderer2 = new GraphicsRenderer(new FixedModuleSize(5, QuietZoneModules.Two), Brushes.Black, Brushes.White);
                using (MemoryStream ms2 = new MemoryStream())
                {
                    renderer2.WriteToStream(qrCode.Matrix, ImageFormat.Png, ms2);
                    var ima2 = new Bitmap(ms2);
                    var imag2 = new Bitmap(ima2);

                    Image newImage2 = Image.FromStream(ms2);
                    e.Graphics.DrawImage(newImage2, 330, 3, 60, 60);//qrcode etiqueta 2
                    ms2.Close();
                    imag2.Dispose();
                }

                e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar1 + "/" + totar1, fuente_encabezados, color, 3, 5);// primera linea etiqueta 1
                e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar2 + "/" + totar2, fuente_encabezados, color, 205, 5);// primera linea etiqueta 2
                if (aux2.Length > 25)
                {
                    e.Graphics.DrawString(aux2.Substring(0, 24), nota, color, 3, 15);//nombre en ingles parte 1 etiqueta 1
                    e.Graphics.DrawString(aux2.Substring(25), nota, color, 3, 28);//nombre en ingles parte 2 etiqueta 1

                    e.Graphics.DrawString(aux2.Substring(0, 24), nota, color, 205, 15);//nombre en ingles parte 1 etiqueta 2
                    e.Graphics.DrawString(aux2.Substring(25), nota, color, 205, 28);//nombre en ingles parte 2 etiqueta 2
                }
                else
                {
                    e.Graphics.DrawString(aux2, nota, color, 3, 15);//nombre en ingles completo etiqueta 1
                    e.Graphics.DrawString(aux2, nota, color, 205, 15);//nombre en ingles completo etiqueta 2
                }
                if (aux5.Length > 25)
                {
                    e.Graphics.DrawString(aux5.Substring(0, 24), nota, color, 3, 41);//nombre en español parte 1 etiqueta 1
                    e.Graphics.DrawString(aux5.Substring(25), nota, color, 3, 54);//nombre en español parte 2 etiqueta 1

                    e.Graphics.DrawString(aux5.Substring(0, 24), nota, color, 205, 41);//nombre en español parte 1 etiqueta 2
                    e.Graphics.DrawString(aux5.Substring(25), nota, color, 205, 54);//nombre en español parte 2 etiqueta 2
                }
                else
                {
                    e.Graphics.DrawString(aux5, nota, color, 3, 41);//nombre en español completo etiqueta 1
                    e.Graphics.DrawString(aux5, nota, color, 205, 41);//nombre en español completo etiqueta 2
                }

                //BarcodeLib.Barcode.Linear code128 = new BarcodeLib.Barcode.Linear();
                //code128.Type = BarcodeLib.Barcode.BarcodeType.CODE128;
                //code128.Data = aux3;
                //code128.AddCheckSum = true;
                //code128.UOM = BarcodeLib.Barcode.UnitOfMeasure.PIXEL;
                //code128.BarHeight = 50;
                //code128.BarWidth = 2;
                //code128.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                //code128.drawBarcode("C:/sisgabweb/code128.jpg");
                Image newImage1 = Image.FromFile(@"C:/sisgabweb/code128.jpg");
                e.Graphics.DrawImage(newImage1, 0, 80, 150, 30);// codebar 128 etiqueta 1

                //BarcodeLib.Barcode.Linear code1282 = new BarcodeLib.Barcode.Linear();
                //code1282.Type = BarcodeLib.Barcode.BarcodeType.CODE128;
                //code1282.Data = aux3;
                //code1282.AddCheckSum = true;
                //code1282.UOM = BarcodeLib.Barcode.UnitOfMeasure.PIXEL;
                //code1282.BarHeight = 50;
                //code1282.BarWidth = 2;
                //code1282.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                //code1282.drawBarcode("C:/sisgabweb/code1282.jpg");
                //Image newImage3 = Image.FromFile(@"C:/sisgabweb/code1282.jpg");
                e.Graphics.DrawImage(newImage1, 205, 80, 150, 30);// codebar 128 etiqueta 2

                StringFormat drawFormat = new StringFormat();
                drawFormat.FormatFlags = StringFormatFlags.DirectionVertical;
                e.Graphics.DrawString(mfeccad, cadu, color, 170, 80, drawFormat);//fecha de caducidad etiqueta 1

                StringFormat drawFormat2 = new StringFormat();
                drawFormat2.FormatFlags = StringFormatFlags.DirectionVertical;
                e.Graphics.DrawString(mfeccad, cadu, color, 370, 80, drawFormat2);//fecha de caducidad etiqueta 2


                e.Graphics.DrawString("C: " + auxnum1, box, color, 150, 67);//número de caja etiqueta 1
                e.Graphics.DrawString("C: " + auxnum2, box, color, 350, 67);//número de caja etiqueta 2

                Pen whitepen = new Pen(Color.White, 5);
                int fin = 105;
                for (int m = 0; m < 10; m++)
                {
                    Rectangle recta = new Rectangle(50, fin, 100, 10);
                    e.Graphics.DrawRectangle(whitepen, recta);
                    fin++;
                }
                e.Graphics.DrawString("(01)" + aux3, nota, color, 30, 105); //numero codigo de barras etiqueta 1

                fin = 105;
                for (int m = 0; m < 10; m++)
                {
                    Rectangle recta = new Rectangle(230, fin, 100, 10);
                    e.Graphics.DrawRectangle(whitepen, recta);
                    fin++;
                }
                e.Graphics.DrawString("(01)" + aux3, nota, color, 230, 105); //numero codigo de barras etiqueta 2  
                j += 2;
            }

            if (co == 2)
            {
                co = 0;
                e.Graphics.DrawString("|    |   |   |   |", nota, color, 230, 105); //numero codigo de barras etiqueta 1
                e.Graphics.DrawString("|    |   |   |   |", nota, color, 3, 105); //numero codigo de barras etiqueta 2  
            }

            if (j > renglon_selec)
            {
                if (co == 0)
                {
                    e.HasMorePages = false;
                    j = 0;
                    co = 0;
                }
                else
                {
                    e.HasMorePages = true;
                    co = 2;
                }
            }
            else
            {
                e.HasMorePages = true;
            }                                                                                                     */
        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (cbgrado.SelectedIndex == 0)
            {
                grado = "GRADO";
                num = "2";
            }
            if (cbgrado.SelectedIndex == 1)
            {
                grado = "GRADO";
                num = "3";
            }
            if (cbgrado.SelectedIndex == 2)
            {
                grado = "GRADO";
                num = "4";
            }

            for (int i = 0; i < DGV4.Rows.Count; i++)
            {
                if (fila.Trim() == Convert.ToString(DGV4.Rows[i].Cells["prod_clave"].Value).Trim())
                    DGV4.Rows[i].Cells["grad"].Value = grado;
            }
            DGV4.Refresh();
            btnactu.Focus();
        }

        private string TrarItem(string CveProd)
        {
            string Item = "";
            //thisConnection.Open();
            SqlCommand cmd = new SqlCommand("Select prod_clave_gabinc from tb_cat_producto where prod_clave = '" + CveProd + "'", thisConnection);
            Item = Convert.ToString(cmd.ExecuteScalar());
            //thisConnection.Close();
            return Item;

        }

        private void DGV4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DGV4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 5)
            {
                string x = DGV4.Rows[e.RowIndex].Cells[5].Value.ToString();
                if (x == "False" && existetrazabilidad == "")
                {
                    DGV4.Rows[e.RowIndex].Cells[5].Value = true;
                }
            }
        }

        private void DGV5_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 7)
            {
                string x = DGV5.Rows[e.RowIndex].Cells[7].Value.ToString();
                if (x == "False" && existetrazabilidad == "")
                {
                    DGV5.Rows[e.RowIndex].Cells[7].Value = true;
                }
            }
        }

        public string getPesoXCaja(string recibo, string prodclave)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Utilerias.Class1.ConnectionString))
                {
                    using (SqlCommand command = new SqlCommand("spGetPesoXCaja", connection))
                    {
                        // Configurar el comando como un procedimiento almacenado
                        command.CommandType = CommandType.StoredProcedure;

                        // Agregar el parámetro del procedimiento
                        command.Parameters.AddWithValue("@Recibo", recibo);
                        command.Parameters.AddWithValue("@ProdClave", prodclave);

                        // Abrir la conexión
                        connection.Open();


                        // Ejecutar el comando y leer los resultados
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                // Obtener el valor del peso unitario
                                return reader["PesoUnitario"].ToString();
                            }
                            else
                            {
                                return "0.00";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Manejo de errores
                //Console.WriteLine($"Error al obtener el peso unitario: {ex.Message}");
                return "0.0";
            }
        }
    }
}
