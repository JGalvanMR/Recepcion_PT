using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;
using System.Data.OleDb;
using System.Drawing.Printing;
using System.Text.RegularExpressions;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Diagnostics;
using System.Threading;
using System.Globalization;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Recepcion_PT
{
    public partial class Recepcion_PT : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlCommand cmnd1 = new SqlCommand();
        SqlCommand cmnd11 = new SqlCommand();
        SqlCommand cmnd111 = new SqlCommand();
        SqlCommand cmnd1111 = new SqlCommand();
        SqlDataReader reader1, reader11, reader111, reader1111;
        SqlDataAdapter da;
        DataSet ds = new DataSet();

        string Actualizar = "N";

        MySqlConnection mySqlConn = new MySqlConnection("server=gab.mrlucky.com.mx;userid=www1166;password=taQ17Zm;database=campo");
        MySqlCommand cmnd = new MySqlCommand();
        MySqlDataReader reader;
        MySqlDataReader readr1;

        //OleDbConnection MyConnection = new OleDbConnection(@"Provider= VFPOLEDB.1;Data Source=c:\mr_lucky\base_de_datos;Collating Sequence=general;");
        OleDbConnection MyConnection = new OleDbConnection(Utilerias.Class1.ConnectionStringFox);
        OleDbCommand cmd1 = new OleDbCommand();
        //OleDbDataReader read1, read11, read111, read1111;

        public static int opcion = 0, recibo = 0, cantidad = 0, tarm = 0, rpt_recibo, num_prod, tam, coor_x, num_viaje, folio_no_conformidad = 0;
        public static string varieda = "", producto = "", nombre = "", fecha_cad = "", envas = "", provee = "", rancho = "", tabla = "", linea = "", nom_tabla = "", tipo = "",
                             env_clave = "", situacion = "S", cletiqueta, proveedor, rancho_nom, hora, estatus, tipo_notcre = "", folnotcre = "", flt = "", query = "", pais = "",
                             email = "", tipo_recepcion = "", evaluacion = "", ruta_pdf = "", contrato = "", HrCaptura = "", fcn_folioNCR = "";
        public static DateTime FechaEla;
        public static decimal b_p = 0, tar = 0, p_u = 0, p_t = 0, peso_env, invtp_entradas_kg = 0, invpt_entradas_un = 0, invpt_inicial_kg = 0, invpt_salidas_kg = 0, invpt_inicial_un = 0, invpt_salidas_un = 0;
        public static decimal var_dec_inv_inicial_kg = 0, var_dec_inv_inicial_un = 0, var_dec_peso_total = 0, invpt_entradas_kg = 0, hrp_num_unidades, hrp_peso_neto, hrp_reman_unidades;

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            txtrecibo.Text = cbrecpen.SelectedItem.ToString().Trim();
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "select rpt_estatus from tb_mstr_recepcion_pt where rpt_recibo = '" + txtrecibo.Text.Trim() + "'";
            string st = Convert.ToString(cmnd1.ExecuteScalar()).Trim();
            thisConnection.Close();
            if (st.Trim() == "C")
            {
                MessageBox.Show("No se pueden imprimir las etiquetas debido a que el recibo esta cancelado", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            txtrecibo_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
        }

        public static bool notacre, completo;
        public static string ini_captura = "", fin_captura = "";
        ComboBox ticket_bascula = new ComboBox();
        ComboBox prod_clave = new ComboBox();
        ComboBox variedad = new ComboBox();
        ComboBox varie = new System.Windows.Forms.ComboBox();
        ComboBox env = new System.Windows.Forms.ComboBox();
        ComboBox nume_prod = new System.Windows.Forms.ComboBox();
        ComboBox nom_prod = new System.Windows.Forms.ComboBox();

        DataTable proveedores = new DataTable("proveedores");
        DataTable ranchos = new DataTable("ranchos");
        DataTable tablas = new DataTable("tablas");
        DataTable subtablas = new DataTable("subtablas");
        DataTable lineas = new DataTable("lineas");
        DataTable variedades = new DataTable("variedades");
        DataTable envases = new DataTable("envases");
        DataTable procesos = new DataTable("procesos");
        DataTable danos = new DataTable("danos");
        DataTable productos = new DataTable("productos");
        DataTable not_cre = new DataTable("not_cre");
        DataTable campo = new DataTable("campo");

        public Recepcion_PT()
        {
            InitializeComponent();
            not_cre.Columns.Add("producto");
            not_cre.Columns.Add("cantidad");

            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);

            string carpeta = @"C:\\no_conformidad";
            if (!(Directory.Exists(carpeta)))
            {
                Directory.CreateDirectory(carpeta);
            }

            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT COUNT(*) FROM TB_TMP_MSTR_NOTA WHERE AUTORIZADO = 'S' AND AFECTADO = ' '";
            reader1 = cmnd1.ExecuteReader();
            while (reader1.Read())
            {
                if (reader1.GetValue(0).ToString().Trim() != "0")
                {
                    MessageBox.Show("TIENE " + reader1.GetValue(0).ToString().Trim() + " MOVIMIENTOS PENDIENTES POR INGRESAR DE LAS NOTAS DE CREDITO", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            reader1.Dispose();
            thisConnection.Close();
            label9.Text = File.GetLastWriteTime(@"C:\sisgabweb\Recepcion_pt.exe").ToString("dd/MMM/yyyy hh:MM tt").ToUpper().Replace(".", "");
        }

        private void Recepcion_PT_Load(object sender, EventArgs e)
        {
            if (Utilerias.Class1.Usu_login.Trim() == "N")
            {
                pictureBox2.Visible = true;
                BtnFecCad.Visible = true;
            }
            if (System.Environment.MachineName.Trim().ToUpper() == "TUBO")
                BtnFecCad.Visible = true;

            lbfecha.Text = DateTime.Now.ToShortDateString();
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "select semana from tb_cat_semanas WHERE fecha1 <= '" + lbfecha.Text + "' AND fecha2 >= '" + lbfecha.Text + "'";
            reader1 = cmnd1.ExecuteReader();
            while (reader1.Read())
            {
                DateTime dateValue = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
                lbcletiqueta.Text = reader1.GetValue(0).ToString() + "-" + dateValue.ToString("ddd").ToUpper();
                cletiqueta = lbcletiqueta.Text.Substring(0, 5);
                if (cletiqueta.Trim().Contains("Á"))
                    cletiqueta = cletiqueta.Replace('Á', 'A');

                if (cletiqueta.Trim().Contains("É"))
                    cletiqueta = cletiqueta.Replace('É', 'E');

                lbcletiqueta.Text = cletiqueta;
            }
            reader1.Dispose();

            query = "select * from tb_cat_linea order by lin_nombre";
            da = new SqlDataAdapter(query, thisConnection);
            da.Fill(ds, "lineas");
            lineas = ds.Tables["lineas"];

            foreach (DataRow row in lineas.Rows)
                CBlin_nombre.Items.Add(Convert.ToString(row["lin_nombre"].ToString().Trim()));

            query = "select * from tb_cat_envases order by env_nombre";
            da = new SqlDataAdapter(query, thisConnection);
            da.Fill(ds, "envases");
            envases = ds.Tables["envases"];

            foreach (DataRow row in envases.Rows)
            {
                enva.Items.Add(Convert.ToString(row["env_nombre"].ToString().Trim()));
                env.Items.Add(Convert.ToString(row["env_nombre"].ToString().Trim()));
            }

            query = "select * from tb_cat_proveedor order by prov_nombre";
            da = new SqlDataAdapter(query, thisConnection);
            da.Fill(ds, "proveedores");
            proveedores = ds.Tables["proveedores"];

            query = "select * from tb_cat_ranchos order by rch_nombre";
            da = new SqlDataAdapter(query, thisConnection);
            da.Fill(ds, "ranchos");
            ranchos = ds.Tables["ranchos"];

            query = "select * from tb_cat_tablas order by tbl_nombre";
            da = new SqlDataAdapter(query, thisConnection);
            da.Fill(ds, "tablas");
            tablas = ds.Tables["tablas"];

            query = "select * from tb_cat_subtablas order by tbl_codigo";
            da = new SqlDataAdapter(query, thisConnection);
            da.Fill(ds, "subtablas");
            subtablas = ds.Tables["subtablas"];

            query = "select * from tb_cat_danos where dno_tipo = 'PT' and dno_estatus = 'ACTIVO' order by dno_clave";
            da = new SqlDataAdapter(query, thisConnection);
            da.Fill(ds, "danos");
            danos = ds.Tables["danos"];

            query = "select * from tb_cat_procesos where proc_tipo ='PT' and proc_estatus = 'ACTIVO' order by proc_clave";
            da = new SqlDataAdapter(query, thisConnection);
            da.Fill(ds, "procesos");
            procesos = ds.Tables["procesos"];

            query = "select * from tb_cat_producto where prod_tipo <> 'MP' and prod_tipo <> 'ING' and estatus = 'A' order by prod_nombre";
            da = new SqlDataAdapter(query, thisConnection);
            da.Fill(ds, "productos");
            productos = ds.Tables["productos"];

            query = "select * from tb_cat_variedad  order by vari_nombre";
            da = new SqlDataAdapter(query, thisConnection);
            da.Fill(ds, "variedades");
            variedades = ds.Tables["variedades"];
            thisConnection.Close();
            cbevaluacion.SelectedIndex = 0;

            /*cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "select lin_nombre from tb_cat_linea";
            reader1 = cmnd1.ExecuteReader();
            while(reader1.Read())
            {
                CBlin_nombre.Items.Add(reader1.GetValue(0).ToString().Trim());
            }
            reader1.Dispose();

            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "select env_nombre from tb_cat_envases order by env_nombre";
            reader1 = cmnd1.ExecuteReader();
            while(reader1.Read())
            {
                enva.Items.Add(reader1.GetValue(0).ToString().Trim());
                env.Items.Add(reader1.GetValue(0).ToString().Trim());
            }
            reader1.Dispose();
            thisConnection.Close();*/

            //txthora.Text = DateTime.Now.ToString("HH:mm");
        }

        #region METODO PARA VALIDAR LA ACTUALIZACION DEL EJECUTABLE
        public string ValidaActualizacion()
        {
            string Actu = "N";
            if (File.Exists(@"\\gabira1\sisgabweb\Valida.txt"))
            {
                DateTime FechaLocal = File.GetLastWriteTime(@"c:\sisgabweb\Recepcion_PT.exe");
                DateTime FechaServer = File.GetLastWriteTime(@"\\gabira1\sisgabweb\Recepcion_PT.exe");
                if (FechaServer > FechaLocal)
                {
                    MessageBox.Show("Hay una VERSION MAS Reciente se va a Cerrar el Sistema para que se Actualice, hay que volver abrir el programa!!", "Actualizacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return Actu = "S";
                }
            }
            return Actu;
        }
        #endregion

        private void btnAlta_Click_1(object sender, EventArgs e)
        {
            try
            {
                #region VALIDA Y FORZA ACTUALIZACION
                if (ValidaActualizacion() == "S")
                {
                    Actualizar = "S";
                    this.Close();
                }
                #endregion
                cbrecpen.Enabled = true;
                cbreccamp.Enabled = true;
                opcion = 1;
                thisConnection.Open();
                //trae el consecutivo de la tb_mstr_recepcion_mp
                cmnd1 = thisConnection.CreateCommand();
                //cmnd1.CommandText = "SELECT max(rpt_recibo) FROM tb_mstr_recepcion_pt";
                cmnd1.CommandText = "SELECT max(rpt_recibo) FROM tb_consecutivo_pt";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    recibo = Convert.ToInt32(reader1.GetValue(0).ToString()) + 1;
                }
                reader1.Dispose();
                txtrecibo.Text = recibo.ToString();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT A.id_ticket, A.prod_clave, B.prod_nombre, A.variedad, C.hora_peso, A.tipo_prod, A.cantidad, A.rch_clave, A.num_prod " +
                                    "FROM tb_det_recepcion_bascula A, tb_cat_producto B, tb_mstr_recepcion_bascula C " +
                                    "WHERE A.tipo_prod <> 'MP' and B.prod_clave = A.prod_clave and C.id_ticket = A.id_ticket " +
                                    "and A.prod_clave <> '' AND A.estatus ='P' order by A.id_ticket desc";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    cbticket.Items.Add(reader1.GetValue(0).ToString().Trim() + " -> " + reader1.GetValue(1).ToString().Trim() + " -> " + reader1.GetValue(4).ToString().Trim() + " -> " + reader1.GetValue(6).ToString().Trim() + " -> " + reader1.GetValue(7).ToString().Trim() + " -> " + reader1.GetValue(8).ToString().Trim());
                    ticket_bascula.Items.Add(reader1.GetValue(0).ToString().Trim());
                    prod_clave.Items.Add(reader1.GetValue(1).ToString().Trim());
                    variedad.Items.Add(reader1.GetValue(3).ToString().Trim());
                    nume_prod.Items.Add(reader1.GetValue(8).ToString().Trim());
                }
                reader1.Dispose();

                cbrecpen.Items.Clear();
                DataTable etiq = new DataTable("etiq");
                DataTable pen = new DataTable("pen");
                DataSet sd = new DataSet();
                //etiquetas
                /*string query = "select rpt_recibo from tb_mstr_recepcion_pt where rpt_fecha = '" + DateTime.Now.ToShortDateString() + "' and rpt_estatus <> 'F'  and "+
                               "(rpt_situacion <> 'A' or rpt_situacion  IS NULL) and rpt_flete <> 'N.CRED' and rpt_flete <> 'DEV' order by rpt_recibo";*/
                //string query = "select rpt_recibo from tb_mstr_recepcion_pt where ind = '' and rpt_estatus <> 'F'  and " +
                //               "(rpt_situacion <> 'A' or rpt_situacion  IS NULL) and rpt_flete <> 'N.CRED' and rpt_flete <> 'DEV' AND rpt_pesador <> 'AGUILARES' and "+
                //               "rpt_evaluador <> 'AGUILARES' and prov_clave <> 'AJUSTE' and prov_clave <> 'TAYLOR' AND responsable <> '' order by rpt_recibo";
                string query = "select a.rpt_recibo, b.prod_clave, c.prod_nombre from tb_mstr_recepcion_pt a, tb_det_recepcion_pt b, tb_cat_producto c " +
                               "where a.ind = '' and a.rpt_estatus <> 'F'  and (a.rpt_situacion <> 'A' or a.rpt_situacion  IS NULL) and a.rpt_flete <> 'N.CRED' " +
                               "and a.rpt_flete <> 'DEV' AND a.rpt_pesador <> 'AGUILARES' and a.rpt_evaluador <> 'AGUILARES' and a.prov_clave <> 'AJUSTE' " +
                               "and a.prov_clave <> 'TAYLOR' AND a.responsable <> '' and b.rpt_recibo = a.rpt_recibo and c.prod_clave = b.prod_clave " +
                               "order by a.rpt_recibo";
                da = new SqlDataAdapter(query, thisConnection);
                da.Fill(sd, "etiq");
                etiq = sd.Tables["etiq"];

                //elimina lo de taylor
                foreach (DataRow row in etiq.Rows)
                {
                    if (Convert.ToString(row["prod_nombre"].ToString().Trim()).Contains("TAYLOR"))
                        row.Delete();
                }
                etiq.AcceptChanges();

                query = "select recibo from tb_det_trazabilidad where pti_fecha = '" + DateTime.Now.ToShortDateString() + "' and tipo = 'PTC'";
                da = new SqlDataAdapter(query, thisConnection);
                da.Fill(sd, "pen");
                pen = sd.Tables["pen"];

                //query = "select a.rpt_recibo, a.id_ticket from tb_mstr_recepcion_bascula a, tb_mstr_recepcion_pt b where b.rpt_recibo = a.rpt_recibo order by a.rpt_recibo";
                query = "select a.rpt_recibo, a.id_ticket, c.prod_nombre from tb_det_recepcion_bascula a, tb_det_recepcion_pt b, tb_cat_producto c where b.rpt_recibo = a.rpt_recibo and " +
                        "a.estatus = 'P' and c.prod_clave = a.prod_clave";
                da = new SqlDataAdapter(query, thisConnection);
                da.Fill(sd, "campo");
                campo = sd.Tables["campo"];
                thisConnection.Close();

                bool encontro = false;
                foreach (DataRow row in etiq.Rows)
                {
                    encontro = false;

                    foreach (DataRow row1 in pen.Select("recibo = '" + Convert.ToString(row["rpt_recibo"].ToString()) + "'"))
                    {
                        encontro = true;
                        break;
                    }
                    if (encontro == false)
                        cbrecpen.Items.Add(Convert.ToString(row["rpt_recibo"].ToString()));
                }

                cbreccamp.Items.Clear();
                foreach (DataRow row in campo.Rows)
                    cbreccamp.Items.Add(row["rpt_recibo"].ToString().Trim() + " ~ " + row["id_ticket"].ToString().Trim() + " ~ " + row["prod_nombre"].ToString().Trim());

                //thisConnection.Close();
                cbticket.Enabled = true;
                cbticket.Focus();
                btnAlta.Enabled = false;
                btnConsulta.Enabled = false;
                btnCancel.Enabled = true;
                btnGuardar.Enabled = true;
                cbtipo.Enabled = true;
                cbtipo.SelectedIndex = 0;
                cbtipo_SelectionChangeCommitted(sender, e);
                btnntacre.Enabled = true;
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                //error = ex.ToString();
                //ThreadStart delegado = new ThreadStart(CorrerProceso);
                //Thread hilo = new Thread(delegado);
                //hilo.Start();
                //Thread.Sleep(1000);
                MessageBox.Show(ex.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "Sistem@s2026$", ex.ToString().Trim());
                return;
            }
            catch (Exception ex1)
            {
                thisConnection.Close();
                MessageBox.Show(ex1.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "Sistem@s2026$", ex1.ToString().Trim());
                return;
            }
        }

        //string error = "";
        private void CorrerProceso()
        {
            //Rectangle bounds = Screen.GetBounds(Point.Empty);
            //using (Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height))
            //{
            //    using (Graphics g = Graphics.FromImage(bitmap))
            //    {
            //        g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
            //    }
            //    bitmap.Save(@"C:\SisGabWeb\error.jpg");
            //}
            //Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", error.ToString().Trim());
            //return;
        }

        private void cbticket_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                txtrecibo.Text = recibo.ToString();
                ticket_bascula.SelectedIndex = cbticket.SelectedIndex;
                prod_clave.SelectedIndex = cbticket.SelectedIndex;
                variedad.SelectedIndex = cbticket.SelectedIndex;
                nume_prod.SelectedIndex = cbticket.SelectedIndex;
                txtprov_clave.ReadOnly = false;
                txtrch_clave.ReadOnly = false;
                txttbl_clave.ReadOnly = false;
                txtrecibo.ReadOnly = true;
                fecha_cad = ""; Convert.ToDateTime(lbfecha.Text).AddDays(15).ToShortDateString();
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "select A.prod_clave, B.prod_nombre, A.cantidad, A.envase, A.tara, A.peso_bruto, A.peso_neto, A.id_tabla, A.tabla, A.variedad, A.lin_clave, " +
                                    "A.cantidad, A.num_tarimas, A.prov_clave, A.rch_clave, C.folio, A.num_prod, A.flete, C.folio, C.hora_peso, B.prod_paisorigen " +
                                    "from tb_det_recepcion_bascula A, tb_cat_producto B, tb_mstr_recepcion_bascula C where A.prod_clave = '" + prod_clave.SelectedItem.ToString() + "' and A.id_ticket = " + ticket_bascula.SelectedItem.ToString() + " AND A.num_prod = " + Convert.ToInt32(nume_prod.SelectedItem.ToString().Trim()) + " and " +
                                    "B.prod_clave = A.prod_clave and A.estatus = 'P' and C.id_ticket = A.id_ticket";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    producto = reader1.GetValue(0).ToString().Trim();
                    nombre = reader1.GetValue(1).ToString();
                    envas = reader1.GetValue(3).ToString().Trim();
                    tar = Convert.ToDecimal(reader1.GetValue(4).ToString());
                    b_p = Convert.ToDecimal(reader1.GetValue(5).ToString());
                    p_t = Convert.ToDecimal(reader1.GetValue(6).ToString());
                    tabla = reader1.GetValue(7).ToString().Trim();
                    nom_tabla = reader1.GetValue(8).ToString().Trim();
                    varieda = reader1.GetValue(9).ToString().Trim();
                    linea = reader1.GetValue(10).ToString().Trim();
                    cantidad = Convert.ToInt32(reader1.GetValue(11).ToString());
                    tarm = Convert.ToInt32(reader1.GetValue(12).ToString());
                    provee = reader1.GetValue(13).ToString().Trim();
                    rancho = reader1.GetValue(14).ToString().Trim();
                    if ((reader1.GetValue(17).ToString().Trim() != "")) // No. FLETE
                    {
                        if (reader1.GetValue(17).ToString().Trim() != "0")
                            flt = reader1.GetValue(17).ToString().Trim();
                    }
                    else
                        flt = "";

                    num_prod = Convert.ToInt32(reader1.GetValue(16).ToString());
                    txthora.Text = reader1.GetValue(19).ToString();
                    pais = reader1.GetValue(20).ToString().Trim();
                    txtflete.Text = flt;
                }
                reader1.Dispose();

                if (DGV1.Rows.Count == 0)
                {
                    txttbl_clave.Text = tabla;
                    txttabla.Text = nom_tabla;
                    DGV2.Rows.Clear();
                    DGV3.Rows.Clear();

                    //trae el proveedor
                    foreach (DataRow row in proveedores.Select("prov_clave = '" + provee + "'"))
                    {
                        txtprov_clave.Text = Convert.ToString(row["prov_clave"].ToString().Trim());
                        txtprov.Text = Convert.ToString(row["prov_nombre"].ToString().Trim());
                        email = Convert.ToString(row["prov_email"].ToString().Trim());
                    }

                    /*cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select prov_nombre from tb_cat_proveedor where prov_clave = '" + provee + "'";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        txtprov_clave.Text = provee;
                        txtprov.Text = reader1.GetValue(0).ToString().Trim();
                    }
                    reader1.Dispose();*/

                    //trae el rancho
                    foreach (DataRow row in ranchos.Select("prov_clave= '" + txtprov_clave.Text + "' and rch_clave = '" + rancho + "'"))
                    {
                        txtrch_clave.Text = Convert.ToString(row["rch_clave"].ToString().Trim());
                        txtrancho.Text = Convert.ToString(row["rch_nombre"].ToString().Trim());
                    }

                    /*cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select rch_nombre from tb_cat_ranchos where prov_clave = '" + txtprov_clave.Text + "' and rch_clave = '" + rancho + "'";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        txtrch_clave.Text = rancho;
                        txtrancho.Text = reader1.GetValue(0).ToString().Trim();
                    }
                    reader1.Dispose();*/

                    //trae el codigo
                    foreach (DataRow row in subtablas.Select("prov_clave = '" + txtprov_clave.Text + "' and rch_clave = '" + txtrch_clave.Text + "' and tbl_clave = '" + tabla + "'"))
                        CBcodigo.Items.Add(Convert.ToString(row["tbl_codigo"].ToString().Trim()));

                    /*cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select tbl_codigo from tb_cat_subtablas where prov_clave = '" + txtprov_clave.Text + "' and rch_clave = '" + txtrch_clave.Text + "' and tbl_clave = '" + tabla + "'";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        CBcodigo.Items.Add(reader1.GetValue(0).ToString().Trim());
                    }
                    reader1.Dispose();*/

                    //selecciona la linea
                    foreach (DataRow row in lineas.Select("lin_clave = '" + linea + "'"))
                    {
                        txtlin_clave.Text = Convert.ToString(row["lin_clave"].ToString().Trim());
                        CBlin_nombre.SelectedItem = Convert.ToString(row["lin_nombre"].ToString().Trim());
                    }

                    /*cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select lin_nombre from tb_cat_linea where lin_clave = '" + linea + "' order by lin_nombre";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        txtlin_clave.Text = linea;
                        CBlin_nombre.SelectedItem = reader1.GetValue(0).ToString().Trim();
                    }
                    reader1.Dispose();*/

                    //carga el combo de productos de acuerdo a la linea
                    foreach (DataRow row in productos.Select("lin_clave ='" + txtlin_clave.Text + "'"))
                        nomb.Items.Add(Convert.ToString(row["prod_nombre"].ToString().Trim()));

                    /*cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select prod_nombre from tb_cat_producto where lin_clave = '" + txtlin_clave.Text + "' and prod_tipo <> 'MP' and prod_tipo <> 'ING' order by prod_nombre";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        nomb.Items.Add(reader1.GetValue(0).ToString());
                    }
                    reader1.Dispose();*/

                    //trae los datos de tb_mstr_recepcion_bascula
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select folio, recibio_bascula from tb_mstr_recepcion_bascula where id_ticket = " + ticket_bascula.SelectedItem.ToString();
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        //txtflete.Text = reader1.GetValue(0).ToString().Trim();
                        cmnd11 = thisConnection.CreateCommand();
                        cmnd11.CommandText = "select bas_nombre from tb_cat_basculeros where bas_clave = '" + reader1.GetValue(1).ToString().Trim() + "'";
                        reader11 = cmnd11.ExecuteReader();
                        while (reader11.Read())
                        {
                            txtpesador.Text = reader11.GetValue(0).ToString().Trim();
                        }
                        reader11.Dispose();
                    }
                    reader1.Dispose();

                    //trae la variedad    
                    foreach (DataRow row in variedades.Select("lin_clave= '" + txtlin_clave.Text + "'"))
                        cbvariedad.Items.Add(Convert.ToString(row["vari_nombre"].ToString().Trim()));

                    //if (varieda.Trim() == "")
                    //{
                    /*cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select vari_nombre from tb_cat_variedad where lin_clave = '" + txtlin_clave.Text + "' order by vari_nombre";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        cbvariedad.Items.Add(reader1.GetValue(0).ToString().Trim());
                        //vari.Items.Add(reader1.GetValue(0).ToString().Trim());
                    }
                    reader1.Dispose();*/
                    //}
                    //else
                    //{
                    //    cmnd1 = thisConnection.CreateCommand();
                    //    cmnd1.CommandText = "select vari_clave from tb_cat_variedad where vari_nombre = '" + varieda + "' and lin_clave = '"+txtlin_clave.Text+"'";
                    //    txtvariedad.Text = Convert.ToString(cmnd1.ExecuteScalar()).Trim();
                    //    cbvariedad.Items.Add(varieda);
                    //    cbvariedad.SelectedItem = varieda;
                    //}
                    cbvariedad.Enabled = true;
                    txtvariedad.ReadOnly = false;
                    //trae el peso del envase
                    foreach (DataRow row in envases.Select("env_nombre = '" + envas + "'"))
                    {
                        peso_env = Convert.ToDecimal(row["env_peso"].ToString());
                        env_clave = Convert.ToString(row["env_clave"].ToString());
                    }
                    /*cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select env_peso, env_clave from tb_cat_envases where env_nombre = '" + envas + "'";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        peso_env = Convert.ToDecimal(reader1.GetValue(0).ToString());
                        env_clave = reader1.GetValue(1).ToString().Trim();
                    }
                    reader1.Dispose();*/

                    //trae los defectos de acuerdo a la linea
                    foreach (DataRow row in danos.Select("lin_clave = '" + linea + "'"))
                        DGV2.Rows.Add(Convert.ToString(row["dno_nombre"].ToString().Trim()), 0, Convert.ToString(row["dno_clave"].ToString().Trim()), 0, row["dno_Calidad"].ToString());

                    /*cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select dno_nombre, dno_clave from tb_cat_danos where lin_clave = '" + linea + "' and dno_tipo = 'PT' order by dno_clave";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        DGV2.Rows.Add(reader1.GetValue(0).ToString().Trim(), 0, reader1.GetValue(1).ToString().Trim());
                    }
                    reader1.Dispose();*/

                    //trae los procesos de acuerdo a la linea
                    foreach (DataRow row in procesos.Select("lin_clave = '" + linea + "'"))
                        DGV3.Rows.Add(Convert.ToString(row["proc_nombre"].ToString().Trim()), 0, Convert.ToString(row["proc_clave"].ToString().Trim()));

                    /*cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select proc_nombre, proc_clave from tb_cat_procesos where lin_clave = '" + linea + "' and proc_tipo ='PT' order by proc_clave";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        DGV3.Rows.Add(reader1.GetValue(0).ToString().Trim(), 0, reader1.GetValue(1).ToString().Trim());
                    }
                    reader1.Dispose();*/
                    thisConnection.Close();

                    num_viaje = Convert.ToInt32(txtnumviaje.Text);
                    proveedor = txtprov.Text;
                    rancho_nom = txttabla.Text;
                    hora = txthora.Text;

                    CBcodigo.Enabled = true;
                    CBlin_nombre.Enabled = true;
                    txtcodigo_clave.ReadOnly = false;
                    txtlin_clave.ReadOnly = false;
                    txtnumviaje.ReadOnly = false;
                    txtnumped.ReadOnly = false;
                    CBIngProd_fis.Enabled = true;
                    txtevaluador.ReadOnly = false;
                    txtobs.ReadOnly = false;
                    txthora.ReadOnly = false;
                    DGV2.Columns[1].ReadOnly = false;
                    DGV3.Columns[1].ReadOnly = false;
                    env.SelectedItem = envas;
                    cbevaluacion.Enabled = true;
                    TxtPieza.ReadOnly = false;
                    //p_t = b_p - tar - (cantidad - peso_env) - (tarm * 20);
                    p_t = b_p - tar - (cantidad * peso_env) - (tarm * 20);

                    if (cantidad > 0)
                        p_u = p_t / cantidad;
                    else
                        p_u = 0;

                    DGV1.Rows.Add(producto, nombre, ticket_bascula.SelectedItem.ToString(), b_p.ToString(), tar.ToString(), envas, cantidad.ToString(), tarm.ToString(), p_u, p_t, fecha_cad, "", env_clave, "", peso_env, num_prod, pais);
                    //DGV1.Rows.Add(producto, nombre, ticket_bascula.SelectedItem.ToString(), b_p.ToString(), 0, envas, cantidad.ToString(), tarm.ToString(), p_u, p_t, fecha_cad, "", env_clave, "", peso_env, num_prod, pais);
                    //DGV1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    DGV2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    DGV3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    CBlin_nombre.Enabled = false;
                    txtlin_clave.Enabled = false;
                    CBMoneda.SelectedIndex = 0;
                    CBMoneda.Enabled = true;
                    txtflete.ReadOnly = false;
                    TxtPieza.Text = "0";
                }
                //if (DGV1.Rows.Count > 0)
                else
                {
                    if ((provee.Trim() != txtprov_clave.Text.Trim()) || (rancho.Trim() != txtrch_clave.Text.Trim()) || (tabla.Trim() != txttbl_clave.Text.Trim()) || (linea.Trim() != txtlin_clave.Text.Trim()))
                    {
                        thisConnection.Close();
                        MessageBox.Show("No se puede agregar el producto debido a que algunos de los datos (proveedor, rancho, tabla y/o linea) no concuerdan", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    else
                    {
                        thisConnection.Close();
                        for (int i = 0; i < DGV1.Rows.Count; i++)
                        {
                            if ((Convert.ToString(DGV1.Rows[i].Cells["produc"].Value).Trim() == producto.Trim()) && (Convert.ToString(DGV1.Rows[i].Cells["bascula_ticket"].Value).Trim() == ticket_bascula.SelectedItem.ToString().Trim()))
                            {
                                MessageBox.Show("El producto " + DGV1.Rows[i].Cells["Nomb"].Value.ToString().Trim() + " ya fue agregado", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                thisConnection.Close();
                                return;
                            }
                        }
                        //p_t = b_p - tar - (cantidad - peso_env) - (tarm * 20);
                        p_t = b_p - tar - (cantidad * peso_env) - (tarm * 20);

                        if (cantidad > 0)
                            p_u = p_t / cantidad;
                        else
                            p_u = 0;

                        //trae el peso del envase
                        foreach (DataRow row in envases.Select("env_nombre = '" + envas + "'"))
                        {
                            peso_env = Convert.ToDecimal(row["env_peso"].ToString());
                            env_clave = Convert.ToString(row["env_clave"].ToString());
                        }
                        //cmnd1 = thisConnection.CreateCommand();
                        //cmnd1.CommandText = "select env_peso, env_clave from tb_cat_envases where env_nombre = '" + envas + "'";
                        //reader1 = cmnd1.ExecuteReader();
                        //while (reader1.Read())
                        //{
                        //    peso_env = Convert.ToDecimal(reader1.GetValue(0).ToString());
                        //    env_clave = reader1.GetValue(1).ToString().Trim();
                        //}
                        //reader1.Dispose();
                        //thisConnection.Close();
                        DGV1.Rows.Add(producto, nombre, ticket_bascula.SelectedItem.ToString(), b_p.ToString(), tar.ToString(), envas, cantidad.ToString(), tarm.ToString(), p_u, p_t, fecha_cad, "", env_clave, "", peso_env, num_prod, pais);
                        //DGV1.Rows.Add(producto, nombre, ticket_bascula.SelectedItem.ToString(), b_p.ToString(), 0, envas, cantidad.ToString(), tarm.ToString(), p_u, p_t, fecha_cad, "", env_clave, "", peso_env, num_prod, pais);
                        //DGV1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    }
                }

                cbcontrato.Items.Clear();
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "select a.cont_folio from tb_mstr_contratos a, tb_det_contratos b, tb_cat_producto c where a.prov_clave = '" + txtprov_clave.Text + "' " +
                                    "and a.cont_status = 'A' and b.cont_folio = a.cont_folio and b.rch_clave = '" + txtrch_clave.Text + "' " +
                                    "and b.tbl_clave = '" + txttbl_clave.Text + "' and c.lin_clave = '" + txtlin_clave.Text + "' and a.cult_folio = c.prod_clave";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                    cbcontrato.Items.Add(reader1.GetValue(0).ToString().Trim());

                reader1.Dispose();
                thisConnection.Close();
                cbcontrato.Enabled = true;
                DGV1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                for (int i = 0; i < DGV1.Rows.Count; i++)
                {
                    this.DGV1.Rows[i].HeaderCell.Value = (i + 1).ToString();
                }
                ini_captura = DateTime.Now.ToString();
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString().Trim());
                MessageBox.Show(ex.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            catch (Exception ex1)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString().Trim());
                MessageBox.Show(ex1.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void DGV1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception.Message == "El valor de DataGridViewComboBoxCell no es válido." || e.Exception.Message == "DataGridViewComboBoxCell value is not valid.")
            {
                object value = DGV1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (!((DataGridViewComboBoxColumn)DGV1.Columns[e.ColumnIndex]).Items.Contains(value))
                {
                    ((DataGridViewComboBoxColumn)DGV1.Columns[e.ColumnIndex]).Items.Add(value);
                    e.ThrowException = false;
                }
            }
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            #region guardar nuevo registro
            if (opcion == 1)
            {
                for (int i = 0; i < DGV1.Rows.Count; i++)
                {
                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_uni"].Value) == 0)
                    {
                        MessageBox.Show("Favor de agregar el peso unitario en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["pes_uni"].Selected = true;
                        return;
                    }

                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["tarim"].Value) == 0)
                    {
                        MessageBox.Show("Favor de agregar número de tarimas en la fila" + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["tarim"].Selected = true;
                        return;
                    }

                    if (Convert.ToString(DGV1.Rows[i].Cells["enva_clave"].Value) == "")
                    {
                        MessageBox.Show("Favor de seleccionar el envase en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["enva"].Selected = true;
                        return;
                    }


                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_uni"].Value) < 0)
                    {
                        MessageBox.Show("El Peso Unitario debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["pes_uni"].Selected = true;
                        return;
                    }

                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_tot"].Value) < 0)
                    {
                        MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["pes_tot"].Selected = true;
                        return;
                    }

                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) < 0)
                    {
                        MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                        return;
                    }

                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value) < 0)
                    {
                        MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["ta"].Selected = true;
                        return;
                    }

                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) <= Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value))
                    {
                        MessageBox.Show("El Peso Bruto debe ser mayor a la Tara en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                        return;
                    }


                    if (DGV1.Rows[i].Cells[10].Value.ToString().Trim().Length > 0)
                    {
                        try
                        {
                            DateTime fechatar = DateTime.ParseExact(Convert.ToString(DGV1.Rows[i].Cells[10].Value.ToString().Trim()), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                        }
                        catch
                        {
                            MessageBox.Show("Favor de Ingresar la Fecha de Caducidad con un formato Valido, dd/MM/yyyy", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Rows[i].Cells["FechaCad"].Selected = true;
                            return;
                        }
                    }
                    if (Convert.ToInt32(DGV1.Rows[i].Cells["tarim"].Value.ToString().Trim()) > 999)
                    {
                        MessageBox.Show("El numero de Tarimas no puede superar las 999, favor de recalcular", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Rows[i].Cells["tarim"].Selected = true;
                        return;
                    }
                }

                if ((TxtPieza.Text.Trim() == "0" || TxtPieza.Text.Trim().Length == 0) && txtflete.Text != "N.CRED")
                {
                    MessageBox.Show("Favor de Capturar el Numero de Piezas de la Evaluacion", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtPieza.Focus();
                    return;
                }
                if (Convert.ToDecimal(txtnumviaje.Text) == 0)
                {
                    MessageBox.Show("Favor de escribir el número de viaje", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtnumviaje.Focus();
                    return;
                }

                if (txtvariedad.Text.Trim().Length == 0 && txtflete.Text.Trim() != "N.CRED")
                {
                    MessageBox.Show("Favor de Capturar la Variedad en caso de que no aplique dar de alta un concepto que diga No Aplica", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtvariedad.Focus();
                    return;
                }
                if (txtflete.Text == "N.CRED")
                    TxtPieza.Text = "0";
                /*if (txtflete.Text.Trim() == "N.CRED")
                {
                    if (DGV4.Rows.Count == 0)
                    {
                        MessageBox.Show("La devolución debe de tener al menos un folio", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    int count = 0, cont= 0;
                    for (int i = 0; i < DGV4.Rows.Count; i++)
                    {
                        //count = Convert.ToInt32(DGV4.Rows[i].Cells["boxes"].Value);
                        if (i == 0)
                            count = Convert.ToInt32(DGV4.Rows[i].Cells["boxes"].Value);
                        else
                        { 
                            //si el producto cambia
                            if (Convert.ToString(DGV4.Rows[i].Cells["cve_prod"].Value).Trim() != Convert.ToString(DGV4.Rows[i - 1].Cells["cve_prod"].Value).Trim())
                            {
                                cont = 0;
                                foreach (DataRow row in not_cre.Select("producto = '" + Convert.ToString(DGV4.Rows[i - 1].Cells["cve_prod"].Value) + "'"))                            
                                    cont += Convert.ToInt32(row["cantidad"].ToString());

                                if (count != cont)
                                {
                                    MessageBox.Show("Las cantidades de caja no son iguales del producto " + Convert.ToString(DGV4.Rows[i - 1].Cells["cve_prod"].Value), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                                count = Convert.ToInt32(DGV4.Rows[i].Cells["boxes"].Value);
                            }
                            else
                                count += Convert.ToInt32(DGV4.Rows[i].Cells["boxes"].ToString().Trim());
                        }
                    }
                }*/

                //if (notacre != true)
                //{
                //    if (DGV1.Rows.Count == 0 || cbticket.SelectedIndex == -1)
                //    {
                //        MessageBox.Show("Favor de seleccionar un ticket de báscula y/o no hay producto(s) agregado(s)", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //        return;
                //    }
                //}
                if (cbcontrato.SelectedIndex == -1)
                    contrato = "";
                else
                    contrato = cbcontrato.SelectedItem.ToString().Trim();

                label7.Visible = true;

                #region SQL
                string querys = "";
                try
                {
                    for (int i = 0; i < DGV1.Rows.Count; i++)
                    {
                        #region VALIDACION DE PESO BRUTO Y TARA SON MAYOR A 0
                        if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) < 0)
                        {
                            MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Focus();
                            DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                            return;
                        }

                        if (Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value) < 0)
                        {
                            MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Focus();
                            DGV1.Rows[i].Cells["ta"].Selected = true;
                            return;
                        }

                        if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_uni"].Value) < 0)
                        {
                            MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Focus();
                            DGV1.Rows[i].Cells["ta"].Selected = true;
                            return;
                        }

                        if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) <= Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value))
                        {
                            MessageBox.Show("El Peso Bruto debe ser mayor a la Tara en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Focus();
                            DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                            return;
                        }
                        #endregion
                    }

                    thisConnection.Open();
                    #region guarda en tb_mstr_recepcion_pt
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "insert into tb_consecutivo_pt (fecha) values ('" + lbfecha.Text + "') select scope_identity()";
                    string error = Convert.ToString(cmnd1.ExecuteScalar());

                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "BEGIN TRY " +
                                        "insert into tb_mstr_recepcion_pt (rpt_recibo, id_regpt, rpt_fecha, prov_clave, rch_clave, tbl_clave, tbl_vieja, rpt_codigo, lin_clave, " +
                                        "rpt_tipo, rpt_estatus, rpt_flete, rpt_observaciones, vari_clave, rpt_pesador, rpt_evaluador, rpt_recref, rpt_tiporef, " +
                                        "rpt_cve_fecha, rpt_viaje, rpt_hora, rpt_enviado, rpt_inventario, rpt_situacion, id_flete, id_ticket, id_registro, orde, " +
                                        "transporte, ind, nc, inicio_captura, responsable, um_clave, rpt_evaluacion, folio_no_conformidad, numero_pedimento, contrato_proveedor, rpt_piezasEvaluacion) values " +
                                        "('" + error + "', 0, '" + lbfecha.Text + "', '" + txtprov_clave.Text + "', " +
                                        "'" + txtrch_clave.Text + "', '" + txttbl_clave.Text + "', '', '" + txtcodigo_clave.Text + "', '" + txtlin_clave.Text.Trim() + "', " +
                                        "'" + tipo + "', '', '" + txtflete.Text + "', '" + txtobs.Text + "', '" + txtvariedad.Text + "', '" + txtpesador.Text + "', " +
                                        "'" + txtevaluador.Text + "', '', '', '" + lbcletiqueta.Text.Substring(0, 5) + "', " + Convert.ToInt32(txtnumviaje.Text) + ", " +
                                        "'" + txthora.Text + "', '', '" + situacion + "', 'C', 0, 0, 0, '', '', '', '', '" + ini_captura.Trim() + "', " +
                                        "'" + Environment.MachineName + "', '" + CBMoneda.SelectedItem.ToString().Trim() + "', '" + cbevaluacion.SelectedItem.ToString().Trim() + "', " +
                                        "'" + txtfolnocon.Text + "', '" + txtnumped.Text + "', '" + contrato + "','" + TxtPieza.Text + "') END TRY BEGIN CATCH SELECT ERROR_MESSAGE() AS WEY END CATCH";
                    reader1 = cmnd1.ExecuteReader();
                    //string error = txtrecibo.Text;
                    //string error = Convert.ToString(cmnd1.ExecuteScalar());

                    if (Char.IsLetter(error, 0))
                    {
                        thisConnection.Close();
                        MessageBox.Show(error.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    txtrecibo.Text = error;
                    recibo = Convert.ToInt32(error);
                    #endregion

                    #region guarda en tb_det_recepcion_pt
                    for (int i = 0; i < DGV1.Rows.Count; i++)
                    {
                        #region VALIDACION DE PESO BRUTO Y TARA SON MAYOR A 0
                        if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) < 0)
                        {
                            MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Focus();
                            DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                            return;
                        }

                        if (Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value) < 0)
                        {
                            MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Focus();
                            DGV1.Rows[i].Cells["ta"].Selected = true;
                            return;
                        }

                        if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_uni"].Value) < 0)
                        {
                            MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Focus();
                            DGV1.Rows[i].Cells["ta"].Selected = true;
                            return;
                        }

                        if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) <= Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value))
                        {
                            MessageBox.Show("El Peso Bruto debe ser mayor a la Tara en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Focus();
                            DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                            return;
                        }
                        #endregion

                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "insert into tb_det_recepcion_pt (rpt_recibo, rptd_ticket, lin_clave, prod_clave, rptd_peso_bruto, rptd_tara, env_clave, rptd_tarimas, rptd_cantidad, " +
                                            "rptd_fechacad, producto, caj_tarimas, resto, peso_neto, kg_caja, kg_cajas_vacias, id_tabla, tabla, variedad, variedad_ofc, linea_ofc, observaciones, " +
                                            "peso_unitario, postura, rptd_estatus) values (" + Convert.ToInt32(txtrecibo.Text) + ", '" + Convert.ToString(DGV1.Rows[i].Cells[2].Value) + "', '" + txtlin_clave.Text + "', " +
                                            "'" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "', " + Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) + ", " + Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) + ", " +
                                            "'" + Convert.ToString(DGV1.Rows[i].Cells[12].Value) + "', " + Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) + ", " + Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value) + ", " +
                                            "'" + Convert.ToString(DGV1.Rows[i].Cells[10].Value) + "', '', 0, 0, '" + Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) + "', 0, 0, '', '', '" + txtvariedad.Text + "', '', '', '', 0, '', '')";
                        reader1 = cmnd1.ExecuteReader();
                        reader1.Dispose();
                    }
                    #endregion

                    #region guarda defectos
                    if (DGV2.Rows.Count > 0)
                    {
                        for (int i = 0; i < DGV2.Rows.Count; i++)
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "insert into tb_mstr_danos (rmp_recibo, lin_clave, dno_clave, dnm_cantidad, dnm_tipo, dnm_peso) " +
                                                "values ('" + txtrecibo.Text + "', '" + txtlin_clave.Text + "', '" + Convert.ToString(DGV2.Rows[i].Cells[2].Value) + "', " +
                                                "" + Convert.ToDecimal(DGV2.Rows[i].Cells[1].Value) + ", 'PT', 0.000)";
                            reader1 = cmnd1.ExecuteReader();
                            reader1.Dispose();
                        }
                    }
                    #endregion

                    #region guarda procesos
                    if (DGV3.Rows.Count > 0)
                    {
                        for (int i = 0; i < DGV3.Rows.Count; i++)
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "insert into tb_mstr_procesos (rmp_recibo, lin_clave, proc_clave, prom_cantidad, prom_tipo, prom_peso) values " +
                                                "('" + txtrecibo.Text + "', '" + txtlin_clave.Text + "', '" + Convert.ToString(DGV3.Rows[i].Cells[2].Value) + "', " +
                                                "" + Convert.ToDecimal(DGV3.Rows[i].Cells[1].Value) + ", 'PT', 0)";
                            reader1 = cmnd1.ExecuteReader();
                            reader1.Dispose();
                        }
                    }
                    #endregion

                    #region guarda en tb_hist_recepcion
                    for (int i = 0; i < DGV1.Rows.Count; i++)
                    {
                        decimal hrp_peso_neto = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                        decimal hrp_clase1 = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                        decimal hrp_peso_util = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                        decimal hrp_reman_kg = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                        int hrp_num_unidades = Convert.ToInt32(DGV1.Rows[i].Cells[6].Value);
                        int hrp_reman_unidades = Convert.ToInt32(DGV1.Rows[i].Cells[6].Value);

                        cmnd1 = thisConnection.CreateCommand();
                        if (cbevaluacion.SelectedIndex == 0 || cbevaluacion.SelectedIndex == 1)
                            cmnd1.CommandText = "INSERT INTO tb_hist_recepcion (hrp_recibo, hrp_fecha, lin_clave, hrp_tipo_recepcion, hrp_estatus, prod_clave, hrp_peso_neto, hrp_clase1, " +
                                                "hrp_peso_util, " +
                                                "hrp_num_unidades, hrp_reman_kg, hrp_reman_unidades, hrp_situacion, hrp_clase2, hrp_proceso, hrp_surtido, hrp_liquidado, hrp_numliq) values " +
                                                "(" + txtrecibo.Text + ", '" + lbfecha.Text + "', '" + txtlin_clave.Text + "', 'PTC', 'T', " +
                                                "'" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "'," + hrp_peso_neto + ", " + hrp_clase1 + ", " + hrp_peso_util + "," + hrp_num_unidades + ", " + hrp_reman_kg + ", " +
                                                "" + hrp_reman_unidades + ", '" + tipo + "', 0, 0, '" + hrp_peso_neto + "', '', '')";
                        if (cbevaluacion.SelectedIndex == 2)
                            cmnd1.CommandText = "INSERT INTO tb_hist_recepcion (hrp_recibo, hrp_fecha, lin_clave, hrp_tipo_recepcion, hrp_estatus, prod_clave, hrp_peso_neto, hrp_clase1, " +
                                                "hrp_peso_util, " +
                                                "hrp_num_unidades, hrp_reman_kg, hrp_reman_unidades, hrp_situacion, hrp_clase2, hrp_proceso, hrp_surtido, hrp_liquidado, hrp_numliq) " +
                                                "values (" + txtrecibo.Text + ", " +
                                                "'" + lbfecha.Text + "', '" + txtlin_clave.Text + "', 'PTC', 'C', " +
                                                "'" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "'," + hrp_peso_neto + ", " + hrp_clase1 + ", " + hrp_peso_util + ", " +
                                                "" + hrp_num_unidades + ", " + hrp_reman_kg + ", " +
                                                "" + hrp_reman_unidades + ", '" + tipo + "', 0, 0, '" + hrp_peso_neto + "', '', '')";
                        reader1 = cmnd1.ExecuteReader();
                        reader1.Dispose();
                    }
                    #endregion

                    #region afecta inventario pt
                    if (cbevaluacion.SelectedIndex == 0 || cbevaluacion.SelectedIndex == 1)
                    {
                        for (int i = 0; i < DGV1.Rows.Count; i++)
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "SELECT invpt_entradas_kg, invpt_entradas_un FROM TB_MSTR_INVENTARIO_PT where lin_clave ='" + txtlin_clave.Text + "' and prod_clave ='" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "' and " +
                                              "cast(substring(cast (invpt_fecha as varchar(20)),1,11) as datetime) between '" + DateTime.Now.ToShortDateString() + "' and '" + DateTime.Now.ToShortDateString() + "'";
                            reader1 = cmnd1.ExecuteReader();
                            if (reader1.HasRows)
                            {
                                while (reader1.Read())
                                {
                                    invtp_entradas_kg = Convert.ToDecimal(reader1.GetValue(0).ToString()) + Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20); ;
                                    invpt_entradas_un = Convert.ToDecimal(reader1.GetValue(1).ToString()) + Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value);
                                }
                                reader1.Dispose();

                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "update tb_mstr_inventario_pt set invpt_entradas_kg = " + invtp_entradas_kg + ", invpt_entradas_un = " + invpt_entradas_un + " where lin_clave ='" + txtlin_clave.Text + "' and prod_clave ='" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "' and " +
                                                    "cast(substring(cast (invpt_fecha as varchar(20)),1,11) as datetime) between '" + DateTime.Now.ToShortDateString() + "' and '" + DateTime.Now.ToShortDateString() + "'";
                                reader1 = cmnd1.ExecuteReader();
                                reader1.Dispose();
                                opcion = 1;
                            }
                            else
                            {
                                reader1.Dispose();
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "select top 1 invpt_inicial_kg, invpt_entradas_kg, isnull (invpt_salidas_kg, 0), invpt_inicial_un, invpt_entradas_un, isnull (invpt_salidas_un, 0) from tb_mstr_inventario_pt where lin_clave = '" + txtlin_clave.Text + "' and prod_clave ='" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "' order by invpt_fecha desc";
                                reader1 = cmnd1.ExecuteReader();
                                if (reader1.HasRows)
                                {
                                    while (reader1.Read())
                                    {
                                        var_dec_inv_inicial_kg = Convert.ToDecimal(reader1.GetValue(0).ToString()) + Convert.ToDecimal(reader1.GetValue(1).ToString()) - Convert.ToDecimal(reader1.GetValue(2).ToString());
                                        var_dec_inv_inicial_un = Convert.ToDecimal(reader1.GetValue(3).ToString()) + Convert.ToDecimal(reader1.GetValue(4).ToString()) - Convert.ToDecimal(reader1.GetValue(5).ToString());

                                        var_dec_inv_inicial_kg = Fn_peso_unitario(txtlin_clave.Text, Convert.ToString(DGV1.Rows[i].Cells[0].Value));

                                        invpt_entradas_kg = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);

                                        cmnd11 = thisConnection.CreateCommand();
                                        cmnd11.CommandText = "insert into tb_mstr_inventario_pt (invpt_fecha, lin_clave, prod_clave, invpt_inicial_un, invpt_inicial_kg, invpt_entradas_kg, invpt_entradas_un) values " +
                                                              "('" + DateTime.Now.ToShortDateString() + "', '" + txtlin_clave.Text + "', '" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "', " + var_dec_inv_inicial_un + ", " +
                                                              "" + (var_dec_inv_inicial_kg * var_dec_inv_inicial_un) + ", " + invpt_entradas_kg + ", " + Convert.ToInt32(DGV1.Rows[i].Cells[6].Value) + ")";
                                        reader11 = cmnd11.ExecuteReader();
                                        reader11.Dispose();
                                        opcion = 2;
                                    }
                                }
                                else
                                {
                                    decimal invpt_entradas_kg = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                                    cmnd11 = thisConnection.CreateCommand();
                                    cmnd11.CommandText = "insert into tb_mstr_inventario_pt (invpt_fecha, lin_clave, prod_clave, invpt_inicial_un, invpt_inicial_kg, invpt_entradas_kg, invpt_entradas_un) values " +
                                                       "('" + DateTime.Now.ToShortDateString() + "', '" + txtlin_clave.Text + "', '" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "', 0, 0, " + invpt_entradas_kg + ", " +
                                                       "" + Convert.ToInt32(DGV1.Rows[i].Cells[6].Value) + ")";
                                    reader11 = cmnd11.ExecuteReader();
                                    reader11.Dispose();
                                    opcion = 3;
                                }
                                reader1.Dispose();
                            }
                        }
                    }
                    #endregion

                    #region cambia el estatus en det_recepcion_bascula
                    if (ticket_bascula.SelectedIndex != -1)
                    {
                        for (int i = 0; i < DGV1.Rows.Count; i++)
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "update tb_det_recepcion_bascula set estatus ='R' where id_ticket = '" + Convert.ToString(DGV1.Rows[i].Cells[2].Value) + "' and num_prod = " + Convert.ToInt32(DGV1.Rows[i].Cells[15].Value) + " and prod_clave = '" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "'";
                            reader1 = cmnd1.ExecuteReader();
                            reader1.Dispose();
                        }
                    }
                    #endregion

                    #region si es nota de credito
                    if (notacre == true)
                    {

                        for (int i = 0; i < DGV1.Rows.Count; i++)
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "select rpt_recibo, lin_clave, prod_clave, rptd_cantidad, fcn_folio, afectado, rpt_Tiporecibo " +
                                                //"from tb_tmp_det_nota where prod_clave = '" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "' and rpt_recibo = '" + folnotcre + "' and rpt_Tiporecibo = '" + tipo_notcre + "'";
                                                "from tb_tmp_det_nota where prod_clave = '" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "' and rpt_recibo = '" + folnotcre + "' and fcn_folio = '" + fcn_folioNCR + "'";
                            querys += cmnd1.CommandText.ToString().Replace("'", "") + System.Environment.NewLine;
                            reader1 = cmnd1.ExecuteReader();
                            if (reader1.HasRows)
                            {
                                cmnd11 = thisConnection.CreateCommand();
                                //cmnd11.CommandText = "update tb_tmp_det_nota set afectado = 'S' where prod_clave = '" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "' and rpt_recibo = '" + folnotcre + "' and rpt_Tiporecibo ='" + tipo_notcre + "'";
                                cmnd11.CommandText = "update tb_tmp_det_nota set afectado = 'S' where prod_clave = '" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "' and rpt_recibo = '" + folnotcre + "' and fcn_folio = '" + fcn_folioNCR + "'";
                                querys += cmnd11.CommandText.ToString().Replace("'", "") + System.Environment.NewLine;
                                reader11 = cmnd11.ExecuteReader();
                                reader11.Dispose();
                            }
                            reader1.Dispose();

                            completo = true;
                            tam = 0;

                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "select rpt_recibo, lin_clave, prod_clave, rptd_cantidad, fcn_folio, afectado, rpt_Tiporecibo " +
                                                //"from tb_tmp_det_nota where rpt_recibo = '" + folnotcre + "' and rpt_Tiporecibo ='" + tipo_notcre + "'";
                                                "from tb_tmp_det_nota where rpt_recibo = '" + folnotcre + "' and fcn_folio = '" + fcn_folioNCR + "'";
                            querys += cmnd1.CommandText.ToString().Replace("'", "") + System.Environment.NewLine;
                            reader1 = cmnd1.ExecuteReader();
                            while (reader1.Read())
                            {
                                if (reader1.GetValue(5).ToString().Trim() != "S")
                                {
                                    completo = false;
                                    break;
                                }
                                tam = tam + 1;
                            }
                            reader1.Dispose();

                            if (completo == true && tam > 0)
                            {
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "update tb_tmp_mstr_nota set afectado ='S' where rpt_recibo = '" + folnotcre + "' and rpt_Tiporecibo ='" + tipo_notcre + "' and fcn_folio = '" + fcn_folioNCR + "'";
                                querys += cmnd1.CommandText.ToString().Replace("'", "") + System.Environment.NewLine;
                                reader1 = cmnd1.ExecuteReader();
                                reader1.Dispose();
                            }
                        }
                    }
                    #endregion

                    #region devolucion
                    if (txtflete.Text.Trim() == "DEV")
                    {
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "update tb_mstr_recepcion_pt set ind = 'R' where rpt_recibo = '" + txtrecibo.Text + "'";
                        reader1 = cmnd1.ExecuteReader();
                        reader1.Dispose();

                        for (int i = 0; i < DGV4.Rows.Count; i++)
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "insert into tb_det_devoluciones (rpt_recibo_nuevo, rpt_recibo_anterior, cajas, tarima, fecha, prod_clave) values " +
                                "('" + txtrecibo.Text + "', '" + Convert.ToString(DGV4.Rows[i].Cells["rec_folio"].Value).Trim() + "', '" + Convert.ToInt32(DGV4.Rows[i].Cells["boxes"].Value) + "', " +
                            "'" + Convert.ToInt32(DGV4.Rows[i].Cells["pallet"].Value) + "', '" + lbfecha.Text + "', '" + Convert.ToString(DGV4.Rows[i].Cells["cve_prod"].Value).Trim() + "')";
                            reader1 = cmnd1.ExecuteReader();
                            reader1.Dispose();

                            int hay = 0;
                            decimal resta = 0;
                            cmnd1 = thisConnection.CreateCommand();
                            //producto de campo
                            if (Convert.ToString(DGV4.Rows[i].Cells["tipo_recep"].Value) == "PTC")
                            {
                                //cmnd1.CommandText = "select etiqueta from tb_det_trazabilidad where recibo = '" + Convert.ToString(DGV4.Rows[i].Cells["rec_folio"].Value).Trim() + "' and " +
                                //                    "prod_clave = '" + Convert.ToString(DGV4.Rows[i].Cells["cve_prod"].Value).Trim() + "' " +
                                //                    "and tarima = '" + Convert.ToInt32(DGV4.Rows[i].Cells["pallet"].Value) + "'";
                                cmnd1.CommandText = "select surtido from tb_det_trazabilidad where recibo = '" + Convert.ToString(DGV4.Rows[i].Cells["rec_folio"].Value).Trim() + "' and " +
                                                    "prod_clave = '" + Convert.ToString(DGV4.Rows[i].Cells["cve_prod"].Value).Trim() + "' " +
                                                    "and tarima = '" + Convert.ToInt32(DGV4.Rows[i].Cells["pallet"].Value) + "'";
                                hay = Convert.ToInt32(cmnd1.ExecuteScalar());

                                resta = hay - Convert.ToInt32(DGV4.Rows[i].Cells["boxes"].Value);

                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "update tb_det_trazabilidad set surtido = '" + resta + "', pti_estatus_sur = '' " +
                                                    "where recibo = '" + Convert.ToString(DGV4.Rows[i].Cells["rec_folio"].Value).Trim() + "' and " +
                                                    "prod_clave = '" + Convert.ToString(DGV4.Rows[i].Cells["cve_prod"].Value).Trim() + "' " +
                                                    "and tarima = '" + Convert.ToInt32(DGV4.Rows[i].Cells["pallet"].Value) + "'";
                                reader1 = cmnd1.ExecuteReader();
                                reader1.Dispose();
                            }
                            //producto de planta
                            if (Convert.ToString(DGV4.Rows[i].Cells["tipo_recep"].Value) == "PTP")
                            {
                                //cmnd1.CommandText = "select num_cajas from tb_det_eti_final where folio = '" + Convert.ToString(DGV4.Rows[i].Cells["rec_folio"].Value).Trim() + "' and " +
                                //                    "cve_prod = '" + Convert.ToString(DGV4.Rows[i].Cells["cve_prod"].Value).Trim() + "' " +
                                //                    "and tarima = '" + Convert.ToInt32(DGV4.Rows[i].Cells["pallet"].Value) + "'";
                                cmnd1.CommandText = "select cajas_sur from tb_det_eti_final where folio = '" + Convert.ToString(DGV4.Rows[i].Cells["rec_folio"].Value).Trim() + "' and " +
                                                    "cve_prod = '" + Convert.ToString(DGV4.Rows[i].Cells["cve_prod"].Value).Trim() + "' " +
                                                    "and tarima = '" + Convert.ToInt32(DGV4.Rows[i].Cells["pallet"].Value) + "'";
                                hay = Convert.ToInt32(cmnd1.ExecuteScalar());

                                resta = hay - Convert.ToInt32(DGV4.Rows[i].Cells["boxes"].Value);

                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "update tb_det_eti_final set cajas_sur = '" + resta + "', estatus_sur = '' " +
                                                    "where folio = '" + Convert.ToString(DGV4.Rows[i].Cells["rec_folio"].Value).Trim() + "' and " +
                                                    "cve_prod = '" + Convert.ToString(DGV4.Rows[i].Cells["cve_prod"].Value).Trim() + "' " +
                                                    "and tarima = '" + Convert.ToInt32(DGV4.Rows[i].Cells["pallet"].Value) + "'";
                                reader1 = cmnd1.ExecuteReader();
                                reader1.Dispose();
                            }
                        }
                    }
                    #endregion

                    thisConnection.Close();

                    #region actualiza e inserta en Internet
                    if (flt != "")
                    {
                        try
                        {
                            mySqlConn.Open();
                            cmnd = mySqlConn.CreateCommand();
                            cmnd.CommandText = "select count(*) from tb_det_flete where id_flete = " + Convert.ToInt32(flt) + " and id_proveedor = '" + txtprov_clave.Text + "'";
                            string existe = Convert.ToString(cmnd.ExecuteScalar()).Trim();

                            if (existe.Trim() != "0")
                            {
                                for (int i = 0; i < DGV1.Rows.Count; i++)
                                {
                                    cmnd = mySqlConn.CreateCommand();
                                    cmnd.CommandText = "update tb_det_flete set num_sipgab = " + txtrecibo.Text + ", estatus= 'R' where id_flete = " + Convert.ToInt32(flt) + " and " +
                                                       "id_proveedor = '" + txtprov_clave.Text + "' and id_producto = '" + Convert.ToString(DGV1.Rows[i].Cells["produc"].Value) + "' and " +
                                                       "cantidad  ='" + Convert.ToString(DGV1.Rows[i].Cells["canti"].Value) + "'";
                                    reader = cmnd.ExecuteReader();
                                    reader.Dispose();

                                    cmnd = mySqlConn.CreateCommand();
                                    cmnd.CommandText = "insert into tb_det_sipgab (id_flete, num_sistema, cantidad, codigo, tarima, unidad, fecha_captura) values (" + Convert.ToInt32(txtflete.Text) + ", " +
                                                     "'" + txtrecibo.Text + "', " + Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value) + ", '" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "', " +
                                                     "" + Convert.ToInt32(DGV1.Rows[i].Cells[7].Value) + ", '" + Convert.ToString(DGV1.Rows[i].Cells[12].Value) + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                                    reader = cmnd.ExecuteReader();
                                    cmnd.Dispose();
                                    reader.Dispose();
                                }

                                cmnd = mySqlConn.CreateCommand();
                                cmnd.CommandText = "update tb_mstr_flete set estatus ='R', recibio = '" + txtevaluador.Text + "' where id_flete = " + Convert.ToInt32(txtflete.Text) + " and " +
                                                   " id_proveedor ='" + txtprov_clave.Text + "'";
                                readr1 = cmnd.ExecuteReader();
                            }
                            reader.Dispose();
                            mySqlConn.Close();
                        }
                        catch (SqlException ex)
                        {
                            thisConnection.Close();
                            Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.3", ex.ToString().Trim(), "PTSQL");
                            Utilerias.Class1.SendMail("ricardo.cortes@mrlucky.com.mx", "ricardo.cortes", "rcedillo", ex.ToString().Trim());
                            MessageBox.Show(ex.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    #endregion

                    #region si el producto es rechazado o sobre inspección
                    if (cbevaluacion.SelectedIndex == 2 || cbevaluacion.SelectedIndex == 1)
                    {
                        No_conformidad.recibo = txtrecibo.Text;
                        No_conformidad.nombre_proveedor = txtprov.Text;
                        No_conformidad.evaluacion = cbevaluacion.SelectedItem.ToString().Trim();
                        No_conformidad nocon = new No_conformidad();
                        foreach (DataGridViewRow row in DGV1.Rows)
                        {
                            DataRow rw = No_conformidad.productos.NewRow();
                            rw["producto"] = Convert.ToString(row.Cells["Nomb"].Value.ToString().Trim());
                            No_conformidad.productos.Rows.Add(rw);
                        }
                        nocon.ShowDialog();
                    }
                    #endregion
                }
                catch (SqlException ex)
                {
                    thisConnection.Close();
                    Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.3", ex.ToString().Trim(), "PTSQL");
                    Utilerias.Class1.SendMail("ricardo.cortes@mrlucky.com.mx", "ricardo.cortes", "rcedillo", ex.ToString().Trim());
                    MessageBox.Show(ex.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                catch (Exception ex1)
                {
                    thisConnection.Close();
                    Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.3", ex1.ToString().Trim(), "PTSQL");
                    Utilerias.Class1.SendMail("ricardo.cortes@mrlucky.com.mx", "ricardo.cortes", "rcedillo", ex1.ToString().Trim());
                    MessageBox.Show(ex1.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                #endregion

                label7.ForeColor = Color.Blue;
                label7.Update();

                #region FOX
                /*try
            {
                MyConnection.Open();
                OleDbCommand dbCmdNull1 = MyConnection.CreateCommand();
                dbCmdNull1.CommandText = "SET NULL OFF";
                dbCmdNull1.ExecuteNonQuery();
                dbCmdNull1.Dispose();

                if (txtrecibo.Text.Length < 6)
                    txtrecibo.Text = "0" + txtrecibo.Text;

                //guarda en tb_mstr_recepcion_pt
                cmd1 = MyConnection.CreateCommand();
                cmd1.CommandText = "insert into tb_mstr_recepcion_pt ([rpt_recibo], [rpt_fecha], [prov_clave], [rch_clave], [tbl_clave], [rpt_codigo], [lin_clave], [rpt_tipo], [rpt_flete], [rpt_observaciones], " +
                                    "[vari_clave], [rpt_pesador], [rpt_evaluador], [rpt_cve_fecha], [rpt_viaje], [rpt_hora], [rpt_inventario], [rpt_situacion]) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
                cmd1.Parameters.Add("@Nombre1", OleDbType.Char).Value = txtrecibo.Text;
                cmd1.Parameters.Add("@Nombre2", OleDbType.Date).Value = Convert.ToDateTime(lbfecha.Text);
                cmd1.Parameters.Add("@Nombre3", OleDbType.Char).Value = txtprov_clave.Text;
                cmd1.Parameters.Add("@Nombre4", OleDbType.Char).Value = txtrch_clave.Text;
                cmd1.Parameters.Add("@Nombre5", OleDbType.Char).Value = txttbl_clave.Text;
                cmd1.Parameters.Add("@Nombre6", OleDbType.Char).Value = txtcodigo_clave.Text;
                cmd1.Parameters.Add("@Nombre7", OleDbType.Char).Value = txtlin_clave.Text;
                cmd1.Parameters.Add("@Nombre8", OleDbType.Char).Value = tipo;
                cmd1.Parameters.Add("@Nombre9", OleDbType.Char).Value = txtflete.Text;
                cmd1.Parameters.Add("@Nombre10", OleDbType.Char).Value = txtobs.Text;
                cmd1.Parameters.Add("@Nombre11", OleDbType.Char).Value = txtvariedad.Text;
                cmd1.Parameters.Add("@Nombre12", OleDbType.Char).Value = txtpesador.Text;
                cmd1.Parameters.Add("@Nombre13", OleDbType.Char).Value = txtevaluador.Text;
                cmd1.Parameters.Add("@Nombre14", OleDbType.Char).Value = lbcletiqueta.Text;
                cmd1.Parameters.Add("@Nombre15", OleDbType.Numeric).Value = float.Parse(txtnumviaje.Text);
                cmd1.Parameters.Add("@Nombre16", OleDbType.Char).Value = txthora.Text;
                cmd1.Parameters.Add("@Nombre19", OleDbType.Char).Value = situacion;
                cmd1.Parameters.Add("@Nombre18", OleDbType.Char).Value = "C";
                read1 = cmd1.ExecuteReader();
                read1.Dispose();

                for (int i = 0; i < DGV1.Rows.Count; i++)
                {
                    dbCmdNull1 = MyConnection.CreateCommand();
                    dbCmdNull1.CommandText = "SET NULL OFF";
                    dbCmdNull1.ExecuteNonQuery();
                    dbCmdNull1.Dispose();

                    //guarda en tb_det_recepcion_pt
                    cmd1 = MyConnection.CreateCommand();
                    cmd1.CommandText = "insert into tb_det_recepcion_pt ([rpt_recibo], [lin_clave], [rptd_ticket], [prod_clave], [rptd_peso_bruto], [rptd_tara], [env_clave], [rptd_tarimas], [rptd_cantidad], " +
                                       "[rptd_fechacad]) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
                    cmd1.Parameters.Add("@Nombre1", OleDbType.Char).Value = txtrecibo.Text;
                    cmd1.Parameters.Add("@Nombre2", OleDbType.Char).Value = txtlin_clave.Text;
                    cmd1.Parameters.Add("@Nombre3", OleDbType.Char).Value = Convert.ToString(DGV1.Rows[i].Cells[2].Value);
                    cmd1.Parameters.Add("@Nombre4", OleDbType.Char).Value = Convert.ToString(DGV1.Rows[i].Cells[0].Value);
                    cmd1.Parameters.Add("@Nombre5", OleDbType.Numeric).Value = float.Parse(Convert.ToString(DGV1.Rows[i].Cells[3].Value));
                    cmd1.Parameters.Add("@Nombre6", OleDbType.Numeric).Value = float.Parse(Convert.ToString(DGV1.Rows[i].Cells[4].Value));
                    cmd1.Parameters.Add("@Nombre7", OleDbType.Char).Value = Convert.ToString(DGV1.Rows[i].Cells[12].Value);
                    cmd1.Parameters.Add("@Nombre8", OleDbType.Numeric).Value = float.Parse(Convert.ToString(DGV1.Rows[i].Cells[7].Value));
                    cmd1.Parameters.Add("@Nombre9", OleDbType.Numeric).Value = float.Parse(Convert.ToString(DGV1.Rows[i].Cells[6].Value));
                    cmd1.Parameters.Add("@Nombre10", OleDbType.Char).Value = Convert.ToString(DGV1.Rows[i].Cells[10].Value);
                    read1 = cmd1.ExecuteReader();
                    read1.Dispose();

                    //guarda en tb_hist_recepcion
                    decimal hrp_peso_neto = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                    decimal hrp_clase1 = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                    decimal hrp_peso_util = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                    decimal hrp_reman_kg = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                    int hrp_num_unidades = Convert.ToInt32(DGV1.Rows[i].Cells[6].Value);
                    int hrp_reman_unidades = Convert.ToInt32(DGV1.Rows[i].Cells[6].Value);

                    cmd1 = MyConnection.CreateCommand();
                    cmd1.CommandText = "insert into tb_hist_recepcion ([hrp_recibo], [hrp_fecha], [lin_clave], [hrp_tipo_recepcion], [hrp_estatus], [prod_clave], [hrp_peso_neto], [hrp_clase1], [hrp_peso_util], " +
                                         "[hrp_num_unidades], [hrp_reman_kg], [hrp_reman_unidades], [hrp_situacion]) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
                    cmd1.Parameters.Add("@Nombre1", OleDbType.Char).Value = txtrecibo.Text;
                    cmd1.Parameters.Add("@Nombre2", OleDbType.Date).Value = Convert.ToDateTime(lbfecha.Text);
                    cmd1.Parameters.Add("@Nombre3", OleDbType.Char).Value = txtlin_clave.Text;
                    cmd1.Parameters.Add("@Nombre4", OleDbType.Char).Value = "PTC";
                    cmd1.Parameters.Add("@Nombre5", OleDbType.Char).Value = "T";
                    cmd1.Parameters.Add("@Nombre6", OleDbType.Char).Value = Convert.ToString(DGV1.Rows[i].Cells[0].Value);
                    cmd1.Parameters.Add("@Nombre7", OleDbType.Numeric).Value = float.Parse(hrp_peso_neto.ToString());
                    cmd1.Parameters.Add("@Nombre8", OleDbType.Numeric).Value = float.Parse(hrp_clase1.ToString());
                    cmd1.Parameters.Add("@Nombre9", OleDbType.Numeric).Value = float.Parse(hrp_peso_util.ToString());
                    cmd1.Parameters.Add("@Nombre10", OleDbType.Numeric).Value = float.Parse(hrp_num_unidades.ToString());
                    cmd1.Parameters.Add("@Nombre11", OleDbType.Numeric).Value = float.Parse(hrp_reman_kg.ToString());
                    cmd1.Parameters.Add("@Nombre12", OleDbType.Numeric).Value = float.Parse(hrp_reman_unidades.ToString());
                    cmd1.Parameters.Add("@Nombre13", OleDbType.Char).Value = tipo;
                    read1 = cmd1.ExecuteReader();
                    read1.Dispose();

                    //solo actualiza el inventario si hay registro
                    if (opcion == 1)
                    {
                        cmd1 = MyConnection.CreateCommand();
                        cmd1.CommandText = "update tb_mstr_inventario_pt set invpt_entradas_kg = ?, invpt_entradas_un = ? where lin_clave ='" + txtlin_clave.Text + "' and prod_clave ='" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "'";
                        cmd1.Parameters.Add("@Nombre1", OleDbType.Numeric).Value = float.Parse(invpt_entradas_kg.ToString());
                        cmd1.Parameters.Add("@Nombre2", OleDbType.Numeric).Value = float.Parse(invpt_entradas_un.ToString());
                        read1 = cmd1.ExecuteReader();
                        read1.Dispose();
                    }
                    //el último registro del producto
                    if (opcion == 2)
                    {
                        cmd1 = MyConnection.CreateCommand();
                        cmd1.CommandText = "insert into tb_mstr_inventario_pt (invpt_fecha, lin_clave, prod_clave, invpt_inicial_un, invpt_inicial_kg, invpt_entradas_kg, invpt_entradas_un) values" +
                                         "(?, ?, ?, ?, ?, ?, ?)";
                        cmd1.Parameters.Add("@Nombre1", OleDbType.Date).Value = Convert.ToDateTime(DateTime.Now.ToShortDateString());
                        cmd1.Parameters.Add("@Nombre2", OleDbType.Char).Value = txtlin_clave.Text;
                        cmd1.Parameters.Add("@Nombre3", OleDbType.Char).Value = Convert.ToString(DGV1.Rows[i].Cells[0].Value);
                        cmd1.Parameters.Add("@Nombre4", OleDbType.Numeric).Value = float.Parse(var_dec_inv_inicial_un.ToString());
                        cmd1.Parameters.Add("@Nombre5", OleDbType.Numeric).Value = float.Parse((var_dec_inv_inicial_kg * var_dec_inv_inicial_un).ToString());
                        cmd1.Parameters.Add("@Nombre6", OleDbType.Numeric).Value = float.Parse(invpt_entradas_kg.ToString());
                        cmd1.Parameters.Add("@Nombre7", OleDbType.Numeric).Value = float.Parse(Convert.ToString(DGV1.Rows[i].Cells[6].Value));
                        read1 = cmd1.ExecuteReader();
                        read1.Dispose();
                    }
                    //si no hay registro
                    if (opcion == 3)
                    {
                        cmd1 = MyConnection.CreateCommand();
                        cmd1.CommandText = "insert into tb_mstr_inventario_pt (invpt_fecha, lin_clave, prod_clave, invpt_inicial_un, invpt_inicial_kg, invpt_entradas_kg, invpt_entradas_un) values " +
                                           "(?, ?, ?, ?, ?, ?, ?)";
                        cmd1.Parameters.Add("@Nombre1", OleDbType.Date).Value = Convert.ToDateTime(DateTime.Now.ToShortDateString());
                        cmd1.Parameters.Add("@Nombre2", OleDbType.Char).Value = txtlin_clave.Text;
                        cmd1.Parameters.Add("@Nombre3", OleDbType.Char).Value = Convert.ToString(DGV1.Rows[i].Cells[0].Value);
                        cmd1.Parameters.Add("@Nombre4", OleDbType.Numeric).Value = 0;
                        cmd1.Parameters.Add("@Nombre5", OleDbType.Numeric).Value = 0;
                        cmd1.Parameters.Add("@Nombre6", OleDbType.Numeric).Value = float.Parse(invpt_entradas_kg.ToString());
                        cmd1.Parameters.Add("@Nombre7", OleDbType.Numeric).Value = float.Parse(Convert.ToString(DGV1.Rows[i].Cells[6].Value));
                        read1 = cmd1.ExecuteReader();
                        read1.Dispose();
                    }

                    //NOTAS DE CREDITO
                    if (notacre == true)
                    {
                        cmd1 = MyConnection.CreateCommand();
                        cmd1.CommandText = "select * from tb_tmp_det_nota where prod_clave = '" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "' and rpt_recibo = '" + folnotcre + "' and rpt_Tiporecibo = '" + tipo_notcre + "'";
                        read1 = cmd1.ExecuteReader();
                        if (read1.HasRows)
                        {
                            while (read1.Read())
                            {
                                cmd1 = MyConnection.CreateCommand();
                                cmd1.CommandText = "update tb_tmp_det_nota set afectado = 'S' where prod_clave = '" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "' and rpt_recibo = '" + folnotcre + "' and rpt_Tiporecibo ='" + tipo_notcre + "'";
                                read11 = cmd1.ExecuteReader();
                                read11.Dispose();
                            }
                        }
                        read1.Dispose();

                        completo = true;
                        tam = 0;

                        cmd1 = MyConnection.CreateCommand();
                        cmd1.CommandText = "select * from tb_tmp_det_nota where rpt_recibo = '" + folnotcre + "' and rpt_Tiporecibo ='" + tipo_notcre + "'";
                        read1 = cmd1.ExecuteReader();
                        while (read1.Read())
                        {
                            if (read1.GetValue(5).ToString().Trim() != "S")
                            {
                                completo = false;
                                break;
                            }
                            tam = tam + 1;
                        }
                        read1.Dispose();
                        if (completo == true && tam > 0)
                        {
                            cmd1 = MyConnection.CreateCommand();
                            cmd1.CommandText = "update tb_tmp_mstr_nota set afectado ='S' where rpt_recibo = '" + folnotcre + "' and rpt_Tiporecibo ='" + tipo_notcre + "'";
                            read1 = cmd1.ExecuteReader();
                            read1.Dispose();
                        }
                    }
                }

                //daños
                dbCmdNull1 = MyConnection.CreateCommand();
                dbCmdNull1.CommandText = "SET NULL OFF";
                dbCmdNull1.ExecuteNonQuery();
                dbCmdNull1.Dispose();


                for (int i = 0; i < DGV2.Rows.Count; i++)
                {
                    cmd1 = MyConnection.CreateCommand();
                    cmd1.CommandText = "insert into tb_mstr_danos ([rmp_recibo], [lin_clave], [dno_clave], [dnm_cantidad], [dnm_tipo]) values (?, ?, ?, ?, ?)";
                    cmd1.Parameters.Add("@Nombre1", OleDbType.Char).Value = txtrecibo.Text;
                    cmd1.Parameters.Add("@Nombre2", OleDbType.Char).Value = linea;
                    cmd1.Parameters.Add("@Nombre3", OleDbType.Char).Value = Convert.ToString(DGV2.Rows[i].Cells[2].Value);
                    cmd1.Parameters.Add("@Nombre4", OleDbType.Numeric).Value = float.Parse(Convert.ToString(DGV2.Rows[i].Cells[1].Value));
                    cmd1.Parameters.Add("@Nombre5", OleDbType.Char).Value = "PT";
                    cmd1.Connection = MyConnection;
                    cmd1.ExecuteNonQuery();
                    cmd1.Dispose();
                }

                //procesos
                dbCmdNull1 = MyConnection.CreateCommand();
                dbCmdNull1.CommandText = "SET NULL OFF";
                dbCmdNull1.ExecuteNonQuery();
                dbCmdNull1.Dispose();

                for (int i = 0; i < DGV3.Rows.Count; i++)
                {
                    cmd1 = MyConnection.CreateCommand();
                    cmd1.CommandText = "insert into tb_mstr_procesos ([rmp_recibo], [lin_clave], [proc_clave], [prom_cantidad], [prom_tipo]) values (?, ?, ?, ?, ?)";
                    cmd1.Parameters.Add("@Nombre1", OleDbType.Char).Value = txtrecibo.Text;
                    cmd1.Parameters.Add("@Nombre2", OleDbType.Char).Value = linea;
                    cmd1.Parameters.Add("@Nombre3", OleDbType.Char).Value = Convert.ToString(DGV3.Rows[i].Cells[2].Value);
                    cmd1.Parameters.Add("@Nombre4", OleDbType.Numeric).Value = float.Parse(Convert.ToString(DGV3.Rows[i].Cells[1].Value));
                    cmd1.Parameters.Add("@Nombre5", OleDbType.Char).Value = "PT";
                    cmd1.Connection = MyConnection;
                    cmd1.ExecuteNonQuery();
                    cmd1.Dispose();
                }

                MyConnection.Close();
            }
            catch (OleDbException ex)
            {
                MyConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.3", ex.ToString().Trim(), "PTFOX");
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString().Trim());
                MessageBox.Show(ex.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            catch (Exception ex1)
            {
                MyConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.3", ex1.ToString().Trim(), "PTFOX");
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString().Trim());
                MessageBox.Show(ex1.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }*/
                #endregion

                label7.Visible = false;

                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.3", txtrecibo.Text, "NUEVO RECIBO DE PT " + txtrecibo.Text, "SIPGAB");
                if (notacre == true)
                    Utilerias.Class1.registrar_movimiento3(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.3", txtrecibo.Text, "RECIBO DE PT DE NOTACRE " + txtrecibo.Text, "SIPGAB", "", querys);

                if (cbevaluacion.SelectedIndex == 2 || cbevaluacion.SelectedIndex == 1)
                {
                    No_conformidad.recibo = txtrecibo.Text;
                    No_conformidad.nombre_proveedor = txtprov.Text;
                    No_conformidad nocon = new No_conformidad();
                    foreach (DataGridViewRow row in DGV1.Rows)
                    {
                        DataRow rw = No_conformidad.productos.NewRow();
                        rw["producto"] = Convert.ToString(row.Cells["Nomb"].Value.ToString().Trim());
                        No_conformidad.productos.Rows.Add(rw);
                    }
                    nocon.ShowDialog();
                }


                //limpiarTextBoxes(this);
                //return;
                try
                {
                    if (MessageBox.Show("¿Desea imprimir la recepción?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        PrintDocument pd = new System.Drawing.Printing.PrintDocument();
                        pd.PrintPage += new PrintPageEventHandler(this.printDocument1_PrintPage);
                        pd.Print();
                        //PrintPreviewDialog VistaPrevia = new PrintPreviewDialog();
                        //VistaPrevia.Document = printDocument1;
                        //VistaPrevia.Show();  
                    }
                }
                catch (SqlException ex)
                {
                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                    //Console.WriteLine("Line: " + trace.GetFrame(0).GetFileLineNumber());
                    MessageBox.Show("Recibo: " + txtrecibo.Text + " " + ex.ToString() + "  " + System.Environment.MachineName + " " + "Line: " + trace.GetFrame(0).GetFileLineNumber(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "Sistem@s2026$", "Recibo: " + txtrecibo.Text + " " + ex.ToString() + "  " + System.Environment.MachineName + " " + "Line: " + trace.GetFrame(0).GetFileLineNumber());
                    return;
                }
                try
                {
                    btnemail_Click(sender, e);
                }
                catch (SqlException ex)
                {
                    //thisConnection.Close();

                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                    //Console.WriteLine("Line: " + trace.GetFrame(0).GetFileLineNumber());
                    MessageBox.Show("Recibo: " + txtrecibo.Text + " " + ex.ToString() + "  " + System.Environment.MachineName + " " + "Line: " + trace.GetFrame(0).GetFileLineNumber(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "Sistem@s2026$", "Recibo: " + txtrecibo.Text + " " + ex.ToString() + "  " + System.Environment.MachineName + " " + "Line: " + trace.GetFrame(0).GetFileLineNumber());
                    return;
                }
            }
            #endregion

            #region modificar pesos
            if (opcion == 3)
            {
                try
                {
                    if (MessageBox.Show("¿Desea guardar la modificación?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        thisConnection.Open();
                        for (int i = 0; i < DGV1.Rows.Count; i++)
                        {

                            #region VALIDACION DE PESO BRUTO Y TARA SON MAYOR A 0
                            if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) < 0)
                            {
                                MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                DGV1.Focus();
                                DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                                return;
                            }

                            if (Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value) < 0)
                            {
                                MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                DGV1.Focus();
                                DGV1.Rows[i].Cells["ta"].Selected = true;
                                return;
                            }

                            if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_uni"].Value) < 0)
                            {
                                MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                DGV1.Focus();
                                DGV1.Rows[i].Cells["ta"].Selected = true;
                                return;
                            }

                            if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) <= Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value))
                            {
                                MessageBox.Show("El Peso Bruto debe ser mayor a la Tara en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                DGV1.Focus();
                                DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                                return;
                            }
                            #endregion


                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "update tb_det_recepcion_pt set peso_neto = '" + Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) + "', " +
                                                "' where rpt_recibo = '" + txtrecibo.Text + "' and " +
                                                "prod_clave = '" + Convert.ToString(DGV1.Rows[i].Cells["produc"].Value).Trim() + "'";
                            reader1 = cmnd1.ExecuteReader();
                            reader1.Dispose();

                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "update tb_hist_recepcion set hrp_surtido = '" + Convert.ToDecimal(DGV1.Rows[i].Cells["pes_tot"].Value) + "' " +
                                                "where hrp_recibo = '" + txtrecibo.Text + "' and hrp_tipo_recepcion = 'PTC' and " +
                                                "prod_clave ='" + Convert.ToString(DGV1.Rows[i].Cells["produc"].Value).Trim() + "'";
                            reader1 = cmnd1.ExecuteReader();
                            reader1.Dispose();
                        }
                        thisConnection.Close();
                        Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "M", "2.3", txtrecibo.Text, "MODIFICACION RECIBO DE PT " + txtrecibo.Text, "SIPGAB");
                        MessageBox.Show("Los pesos se han actualizado con éxito", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (SqlException ex)
                {
                    thisConnection.Close();
                    Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString().Trim());
                    MessageBox.Show(ex.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                catch (Exception ex1)
                {
                    thisConnection.Close();
                    Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString().Trim());
                    MessageBox.Show(ex1.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            #endregion

            #region modificar recibo campo
            if (opcion == 4)
            {
                for (int i = 0; i < DGV1.Rows.Count; i++)
                {
                    #region VALIDACION DE PESO BRUTO Y TARA SON MAYOR A 0
                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_uni"].Value) == 0)
                    {
                        MessageBox.Show("Favor de agregar el peso unitario en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["pes_uni"].Selected = true;
                        return;
                    }

                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["tarim"].Value) == 0)
                    {
                        MessageBox.Show("Favor de agregar número de tarimas en la fila" + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["tarim"].Selected = true;
                        return;
                    }

                    if (Convert.ToString(DGV1.Rows[i].Cells["enva_clave"].Value) == "")
                    {
                        MessageBox.Show("Favor de seleccionar el envase en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["enva"].Selected = true;
                        return;
                    }


                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_uni"].Value) < 0)
                    {
                        MessageBox.Show("El Peso Unitario debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["pes_uni"].Selected = true;
                        return;
                    }

                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_tot"].Value) < 0)
                    {
                        MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["pes_tot"].Selected = true;
                        return;
                    }

                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) < 0)
                    {
                        MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                        return;
                    }

                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value) < 0)
                    {
                        MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["ta"].Selected = true;
                        return;
                    }

                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) <= Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value))
                    {
                        MessageBox.Show("El Peso Bruto debe ser mayor a la Tara en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                        return;
                    }


                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) < 0)
                    {
                        MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                        return;
                    }

                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value) < 0)
                    {
                        MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["ta"].Selected = true;
                        return;
                    }

                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_uni"].Value) < 0)
                    {
                        MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["ta"].Selected = true;
                        return;
                    }

                    if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) <= Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value))
                    {
                        MessageBox.Show("El Peso Bruto debe ser mayor a la Tara en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DGV1.Focus();
                        DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                        return;
                    }
                    #endregion
                }
                if (MessageBox.Show("¿Desea guardar el recibo?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "update tb_mstr_recepcion_pt set id_regpt = '', tbl_vieja = '', rpt_flete = '" + txtflete.Text + "', rpt_observaciones = '" + txtobs.Text.Trim() + "', " +
                                        "rpt_pesador = '" + txtpesador.Text.Trim() + "', rpt_evaluador = '" + txtevaluador.Text.Trim() + "', rpt_recref = '', rpt_tiporef = '', " +
                                        "rpt_viaje = '" + txtnumviaje.Text + "', rpt_hora = '" + txthora.Text + "', rpt_enviado = '', id_flete = '', id_ticket = '', id_registro = '', " +
                                        "orde = '', transporte = '', numecon = '', responsable = '', ind = '', nc = '', inicio_captura = '" + DateTime.Now + "', " +
                                        "fin_captura = '" + DateTime.Now + "', um_clave = '" + CBMoneda.SelectedItem.ToString().Trim() + "', rpt_evaluacion = '', folio_no_conformidad = '', " +
                                        "rpt_estatus = '' where rpt_recibo = '" + txtrecibo.Text + "'";
                    reader1 = cmnd1.ExecuteReader();
                    reader1.Dispose();
                    thisConnection.Close();
                    Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.3.1", txtrecibo.Text, "SE MODIFICAN DATOS DEL RECIBO '" + txtrecibo.Text + "' EN LA TB_MSTR_RECEPCION_PT", "SIPGAB");


                    //actualiza informacion en tb_det_recepcion_pt
                    for (int i = 0; i < DGV1.Rows.Count; i++)
                    {
                        #region VALIDACION DE PESO BRUTO Y TARA SON MAYOR A 0
                        if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) < 0)
                        {
                            MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Focus();
                            DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                            return;
                        }

                        if (Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value) < 0)
                        {
                            MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Focus();
                            DGV1.Rows[i].Cells["ta"].Selected = true;
                            return;
                        }

                        if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_uni"].Value) < 0)
                        {
                            MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Focus();
                            DGV1.Rows[i].Cells["ta"].Selected = true;
                            return;
                        }

                        if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) <= Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value))
                        {
                            MessageBox.Show("El Peso Bruto debe ser mayor a la Tara en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Focus();
                            DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                            return;
                        }
                        #endregion

                        thisConnection.Open();
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "update tb_det_recepcion_pt set rptd_ticket = '" + Convert.ToString(DGV1.Rows[i].Cells["bascula_ticket"].Value) + "', " +
                                            "rptd_peso_bruto = '" + Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) + "', " +
                                            "rptd_tara = '" + Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value) + "', " +
                                            "env_clave = '" + Convert.ToString(DGV1.Rows[i].Cells["enva_clave"].Value) + "', " +
                                            "rptd_tarimas = '" + Convert.ToDecimal(DGV1.Rows[i].Cells["tarim"].Value) + "', " +
                                            "rptd_cantidad = '" + Convert.ToDecimal(DGV1.Rows[i].Cells["canti"].Value) + "', " +
                                            "producto = '', caj_tarimas = 0, resto = 0, peso_neto = '" + Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) + "', " +
                                            "kg_caja = 0, kg_cajas_vacias= 0, id_tabla = '', tabla = '', variedad = '" + txtvariedad.Text + "', variedad_ofc = '', linea_ofc = '', " +
                                            "observaciones = '', peso_unitario = '" + Convert.ToDecimal(DGV1.Rows[i].Cells["pes_uni"].Value) + "', postura = '', rptd_estatus = '' " +
                                            "where rpt_recibo = '" + txtrecibo.Text + "' and prod_clave = '" + Convert.ToString(DGV1.Rows[i].Cells["produc"].Value) + "'";
                        reader1 = cmnd1.ExecuteReader();
                        reader1.Dispose();
                        thisConnection.Close();
                        Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.3.1", txtrecibo.Text,
                                                              "SE MODIFICAN DATOS DEL PRODUCTO = " + DGV1.Rows[i].Cells[0].Value + " EN LA TB_DET_RECEPCION_PT", "SIPGAB");
                    }//for

                    //actualiza informacion en tb_his_recepcion
                    for (int i = 0; i < DGV1.Rows.Count; i++)
                    {
                        #region VALIDACION DE PESO BRUTO Y TARA SON MAYOR A 0
                        if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) < 0)
                        {
                            MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Focus();
                            DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                            return;
                        }

                        if (Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value) < 0)
                        {
                            MessageBox.Show("El peso debe ser mayor a 0 en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Focus();
                            DGV1.Rows[i].Cells["ta"].Selected = true;
                            return;
                        }

                        if (Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) <= Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value))
                        {
                            MessageBox.Show("El Peso Bruto debe ser mayor a la Tara en la fila " + (i + 1).ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DGV1.Focus();
                            DGV1.Rows[i].Cells["pes_bru"].Selected = true;
                            return;
                        }
                        #endregion

                        thisConnection.Open();
                        decimal hrp_peso_neto = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                        decimal hrp_clase1 = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                        decimal hrp_peso_util = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                        decimal hrp_reman_kg = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                        int hrp_num_unidades = Convert.ToInt32(DGV1.Rows[i].Cells[6].Value);
                        int hrp_reman_unidades = Convert.ToInt32(DGV1.Rows[i].Cells[6].Value);

                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "update tb_hist_recepcion set hrp_peso_neto = '" + hrp_peso_neto + "', hrp_clase1 = '" + hrp_clase1 + "', " +
                                            "hrp_peso_util = '" + hrp_peso_util + "', hrp_num_unidades = '" + hrp_num_unidades + "', hrp_reman_kg = '" + hrp_reman_kg + "', " +
                                            "hrp_reman_unidades = '" + hrp_reman_unidades + "', hrp_situacion = '" + tipo + "', hrp_clase2 = 0, " +
                                            "hrp_proceso= 0, hrp_surtido = '" + hrp_peso_neto + "', hrp_liquidado = '', hrp_numliq = '', " +
                                            "hrp_ticket = '" + Convert.ToString(DGV1.Rows[i].Cells["bascula_ticket"].Value) + "' where " +
                                            "hrp_recibo = '" + txtrecibo.Text + "' and prod_clave= '" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "'";
                        reader1 = cmnd1.ExecuteReader();
                        reader1.Dispose();
                        thisConnection.Close();
                        Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.3.1", txtrecibo.Text,
                                                              "SE MODIFICAN DATOS DEL PRODUCTO = " + DGV1.Rows[i].Cells[0].Value + " EN LA TB_HIST_RECEPCION", "SIPGAB");
                    }

                    thisConnection.Open();
                    #region guarda defectos
                    if (DGV2.Rows.Count > 0)
                    {
                        for (int i = 0; i < DGV2.Rows.Count; i++)
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "insert into tb_mstr_danos (rmp_recibo, lin_clave, dno_clave, dnm_cantidad, dnm_tipo, dnm_peso) " +
                                                "values ('" + txtrecibo.Text + "', '" + txtlin_clave.Text + "', '" + Convert.ToString(DGV2.Rows[i].Cells[2].Value) + "', " +
                                                "" + Convert.ToDecimal(DGV2.Rows[i].Cells[1].Value) + ", 'PT', 0.000)";
                            reader1 = cmnd1.ExecuteReader();
                            reader1.Dispose();
                        }
                    }
                    #endregion

                    #region guarda procesos
                    if (DGV3.Rows.Count > 0)
                    {
                        for (int i = 0; i < DGV3.Rows.Count; i++)
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "insert into tb_mstr_procesos (rmp_recibo, lin_clave, proc_clave, prom_cantidad, prom_tipo, prom_peso) values " +
                                                "('" + txtrecibo.Text + "', '" + txtlin_clave.Text + "', '" + Convert.ToString(DGV3.Rows[i].Cells[2].Value) + "', " +
                                                "" + Convert.ToDecimal(DGV3.Rows[i].Cells[1].Value) + ", 'PT', 0)";
                            reader1 = cmnd1.ExecuteReader();
                            reader1.Dispose();
                        }
                    }
                    #endregion

                    #region afecta inventario pt
                    if (cbevaluacion.SelectedIndex == 0 || cbevaluacion.SelectedIndex == 1)
                    {
                        for (int i = 0; i < DGV1.Rows.Count; i++)
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "SELECT invpt_entradas_kg, invpt_entradas_un FROM TB_MSTR_INVENTARIO_PT where lin_clave ='" + txtlin_clave.Text + "' and prod_clave ='" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "' and " +
                                              "cast(substring(cast (invpt_fecha as varchar(20)),1,11) as datetime) between '" + DateTime.Now.ToShortDateString() + "' and '" + DateTime.Now.ToShortDateString() + "'";
                            reader1 = cmnd1.ExecuteReader();
                            if (reader1.HasRows)
                            {
                                while (reader1.Read())
                                {
                                    invtp_entradas_kg = Convert.ToDecimal(reader1.GetValue(0).ToString()) + Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20); ;
                                    invpt_entradas_un = Convert.ToDecimal(reader1.GetValue(1).ToString()) + Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value);
                                }
                                reader1.Dispose();

                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "update tb_mstr_inventario_pt set invpt_entradas_kg = " + invtp_entradas_kg + ", invpt_entradas_un = " + invpt_entradas_un + " where lin_clave ='" + txtlin_clave.Text + "' and prod_clave ='" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "' and " +
                                                    "cast(substring(cast (invpt_fecha as varchar(20)),1,11) as datetime) between '" + DateTime.Now.ToShortDateString() + "' and '" + DateTime.Now.ToShortDateString() + "'";
                                reader1 = cmnd1.ExecuteReader();
                                reader1.Dispose();
                                opcion = 1;
                            }
                            else
                            {
                                reader1.Dispose();
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "select top 1 invpt_inicial_kg, invpt_entradas_kg, isnull (invpt_salidas_kg, 0), invpt_inicial_un, invpt_entradas_un, isnull (invpt_salidas_un, 0) from tb_mstr_inventario_pt where lin_clave = '" + txtlin_clave.Text + "' and prod_clave ='" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "' order by invpt_fecha desc";
                                reader1 = cmnd1.ExecuteReader();
                                if (reader1.HasRows)
                                {
                                    while (reader1.Read())
                                    {
                                        var_dec_inv_inicial_kg = Convert.ToDecimal(reader1.GetValue(0).ToString()) + Convert.ToDecimal(reader1.GetValue(1).ToString()) - Convert.ToDecimal(reader1.GetValue(2).ToString());
                                        var_dec_inv_inicial_un = Convert.ToDecimal(reader1.GetValue(3).ToString()) + Convert.ToDecimal(reader1.GetValue(4).ToString()) - Convert.ToDecimal(reader1.GetValue(5).ToString());

                                        var_dec_inv_inicial_kg = Fn_peso_unitario(txtlin_clave.Text, Convert.ToString(DGV1.Rows[i].Cells[0].Value));

                                        invpt_entradas_kg = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);

                                        cmnd11 = thisConnection.CreateCommand();
                                        cmnd11.CommandText = "insert into tb_mstr_inventario_pt (invpt_fecha, lin_clave, prod_clave, invpt_inicial_un, invpt_inicial_kg, invpt_entradas_kg, invpt_entradas_un) values " +
                                                              "('" + DateTime.Now.ToShortDateString() + "', '" + txtlin_clave.Text + "', '" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "', " + var_dec_inv_inicial_un + ", " +
                                                              "" + (var_dec_inv_inicial_kg * var_dec_inv_inicial_un) + ", " + invpt_entradas_kg + ", " + Convert.ToInt32(DGV1.Rows[i].Cells[6].Value) + ")";
                                        reader11 = cmnd11.ExecuteReader();
                                        reader11.Dispose();
                                        opcion = 2;
                                    }
                                }
                                else
                                {
                                    decimal invpt_entradas_kg = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                                    cmnd11 = thisConnection.CreateCommand();
                                    cmnd11.CommandText = "insert into tb_mstr_inventario_pt (invpt_fecha, lin_clave, prod_clave, invpt_inicial_un, invpt_inicial_kg, invpt_entradas_kg, invpt_entradas_un) values " +
                                                       "('" + DateTime.Now.ToShortDateString() + "', '" + txtlin_clave.Text + "', '" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "', 0, 0, " + invpt_entradas_kg + ", " +
                                                       "" + Convert.ToInt32(DGV1.Rows[i].Cells[6].Value) + ")";
                                    reader11 = cmnd11.ExecuteReader();
                                    reader11.Dispose();
                                    opcion = 3;
                                }
                                reader1.Dispose();
                            }
                        }
                    }
                    #endregion

                    #region cambia el estatus en det_recepcion_bascula
                    if (ticket_bascula.SelectedIndex != -1)
                    {
                        for (int i = 0; i < DGV1.Rows.Count; i++)
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "update tb_det_recepcion_bascula set estatus ='R' " +
                                                "where id_ticket = '" + Convert.ToString(DGV1.Rows[i].Cells[2].Value) + "' and " +
                                                "num_prod = " + Convert.ToInt32(DGV1.Rows[i].Cells[15].Value) + " and " +
                                                "prod_clave = '" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "'";
                            reader1 = cmnd1.ExecuteReader();
                            reader1.Dispose();
                        }
                    }
                    #endregion
                    thisConnection.Close();

                    #region actualiza e inserta en Internet
                    if (flt != "")
                    {
                        mySqlConn.Open();
                        cmnd = mySqlConn.CreateCommand();
                        cmnd.CommandText = "select count(*) from tb_det_flete where id_flete = " + Convert.ToInt32(flt) + " and id_proveedor = '" + txtprov_clave.Text + "'";
                        string existe = Convert.ToString(cmnd.ExecuteScalar()).Trim();

                        if (existe.Trim() != "0")
                        {
                            for (int i = 0; i < DGV1.Rows.Count; i++)
                            {
                                cmnd = mySqlConn.CreateCommand();
                                cmnd.CommandText = "update tb_det_flete set num_sipgab = " + txtrecibo.Text + ", estatus= 'R' where id_flete = " + Convert.ToInt32(flt) + " and " +
                                                   "id_proveedor = '" + txtprov_clave.Text + "' and id_producto = '" + Convert.ToString(DGV1.Rows[i].Cells["produc"].Value) + "' and " +
                                                   "cantidad  ='" + Convert.ToString(DGV1.Rows[i].Cells["canti"].Value) + "'";
                                reader = cmnd.ExecuteReader();
                                reader.Dispose();

                                cmnd = mySqlConn.CreateCommand();
                                cmnd.CommandText = "insert into tb_det_sipgab (id_flete, num_sistema, cantidad, codigo, tarima, unidad, fecha_captura) values (" + Convert.ToInt32(txtflete.Text) + ", " +
                                                 "'" + txtrecibo.Text + "', " + Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value) + ", '" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "', " +
                                                 "" + Convert.ToInt32(DGV1.Rows[i].Cells[7].Value) + ", '" + Convert.ToString(DGV1.Rows[i].Cells[12].Value) + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                                reader = cmnd.ExecuteReader();
                                cmnd.Dispose();
                                reader.Dispose();
                            }

                            cmnd = mySqlConn.CreateCommand();
                            cmnd.CommandText = "update tb_mstr_flete set estatus ='R', recibio = '" + txtevaluador.Text + "' where id_flete = " + Convert.ToInt32(txtflete.Text) + " and " +
                                               " id_proveedor ='" + txtprov_clave.Text + "'";
                            readr1 = cmnd.ExecuteReader();
                        }
                        reader.Dispose();
                        mySqlConn.Close();
                    }
                    #endregion

                    MessageBox.Show("Los datos del recibo han sido actualizados", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            #endregion
            limpiarTextBoxes(this);
            return;
            #region 
            /*etiqueta_verde etiver = new etiqueta_verde();
            num_viaje = Convert.ToInt32(txtnumviaje.Text);
            if (etiver.dt1.Columns.Count == 0 && etiver.dtoriginal.Columns.Count == 0)
            {
                etiver.dt1.Columns.Add("prod_clave");
                etiver.dt1.Columns.Add("prod_nombre");
                etiver.dt1.Columns.Add("numero_tarimas");
                etiver.dt1.Columns.Add("cajas_x_tarima");
                etiver.dt1.Columns.Add("feccad");

                etiver.dtoriginal.Columns.Add("prod_clave");
                etiver.dtoriginal.Columns.Add("prod_nombre");
                etiver.dtoriginal.Columns.Add("numero_tarimas");
                etiver.dtoriginal.Columns.Add("cajas_x_tarima");
                etiver.dtoriginal.Columns.Add("fec_cad");
            }

            for (int i = 0; i < DGV1.Rows.Count; i++)
            {                      
                DataRow row = etiver.dt1.NewRow();
                row[0] = DGV1.Rows[i].Cells[0].Value;
                row[1] = DGV1.Rows[i].Cells[1].Value;
                row[2] = DGV1.Rows[i].Cells[7].Value;
                row[3] = Math.Truncate(Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value) / Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value));
                row[4] = Convert.ToString(DGV1.Rows[i].Cells[10].Value);
                etiver.dt1.Rows.Add(row);

                DataRow rw = etiver.dtoriginal.NewRow();
                rw[0] = DGV1.Rows[i].Cells[0].Value;
                rw[1] = DGV1.Rows[i].Cells[1].Value;
                rw[2] = DGV1.Rows[i].Cells[7].Value;
                rw[3] = DGV1.Rows[i].Cells[6].Value;
                rw[4] = Convert.ToString(DGV1.Rows[i].Cells[10].Value);
                etiver.dtoriginal.Rows.Add(rw);

            }
            hora = txthora.Text;
            etiver.ShowDialog();
            limpiarTextBoxes(this);*/
            #endregion
        }

        //método que limpia todo
        private void limpiarTextBoxes(Control parent)
        {
            //Limpiar de manera rapida
            foreach (Control c in parent.Controls)
            {
                if (c is TextBox)
                {
                    TextBox txt = (TextBox)c;
                    txt.Text = "";
                    txt.ReadOnly = true;
                }
                if (c.Controls.Count > 0)
                {
                    limpiarTextBoxes(c);
                }
            }

            txtnumviaje.Text = "0";
            btnGuardar.Enabled = false;
            btnCancel.Enabled = false;
            btnAlta.Enabled = true;
            btnConsulta.Enabled = true;
            btncancelrecibo.Enabled = false;
            DGV1.Rows.Clear();
            DGV2.Rows.Clear();
            DGV3.Rows.Clear();
            cbticket.SelectedIndex = -1;
            cbticket.Items.Clear();
            cbtipo.SelectedIndex = 0;
            CBcodigo.SelectedIndex = -1;
            CBcodigo.Items.Clear();
            CBlin_nombre.SelectedIndex = -1;
            cbticket.Enabled = false;
            CBlin_nombre.Enabled = false;
            cbticket.Enabled = false;
            CBcodigo.Enabled = false;
            txthora.Text = DateTime.Now.ToString("HH:mm");
            cbtipo.Enabled = false;
            btnntacre.Enabled = false;
            CBMoneda.Enabled = false;
            CBMoneda.SelectedIndex = 0;
            cbrecpen.SelectedIndex = -1;
            cbrecpen.Items.Clear();
            cbrecpen.Enabled = false;
            DGV2.Visible = true;
            DGV3.Visible = true;
            DGV4.Visible = false;
            label8.Visible = false;
            txtfolio.Visible = false;
            lbguion1.Visible = false;
            txtclaveprod.Visible = false;
            lbguion2.Visible = false;
            txttar.Visible = false;
            lbnumcjs.Visible = false;
            txtnumcjs.Visible = false;
            groupBox1.Visible = false;
            btnpdf.Visible = false;
            btnaddfile.Visible = false;
            cbcontrato.SelectedIndex = -1;
            cbcontrato.Items.Clear();
            cbcontrato.Enabled = false;
            opcion = 0; recibo = 0; cantidad = 0; tarm = 0; rpt_recibo = 0; num_prod = 0; tam = 0; coor_x = 0; num_viaje = 0;
            b_p = 0; tar = 0; p_u = 0; p_t = 0; peso_env = 0; invtp_entradas_kg = 0; invpt_entradas_un = 0; invpt_inicial_kg = 0; invpt_salidas_kg = 0; invpt_inicial_un = 0; invpt_salidas_un = 0;
            var_dec_inv_inicial_kg = 0; var_dec_inv_inicial_un = 0; var_dec_peso_total = 0; invpt_entradas_kg = 0; hrp_num_unidades = 0; hrp_peso_neto = 0; hrp_reman_unidades = 0;
            varieda = ""; producto = ""; nombre = ""; fecha_cad = ""; envas = ""; provee = ""; rancho = ""; tabla = ""; linea = ""; nom_tabla = ""; tipo = "";
            env_clave = ""; situacion = "S"; cletiqueta = ""; proveedor = ""; rancho_nom = ""; hora = ""; estatus = ""; tipo_notcre = ""; folnotcre = ""; contrato = "";
        }

        private void cbtipo_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (cbtipo.SelectedIndex == 0)
                tipo = "CM";

            if (cbtipo.SelectedIndex == 2)
                tipo = "TR";

            if (cbtipo.SelectedIndex == 1)
                tipo = "MA";
        }

        private void DGV1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //producto
            if (DGV1.CurrentCell.ColumnIndex == 1)
            {
                nom_prod = e.Control as ComboBox;
                if (env != null)
                {
                    nom_prod.SelectedIndexChanged -= new EventHandler(ComboBox_SelectedIndexChanged);
                    nom_prod.SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);
                }
            }
            //envases
            if (DGV1.CurrentCell.ColumnIndex == 5)
            {
                env = e.Control as ComboBox;
                if (env != null)
                {
                    env.SelectedIndexChanged -= new EventHandler(ComboBox_SelectedIndexChanged);
                    env.SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);
                }
            }
            //variedad
            if (DGV1.CurrentCell.ColumnIndex == 11)
            {
                varie = e.Control as ComboBox;
                if (varie != null)
                {
                    varie.SelectedIndexChanged -= new EventHandler(ComboBox_SelectedIndexChanged);
                    varie.SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);
                }
            }
        }

        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //producto
            if (DGV1.CurrentCell.ColumnIndex == 1)
            {
                foreach (DataRow row in productos.Select("prod_nombre = '" + nom_prod.SelectedItem.ToString() + "'"))
                {
                    producto = Convert.ToString(row["prod_clave"].ToString().Trim());
                    DGV1.Rows[DGV1.CurrentCell.RowIndex].Cells[0].Value = Convert.ToString(row["prod_clave"].ToString().Trim());
                }
                /*string pro = nom_prod.SelectedItem.ToString();
                pro = pro.Replace("'","''");
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "select prod_clave from tb_cat_producto where prod_nombre = '" + pro + "'";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    producto = reader1.GetValue(0).ToString().Trim();
                    DGV1.Rows[DGV1.CurrentCell.RowIndex].Cells[0].Value = reader1.GetValue(0).ToString().Trim();
                }
                reader1.Dispose();
                thisConnection.Close();*/
            }
            //envases
            if (DGV1.CurrentCell.ColumnIndex == 5)
            {
                foreach (DataRow row in envases.Select("env_nombre = '" + env.SelectedItem.ToString() + "'"))
                {
                    envas = Convert.ToString(row["env_clave"].ToString().Trim());
                    DGV1.Rows[DGV1.CurrentCell.RowIndex].Cells[12].Value = Convert.ToString(row["env_clave"].ToString().Trim());
                    DGV1.Rows[DGV1.CurrentCell.RowIndex].Cells[14].Value = Convert.ToString(row["env_peso"].ToString().Trim());
                    peso_env = Convert.ToDecimal(row["env_peso"].ToString().Trim());
                }

                /*thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "select env_clave, env_peso from tb_cat_envases where env_nombre = '" + env.SelectedItem.ToString() + "'";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                     envas = reader1.GetValue(0).ToString().Trim();
                     DGV1.Rows[DGV1.CurrentCell.RowIndex].Cells[12].Value = reader1.GetValue(0).ToString().Trim();
                     DGV1.Rows[DGV1.CurrentCell.RowIndex].Cells[14].Value = reader1.GetValue(1).ToString().Trim();
                     peso_env = Convert.ToDecimal(reader1.GetValue(1).ToString());
                }
                reader1.Dispose();
                thisConnection.Close();*/
            }
            //variedad
            if (DGV1.CurrentCell.ColumnIndex == 11)
            {
                foreach (DataRow row in variedades.Select("vari_nombre ='" + varie.SelectedItem.ToString() + "' and lin_clave = '" + txtlin_clave.Text + "'"))
                {
                    varieda = Convert.ToString(row["vari_clave"].ToString().Trim());
                    DGV1.Rows[DGV1.CurrentCell.RowIndex].Cells[13].Value = Convert.ToString(row["vari_clave"].ToString().Trim());
                }
                /*thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "select vari_clave from tb_cat_variedad where vari_nombre = '" + varie.SelectedItem.ToString() + "' and lin_clave = '" + txtlin_clave.Text + "'";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    varieda = reader1.GetValue(0).ToString().Trim();
                    DGV1.Rows[DGV1.CurrentCell.RowIndex].Cells[13].Value = reader1.GetValue(0).ToString().Trim();
                }
                reader1.Dispose();
                thisConnection.Close();   */
            }
        }

        public decimal Fn_peso_unitario(string lin_clave, string prod_clave)
        {
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "select compp_peso, um_clave from tb_mstr_comp_prod where lin_clave = '" + lin_clave + "' and prod_clave = '" + prod_clave + "'";
            reader1 = cmnd1.ExecuteReader();
            while (reader1.Read())
            {
                cmnd11 = thisConnection.CreateCommand();
                cmnd11.CommandText = "select um_equivalencia from tb_cat_unidad where um_clave = '" + reader1.GetValue(1).ToString() + "'";
                reader11 = cmnd11.ExecuteReader();
                if (reader11.HasRows)
                {
                    if (reader11.Read())
                    {
                        var_dec_peso_total = var_dec_peso_total + (Convert.ToDecimal(reader1.GetValue(0).ToString()) * Convert.ToDecimal(reader11.GetValue(0).ToString()));
                    }
                }
                else
                {
                    var_dec_peso_total = var_dec_peso_total + (Convert.ToDecimal(reader1.GetValue(0).ToString()) * 1);
                }
                reader11.Dispose();
            }
            if (reader1.HasRows == false)
                var_dec_peso_total = 1;

            reader11.Dispose();
            return var_dec_peso_total;
        }

        private void CBIngProd_fis_CheckedChanged(object sender, EventArgs e)
        {
            if (CBIngProd_fis.Checked == true)
                situacion = "S";

            else
                situacion = "N";
        }

        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            SolidBrush color = new SolidBrush(Color.Black);
            Font fuente = new Font("Courier", 10);
            Font fuente2 = new Font("Courier", 8);

            //recibo
            Point recib = new Point(650, 113);
            e.Graphics.DrawString(txtrecibo.Text, fuente, color, recib);


            //flete
            Point fle = new Point(670, 172);
            e.Graphics.DrawString(txtflete.Text, fuente, color, fle);

            //ticket bascula            
            //Point tic = new Point(670, 210);
            //e.Graphics.DrawString(ticket_bascula.SelectedItem.ToString().Trim(), fuente, color, tic);

            //nombre proveedor
            Point prove = new Point(25, 180);
            e.Graphics.DrawString(txtprov.Text, fuente, color, prove);

            //rancho
            Point rch = new Point(25, 270);
            e.Graphics.DrawString(txtrancho.Text, fuente, color, rch);

            //tabla
            Point tbl = new Point(300, 270);
            e.Graphics.DrawString(txttabla.Text, fuente, color, tbl);

            //fecha
            Point day = new Point(630, 270);
            e.Graphics.DrawString(lbfecha.Text, fuente, color, day);

            int cont = 340;
            //grid de los productos
            for (int i = 0; i < DGV1.Rows.Count; i++)
            {
                Point produc = new Point(25, cont);
                string pr = "";
                if (Convert.ToString(DGV1.Rows[i].Cells[1].Value).Length > 32)
                    pr = Convert.ToString(DGV1.Rows[i].Cells[1].Value).Substring(0, 31);
                else
                    pr = Convert.ToString(DGV1.Rows[i].Cells[1].Value);

                e.Graphics.DrawString(pr, fuente, color, produc);

                Point tick = new Point(300, cont);
                e.Graphics.DrawString(Convert.ToString(DGV1.Rows[i].Cells[2].Value), fuente, color, tick);

                Point envas = new Point(370, cont);
                e.Graphics.DrawString(Convert.ToString(DGV1.Rows[i].Cells[5].Value).Substring(0, 10), fuente, color, envas);

                coordenada_x(Convert.ToString(DGV1.Rows[i].Cells[6].Value).Length, "unidades");
                Point cant = new Point(coor_x, cont);
                e.Graphics.DrawString(Convert.ToString(DGV1.Rows[i].Cells[6].Value), fuente, color, cant);

                coordenada_x(Convert.ToString(DGV1.Rows[i].Cells[8].Value).Length, "peso_unitario");
                Point pes_uni = new Point(coor_x, cont);
                e.Graphics.DrawString(Convert.ToDecimal(DGV1.Rows[i].Cells[8].Value).ToString("###,###,###.00"), fuente, color, pes_uni);

                coordenada_x(Convert.ToString(DGV1.Rows[i].Cells[9].Value).Length, "peso_total");
                Point pes_tot = new Point(coor_x, cont);
                e.Graphics.DrawString(Convert.ToDecimal(DGV1.Rows[i].Cells[9].Value).ToString("###,###,###.00"), fuente, color, pes_tot);

                cont = cont + 13;
            }
            // SE IMPRIME LA INFORMACION DEL GRADO 1
            cont = 650;
            Point numpiezas = new Point(25, cont);
            e.Graphics.DrawString("Piezas de la Evaluacion: " + TxtPieza.Text, fuente, color, numpiezas);
            numpiezas = new Point(400, cont);
            e.Graphics.DrawString("% Grado 1 " + LblGra1.Text, fuente, color, numpiezas);
            #region imprime defectos
            if (DGV2.Rows.Count > 1)
            {
                cont = 749;
                DGV2.Sort(DGV2.Columns[1], ListSortDirection.Descending);
                for (int i = 0; i < DGV2.Rows.Count; i++)
                {
                    //if (Convert.ToInt32(DGV2.Rows[i].Cells[1].Value) <= 0)
                    //    continue;
                    Point nom_defecto = new Point(25, cont);
                    e.Graphics.DrawString(Convert.ToString(DGV2.Rows[i].Cells[0].Value), fuente2, color, nom_defecto);

                    Point cantidad_defecto = new Point(250, cont);
                    e.Graphics.DrawString(Convert.ToString(DGV2.Rows[i].Cells[1].Value), fuente2, color, cantidad_defecto);
                    cantidad_defecto = new Point(350, cont);
                    e.Graphics.DrawString(Convert.ToDecimal(DGV2.Rows[i].Cells[3].Value).ToString("###.#0"), fuente2, color, cantidad_defecto);
                    cont = cont + 13;
                }
            }
            #endregion

            #region imprime procesos
            if (DGV3.Rows.Count > 1)
            {
                cont = 892;
                DGV3.Sort(DGV3.Columns[1], ListSortDirection.Descending);
                for (int i = 0; i < DGV3.Rows.Count; i++)
                {
                    //if (Convert.ToInt32(DGV3.Rows[i].Cells[1].Value) <= 0)
                    //    continue;
                    Point nom_proceso = new Point(25, cont);
                    e.Graphics.DrawString(Convert.ToString(DGV3.Rows[i].Cells[0].Value), fuente2, color, nom_proceso);

                    Point cantidad_proceso = new Point(250, cont);
                    e.Graphics.DrawString(Convert.ToString(DGV3.Rows[i].Cells[1].Value), fuente2, color, cantidad_proceso);
                    cantidad_proceso = new Point(350, cont);
                    e.Graphics.DrawString(Convert.ToDecimal(DGV3.Rows[i].Cells[3].Value).ToString("##0.#0"), fuente2, color, cantidad_proceso);
                    cont = cont + 13;
                }
            }
            #endregion

            //tipo de recepción
            Point tipo_recep = new Point(25, 610);
            string type = cbtipo.SelectedItem.ToString().Trim();
            int lon_typy = (type.Length) - 6;
            type = type.Substring(6, lon_typy);
            e.Graphics.DrawString(type, fuente, color, tipo_recep);

            //observaciones                        
            Point obser = new Point(425, 750);
            e.Graphics.DrawString(txtobs.Text, fuente, color, obser);

            //pesador
            //e.Graphics.RotateTransform(270);            
            Point pesador = new Point(180, 610);
            //Point pesador = new Point(-500, 200);            
            e.Graphics.DrawString(txtpesador.Text, fuente, color, pesador);

            if (cbevaluacion.SelectedIndex == 0)
                e.Graphics.DrawString("X", fuente, color, 620, 610);

            if (cbevaluacion.SelectedIndex == 1)
                e.Graphics.DrawString("X", fuente, color, 620, 640);

            if (cbevaluacion.SelectedIndex == 2)
                e.Graphics.DrawString("X", fuente, color, 620, 670);

            //evaluador 
            Point evaluadro = new Point(430, 610);
            e.Graphics.DrawString(txtevaluador.Text, fuente, color, evaluadro);
            //variedad
            if (cbvariedad.SelectedIndex != -1)
            {
                Point varieda = new Point(405, 800);
                e.Graphics.DrawString("Variedad: " + cbvariedad.SelectedItem.ToString(), fuente, color, varieda);
            }
        }

        private void PBbtnsalir_Click(object sender, EventArgs e)
        {
            try
            {
                if (thisConnection.State != ConnectionState.Open)
                    thisConnection.Open();

                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT TOP 1 inicio_sesion, usu_login FROM tb_cat_historial_dia where nombre_maquina = '" + Environment.MachineName + "' and sistema = 'SIPGAB' ORDER BY inicio_sesion desc";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    Utilerias.Class1.Inicio_sesion = reader1.GetSqlDateTime(0).Value;
                    Utilerias.Class1.Usu_login = reader1.GetSqlString(1).ToString();
                    Utilerias.Class1.Nombre_equipo = Environment.MachineName;
                }
                reader1.Close();

                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "update tb_cat_historial_dia set formulario = ' ' where nombre_maquina ='" + Utilerias.Class1.Nombre_equipo + "' and usu_login = '" + Utilerias.Class1.Usu_login + "' and inicio_sesion = '" + Utilerias.Class1.Inicio_sesion.ToString("s") + "' and sistema = 'SIPGAB'";
                reader1 = cmnd1.ExecuteReader();
                reader1.Close();
                thisConnection.Close();
                Application.Exit();
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            limpiarTextBoxes(this);
            btnimpre.Visible = false;
            btnemail.Visible = false;
            btncancelrecibo.Enabled = false;
            btnpesoxtar.Enabled = false;
            btnmodi.Enabled = false;
            btnmodeva.Enabled = false;
            btnpesoxtar.Visible = false;
            PnlProdDev.Visible = false;
            LblObsCan.Visible = false;
            TxtObsCan.Visible = false;
            LblGra1.Text = "0";
        }

        public int coordenada_x(int longitud, string columna)
        {
            #region unidades
            if (columna == "unidades")
            {
                if (longitud == 1)
                    coor_x = 535;

                if (longitud == 2)
                    coor_x = 528;

                if (longitud == 3)
                    coor_x = 522;

                if (longitud == 4)
                    coor_x = 413;

                if (longitud == 5)
                    coor_x = 506;

                if (longitud == 6)
                    coor_x = 499;

                if (longitud == 7)
                    coor_x = 492;
            }
            #endregion

            #region peso unitario
            if (columna == "peso_unitario")
            {
                if (longitud == 1)
                    coor_x = 645;

                if (longitud == 2)
                    coor_x = 638;

                if (longitud == 3)
                    coor_x = 630;

                if (longitud == 4)
                    coor_x = 623;

                if (longitud == 5)
                    coor_x = 616;

                if (longitud == 6)
                    coor_x = 609;

                if (longitud == 7)
                    coor_x = 598;
            }
            #endregion 

            #region peso total
            if (columna == "peso_total")
            {
                if (longitud == 1)
                    coor_x = 755;

                if (longitud == 2)
                    coor_x = 748;

                if (longitud == 3)
                    coor_x = 740;

                if (longitud == 4)
                    coor_x = 733;

                if (longitud == 5)
                    coor_x = 726;

                if (longitud == 6)
                    coor_x = 719;

                if (longitud == 7)
                    coor_x = 706;

                if (longitud == 8)
                    coor_x = 708;

                if (longitud == 9)
                    coor_x = 700;
            }
            #endregion

            return coor_x;
        }

        private void btnntacre_Click(object sender, EventArgs e)
        {
            DGV1.Rows.Clear();
            DGV2.Rows.Clear();
            DGV3.Rows.Clear();
            fcn_folioNCR = "";
            Notas_credito nc = new Notas_credito();
            nc.ShowDialog();
            if (nc.nomline.Trim() != "")
            {
                for (int i = 0; i < nc.ntc.Rows.Count; i++)
                {
                    //foreach (DataRow row in productos.Select("prod_clave = '" + nc.ntc.Rows[i][0].ToString().Trim() + "'"))
                    //{
                    //    foreach (DataRow row1 in envases.Select("env_clave = '" + Convert.ToString(row["prod_presentacion"].ToString()) + "'"))
                    //    {
                    //        //DGV1.Rows.Add(nc.ntc.Rows[i][0].ToString(), nc.ntc.Rows[i][1].ToString(), "0", "0", "0", nc.ntc.Rows[i][2].ToString(), nc.ntc.Rows[i][3].ToString(), "0", "0", "0", "", "", "", "", "", "");
                    //        DGV1.Rows.Add(nc.ntc.Rows[i][0].ToString(), nc.ntc.Rows[i][1].ToString(), "0", "0", "0", row1["env_nombre"].ToString(), nc.ntc.Rows[i][3].ToString(), "0", Convert.ToDecimal(row["prod_peso"].ToString()).ToString("###,##0.00"), "0", "", "", "", "", "", "");
                    //        DataRow rw = not_cre.NewRow();
                    //        rw["producto"] = nc.ntc.Rows[i][0].ToString().Trim();
                    //        rw["cantidad"] = nc.ntc.Rows[i][3].ToString();
                    //        not_cre.Rows.Add(rw);
                    //    }
                    //}
                    DGV1.Rows.Add(nc.ntc.Rows[i][0].ToString(), nc.ntc.Rows[i][1].ToString(), "0", "0", "0", "ENVASIN", nc.ntc.Rows[i][3].ToString(), "0", 0, "0", "", "", "", "", "", "");
                    DataRow rw = not_cre.NewRow();
                    rw["producto"] = nc.ntc.Rows[i][0].ToString().Trim();
                    rw["cantidad"] = nc.ntc.Rows[i][3].ToString();
                    not_cre.Rows.Add(rw);
                }
                CBMoneda.SelectedIndex = 0;
                CBMoneda.Enabled = true;
                txtobs.Text = nc.moti;
                txtlin_clave.Text = nc.line;
                linea = txtlin_clave.Text;
                CBlin_nombre.SelectedItem = nc.nomline.Trim();
                cbtipo.SelectedIndex = 2;
                cbtipo_SelectionChangeCommitted(sender, e);
                txtflete.Text = "N.CRED";
                txtprov_clave.Text = nc.proveclave;
                txtprov.Text = nc.provenom;
                folnotcre = nc.rptrecibo;

                //trae los defectos de acuerdo a la linea
                foreach (DataRow row in danos.Select("lin_clave = '" + linea + "' and dno_tipo = 'PT'"))
                    DGV2.Rows.Add(Convert.ToString(row["dno_nombre"].ToString().Trim()), 0, Convert.ToString(row["dno_clave"].ToString().Trim()));

                /*thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "select dno_nombre, dno_clave from tb_cat_danos where lin_clave = '" + linea + "' and dno_tipo = 'PT' order by dno_clave";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    DGV2.Rows.Add(reader1.GetValue(0).ToString().Trim(), 0, reader1.GetValue(1).ToString().Trim());
                }
                reader1.Dispose();*/

                //trae los procesos de acuerdo a la linea
                foreach (DataRow row in procesos.Select("lin_clave = '" + linea + "' and proc_tipo = 'PT'"))
                    DGV3.Rows.Add(Convert.ToString(row["proc_nombre"].ToString().Trim()), 0, Convert.ToString(row["proc_clave"].ToString().Trim()));

                /*cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "select proc_nombre, proc_clave from tb_cat_procesos where lin_clave = '" + linea + "' and proc_tipo ='PT' order by proc_clave";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    DGV3.Rows.Add(reader1.GetValue(0).ToString().Trim(), 0, reader1.GetValue(1).ToString().Trim());
                }
                reader1.Dispose();*/

                foreach (DataRow row in lineas.Select("lin_clave = '" + linea + "'"))
                    CBlin_nombre.SelectedItem = Convert.ToString(row["lin_nombre"].ToString().Trim());

                /*cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "select lin_nombre from tb_cat_linea where lin_clave = '" + linea + "'";
                CBlin_nombre.SelectedItem = Convert.ToString(cmnd1.ExecuteScalar()).Trim();
                thisConnection.Close();*/

                CBcodigo.Enabled = true;
                //CBlin_nombre.Enabled = true;
                txtcodigo_clave.ReadOnly = false;
                //txtlin_clave.ReadOnly = false;
                txtnumviaje.ReadOnly = false;
                CBIngProd_fis.Enabled = true;
                txtevaluador.ReadOnly = false;
                txtobs.ReadOnly = false;
                txthora.ReadOnly = false;
                txtpesador.ReadOnly = false;
                txtrch_clave.ReadOnly = false;
                txttbl_clave.ReadOnly = false;
                DGV2.Columns[1].ReadOnly = false;
                DGV3.Columns[1].ReadOnly = false;
                env.SelectedItem = envas;
                //DGV2.Visible = false;
                //DGV3.Visible = false;
                //label8.Visible = true;
                //txtfolio.Visible = true;
                //lbguion1.Visible = true;
                //txtclaveprod.Visible = true;
                //lbguion2.Visible = true;
                //txttar.Visible = true;
                //lbnumcjs.Visible = true;
                //txtnumcjs.Visible = true;
                //groupBox1.Visible = true;
                txtclaveprod.Text = DGV1.Rows[0].Cells["produc"].Value.ToString().Trim();
                DGV1.Rows[0].Selected = true;
                DGV1.Update();
            }
            else
            {
                txtprov_clave.Text = "";
                txtprov.Text = "";
                txtflete.Text = "";
                cbtipo.SelectedIndex = 0;
                txtlin_clave.Text = "";
                CBlin_nombre.SelectedIndex = -1;
                txtobs.Text = "";
            }
        }

        private void DGV1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (DGV1.CurrentCell.ColumnIndex == 3)
            //DGV1.CurrentCell.ColumnIndex == 4 || DGV1.CurrentCell.ColumnIndex == 6 || DGV1.CurrentCell.ColumnIndex == 7)
            {
                p_u = 0;
                p_t = 0;
                //           peso bruto                                         tara                                                      Cantidad                                                      Tarima
                p_t = Convert.ToDecimal(DGV1.CurrentRow.Cells[3].Value) - Convert.ToDecimal(DGV1.CurrentRow.Cells[4].Value) - (Convert.ToDecimal(DGV1.CurrentRow.Cells[6].Value) * peso_env) - (Convert.ToDecimal(DGV1.CurrentRow.Cells[7].Value) * 20);
                DGV1.CurrentRow.Cells[9].Value = p_t.ToString("###,###.00");

                if (Convert.ToInt32(DGV1.CurrentRow.Cells[6].Value) > 0)
                    p_u = p_t / Convert.ToInt32(DGV1.CurrentRow.Cells[6].Value);
                else
                    p_u = 0;

                DGV1.CurrentRow.Cells[8].Value = p_u.ToString("###,###.00");
            }
            if (DGV1.CurrentCell.ColumnIndex == 4)
            {
                p_u = 0;
                p_t = 0;

                p_t = Convert.ToDecimal(DGV1.CurrentRow.Cells[3].Value) - Convert.ToDecimal(DGV1.CurrentRow.Cells[4].Value) - (Convert.ToDecimal(DGV1.CurrentRow.Cells[6].Value) * peso_env) - (Convert.ToDecimal(DGV1.CurrentRow.Cells[7].Value) * 20);
                DGV1.CurrentRow.Cells[9].Value = p_t.ToString("###,###.00");

                if (Convert.ToInt32(DGV1.CurrentRow.Cells[6].Value) > 0)
                    p_u = p_t / Convert.ToInt32(DGV1.CurrentRow.Cells[6].Value);
                else
                    p_u = 0;

                DGV1.CurrentRow.Cells[8].Value = p_u.ToString("###,###.00");
            }
            if (DGV1.CurrentCell.ColumnIndex == 6)
            {
                p_u = 0;
                p_t = 0;

                p_t = Convert.ToDecimal(DGV1.CurrentRow.Cells[3].Value) - Convert.ToDecimal(DGV1.CurrentRow.Cells[4].Value) - (Convert.ToDecimal(DGV1.CurrentRow.Cells[6].Value) * peso_env) - (Convert.ToDecimal(DGV1.CurrentRow.Cells[7].Value) * 20);
                DGV1.CurrentRow.Cells[9].Value = p_t.ToString("###,###.00");

                if (Convert.ToInt32(DGV1.CurrentRow.Cells[6].Value) > 0)
                    p_u = p_t / Convert.ToInt32(DGV1.CurrentRow.Cells[6].Value);
                else
                    p_u = 0;

                DGV1.CurrentRow.Cells[8].Value = p_u.ToString("###,###.00");

                foreach (DataRow rw in not_cre.Select("producto = '" + Convert.ToString(DGV1.CurrentRow.Cells[0].Value).Trim() + "'"))
                    rw["cantidad"] = Convert.ToInt32(DGV1.CurrentRow.Cells[6].Value);

            }
            if (DGV1.CurrentCell.ColumnIndex == 7)
            {
                p_u = 0;
                p_t = 0;

                p_t = Convert.ToDecimal(DGV1.CurrentRow.Cells[3].Value) - Convert.ToDecimal(DGV1.CurrentRow.Cells[4].Value) - (Convert.ToDecimal(DGV1.CurrentRow.Cells[6].Value) * peso_env) - (Convert.ToDecimal(DGV1.CurrentRow.Cells[7].Value) * 20);
                DGV1.CurrentRow.Cells[9].Value = p_t.ToString("###,###.00");

                if (Convert.ToInt32(DGV1.CurrentRow.Cells[6].Value) > 0)
                    p_u = p_t / Convert.ToInt32(DGV1.CurrentRow.Cells[6].Value);
                else
                    p_u = 0;

                DGV1.CurrentRow.Cells[8].Value = p_u.ToString("###,###.00");
            }
            if (DGV1.CurrentCell.ColumnIndex == 10)
            {
                string fe = Convert.ToString(DGV1.CurrentRow.Cells[10].Value).Trim();
                if (EsFecha(fe))
                {
                    MessageBox.Show("Fecha correcta");
                }
                else
                {
                    MessageBox.Show("La fecha es incorrecta, favor de verificarla", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    DGV1.CurrentRow.Cells[10].Value = "";
                    return;
                }
            }
            if (Convert.ToDecimal(DGV1.CurrentRow.Cells[8].Value) < 0)
            {
                MessageBox.Show("Revisar Peso Bruto y Tara porque el PESO UNITARIO es NEGATIVO!!", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DGV1.CurrentRow.Cells[DGV1.CurrentCell.ColumnIndex].Selected = true;
                return;
            }
        }

        public static Boolean EsFecha(String fecha)
        {
            try
            {
                DateTime.Parse(fecha);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void btnConsulta_Click(object sender, EventArgs e)
        {
            txtrecibo.ReadOnly = false;
            btnConsulta.Enabled = false;
            btnGuardar.Enabled = false;
            btnntacre.Enabled = false;
            btnAlta.Enabled = false;
            btnCancel.Enabled = true;
            txtrecibo.Focus();
            enva.Items.Clear();
            env.Items.Clear();
            DGV1.ReadOnly = true;
            opcion = 4;
        }

        private void txtrecibo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtrecibo.Text.Length > 0)
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    recibo = Convert.ToInt32(txtrecibo.Text);
                    CBlin_nombre.SelectedIndex = -1;
                    cbvariedad.SelectedIndex = -1;
                    cbvariedad.Items.Clear();
                    DGV1.Rows.Clear();
                    DGV2.Rows.Clear();
                    DGV3.Rows.Clear();
                    DGV4.Rows.Clear();
                    LblGra1.Text = "0";
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select A.rpt_fecha, A.prov_clave, A.rch_clave, A.tbl_clave, A.rpt_codigo, A.lin_clave, A.rpt_tipo, A.rpt_estatus, " +
                                        "A.rpt_flete, A.rpt_observaciones, A.rpt_pesador, A.rpt_evaluador, A.rpt_cve_fecha, A.rpt_viaje, A.rpt_hora, " +
                                        "A.rpt_inventario, A.rpt_codigo, B.prov_nombre, C.lin_nombre, A.vari_clave, A.um_clave, b.prov_email, A.rpt_evaluacion, " +
                                        "A.folio_no_conformidad, B.prov_email, A.numero_pedimento, A.contrato_proveedor, inicio_captura, a.rpt_piezasEvaluacion from tb_mstr_recepcion_pt A, " +
                                        "tb_cat_proveedor B, tb_cat_linea C where A.rpt_recibo ='" + txtrecibo.Text + "' and B.prov_clave = A.prov_clave " +
                                        "and C.lin_clave = A.lin_clave";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        lbfecha.Text = reader1.GetValue(0).ToString().Substring(0, 10);
                        txtprov_clave.Text = reader1.GetValue(1).ToString().Trim();
                        provee = txtprov_clave.Text;
                        txtrch_clave.Text = reader1.GetValue(2).ToString().Trim();
                        rancho = txtrch_clave.Text;
                        txttbl_clave.Text = reader1.GetValue(3).ToString().Trim();
                        tabla = txttbl_clave.Text;
                        txtcodigo_clave.Text = reader1.GetValue(4).ToString().Trim();
                        txtlin_clave.Text = reader1.GetValue(5).ToString().Trim();
                        tipo = reader1.GetValue(6).ToString().Trim();
                        estatus = reader1.GetValue(7).ToString().Trim();
                        txtflete.Text = reader1.GetValue(8).ToString().Trim();
                        txtobs.Text = reader1.GetValue(9).ToString().Trim();
                        txtpesador.Text = reader1.GetValue(10).ToString().Trim();
                        txtevaluador.Text = reader1.GetValue(11).ToString().Trim();
                        lbcletiqueta.Text = reader1.GetValue(12).ToString().Trim();
                        txtnumviaje.Text = reader1.GetValue(13).ToString().Trim();
                        txthora.Text = reader1.GetValue(14).ToString().Trim();
                        situacion = reader1.GetValue(15).ToString().Trim();
                        txtcodigo_clave.Text = reader1.GetValue(16).ToString().Trim();
                        txtprov.Text = reader1.GetValue(17).ToString().Trim();
                        CBlin_nombre.SelectedItem = reader1.GetValue(18).ToString().Trim();
                        CBcodigo.Items.Add(reader1.GetValue(4).ToString().Trim());
                        CBcodigo.SelectedItem = reader1.GetValue(4).ToString().Trim();
                        LblHrCap.Text = reader1["inicio_captura"].ToString();
                        if (LblHrCap.Text.Trim().Length > 12)
                            HrCaptura = LblHrCap.Text.Substring(11);
                        txtvariedad.Text = reader1.GetValue(19).ToString().Trim();
                        if (reader1.GetValue(20).ToString().Trim().Contains("PESOS"))
                            CBMoneda.SelectedIndex = 0;
                        else
                            CBMoneda.SelectedIndex = 1;

                        if (tipo.Equals("CM"))
                            cbtipo.SelectedIndex = 0;
                        if (tipo.Equals("MA"))
                            cbtipo.SelectedIndex = 1;
                        if (tipo.Equals("TR"))
                            cbtipo.SelectedIndex = 2;

                        if (situacion.Equals("S"))
                            CBIngProd_fis.Checked = true;

                        if (situacion.Equals("") || situacion.Equals("N"))
                            CBIngProd_fis.Checked = false;

                        email = reader1.GetValue(21).ToString().Trim();
                        cbevaluacion.SelectedItem = reader1.GetValue(22).ToString().Trim();
                        txtfolnocon.Text = reader1.GetValue(23).ToString().Trim();
                        email = reader1.GetValue(24).ToString().Trim();
                        txtnumped.Text = reader1.GetValue(25).ToString();
                        contrato = reader1.GetValue(26).ToString().Trim();
                        TxtPieza.Text = Convert.ToString(reader1["rpt_piezasEvaluacion"]);
                    }
                    reader1.Dispose();
                    cbcontrato.Items.Clear();
                    cbcontrato.Items.Add(contrato);
                    cbcontrato.SelectedIndex = 0;
                    if (txtrch_clave.Text.Trim() != "")
                    {
                        foreach (DataRow row in ranchos.Select("prov_clave = '" + txtprov_clave.Text + "' and rch_clave = '" + txtrch_clave.Text + "'"))
                            txtrancho.Text = Convert.ToString(row["rch_nombre"].ToString().Trim());

                        /*cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "select rch_nombre from tb_cat_ranchos where prov_clave ='" + txtprov_clave.Text + "' and rch_clave = '" + txtrch_clave.Text + "'";
                        reader1 = cmnd1.ExecuteReader();
                        while(reader1.Read())
                        {
                            txtrancho.Text = reader1.GetValue(0).ToString();
                        }
                        reader1.Dispose();*/
                    }

                    if (txttbl_clave.Text.Trim() != "")
                    {
                        foreach (DataRow row in tablas.Select("prov_clave = '" + txtprov_clave.Text + "' and rch_clave = '" + txtrch_clave.Text + "' and tbl_clave= '" + txttbl_clave.Text + "'"))
                            txttabla.Text = Convert.ToString(row["tbl_nombre"].ToString());

                        /*cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "select tbl_nombre from tb_cat_tablas where prov_clave ='" + txtprov_clave.Text + "' and rch_clave ='" + txtrch_clave.Text + "' and tbl_clave ='" + txttbl_clave.Text + "'";
                        reader1 = cmnd1.ExecuteReader();
                        while(reader1.Read())
                        {
                            txttabla.Text = reader1.GetValue(0).ToString();
                        }
                        reader1.Dispose();*/
                    }

                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select A.prod_clave, B.prod_nombre, A.rptd_ticket, A.rptd_peso_bruto, A.rptd_tara, A.env_clave, C.env_nombre, A.rptd_tarimas, " +
                                        "A.rptd_cantidad, A.rptd_fechacad, C.env_peso, B.prod_paisorigen, b.prod_clave_gabinc from tb_det_recepcion_pt A,tb_cat_producto B, tb_cat_envases C " +
                                        "where A.rpt_recibo = '" + txtrecibo.Text + "' and B.prod_clave = A.prod_clave and C.env_clave = A.env_clave and (A.rptd_estatus <> 'C' or A.RPTD_ESTATUS IS NULL)";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        decimal p_b = Convert.ToDecimal(reader1.GetValue(3).ToString());
                        decimal t = Convert.ToDecimal(reader1.GetValue(4).ToString());
                        decimal can = Convert.ToDecimal(reader1.GetValue(8).ToString());
                        decimal tar = Convert.ToDecimal(reader1.GetValue(7).ToString());
                        peso_env = Convert.ToDecimal(reader1.GetValue(10).ToString());
                        //if (tar != 0)
                        p_t = Convert.ToDecimal(reader1.GetValue(3).ToString()) - Convert.ToDecimal(reader1.GetValue(4).ToString()) - (Convert.ToDecimal(reader1.GetValue(8).ToString()) * peso_env) - (Convert.ToDecimal(reader1.GetValue(7).ToString()) * 20);
                        //else
                        //    p_t = (20 * tar) + (peso_env * can) + (can * p_u);

                        if (Convert.ToInt32(reader1.GetValue(8).ToString()) > 0)
                            p_u = p_t / Convert.ToInt32(reader1.GetValue(8).ToString());
                        else
                            p_u = 0;

                        DGV1.Rows.Add(reader1.GetValue(0).ToString(), reader1.GetValue(1).ToString(), reader1.GetValue(2).ToString(), p_b.ToString("###,###,##0.00"), t.ToString("###,###,##0.00"), reader1.GetValue(6).ToString(), can.ToString("###,###,###"), tar.ToString("###,###,###"), p_u.ToString("###,###,##0.00"), p_t.ToString("###,###,##0.00"), reader1.GetValue(9).ToString(), "", reader1.GetValue(5).ToString(), "", peso_env, "", reader1.GetValue(11).ToString().Trim(), reader1["prod_clave_gabinc"].ToString());


                    }
                    reader1.Dispose();

                    //trae daños
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select A.dno_nombre, B.dnm_cantidad, A.dno_clave, dno_Calidad from tb_cat_danos A, tb_mstr_danos B where A.dno_tipo = 'PT' and b.rmp_recibo = '" + txtrecibo.Text + "' " +
                                        "and A.dno_clave = B.dno_clave and A.lin_clave = '" + txtlin_clave.Text + "' ORDER BY B.dnm_cantidad desc"; // A.dno_clave";
                    reader1 = cmnd1.ExecuteReader();
                    decimal TotPorce = 0;
                    while (reader1.Read())
                    {
                        decimal TotMtra = (TxtPieza.Text.Trim().Length > 0) ? decimal.Parse(TxtPieza.Text) : 0;
                        decimal Porce = 0;
                        if (TotMtra > 0)
                        {

                            Porce = (Convert.ToDecimal(reader1.GetValue(1).ToString()) * 100) / TotMtra;
                            if (reader1.GetValue(3).ToString().Contains("GRADO 1"))
                            {
                                TotPorce += Porce;
                            }
                            LblGra1.Text = (Convert.ToDecimal(100) - TotPorce).ToString("###.#0");
                        }
                        DGV2.Rows.Add(reader1.GetValue(0).ToString(), reader1.GetValue(1).ToString(), reader1.GetValue(2).ToString(), Porce.ToString("##0.#0"), reader1.GetValue(3));
                    }
                    reader1.Dispose();

                    //trae procesos
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select  B.PROC_NOMBRE, a.prom_cantidad, a.proc_clave from tb_mstr_procesos a, tb_cat_procesos b " +
                                        "where a.rmp_recibo ='" + txtrecibo.Text + "' AND a.prom_Tipo ='PT' and " +
                                        "b.proc_clave = a.proc_clave and a.lin_clave ='" + txtlin_clave.Text + "' and b.lin_clave = a.lin_clave order by a.prom_cantidad desc "; //a.proc_clave";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        decimal TotMtra = (TxtPieza.Text.Trim().Length > 0) ? decimal.Parse(TxtPieza.Text) : 0;
                        decimal Porce = 0;
                        if (TotMtra > 0)
                        {
                            Porce = (Convert.ToDecimal(reader1.GetValue(1).ToString()) * 100) / TotMtra;
                            TotPorce += Porce;
                            LblGra1.Text = (Convert.ToDecimal(100) - TotPorce).ToString("###.#0");
                        }
                        DGV3.Rows.Add(reader1.GetValue(0).ToString(), reader1.GetValue(1).ToString(), reader1.GetValue(2).ToString(), Porce.ToString("##0.#0"));
                    }
                    reader1.Dispose();

                    //trae variedad
                    foreach (DataRow row in variedades.Select("vari_clave= '" + txtvariedad.Text + "' and lin_clave = '" + txtlin_clave.Text + "'"))
                        cbvariedad.Items.Add(Convert.ToString(row["vari_nombre"].ToString().Trim()));

                    /*cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select vari_nombre from tb_cat_variedad where lin_clave ='" + txtlin_clave.Text + "' and vari_clave ='" + txtvariedad.Text + "'";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        cbvariedad.Items.Add(reader1.GetValue(0).ToString().Trim());
                    }*/
                    if (cbvariedad.Items.Count > 0)
                        cbvariedad.SelectedIndex = 0;

                    //reader1.Dispose();
                    thisConnection.Close();
                    btnimpre.Visible = true;
                    btnemail.Visible = true;
                    btncancelrecibo.Enabled = true;

                    DGV1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    //DGV2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    //DGV3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    num_viaje = Convert.ToInt32(txtnumviaje.Text);
                    proveedor = txtprov.Text;
                    rancho_nom = txttabla.Text;
                    hora = txthora.Text;
                    btnpdf.Visible = true;
                    btnaddfile.Visible = true;
                    btnmodi.Enabled = true;
                    btnmodeva.Enabled = true;
                    //btnpesoxtar.Visible = true;
                    if (txtflete.Text.Trim() == "DEV")
                    {
                        DGV2.Visible = false;
                        DGV3.Visible = false;
                        DGV4.Visible = true;
                        string tempo = "";
                        thisConnection.Open();
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "select rpt_recibo_anterior, cajas, tarima, prod_clave from tb_det_devoluciones where rpt_recibo_nuevo = '" + txtrecibo.Text + "'";
                        reader1 = cmnd1.ExecuteReader();
                        while (reader1.Read())
                        {
                            foreach (DataRow row in productos.Select("prod_clave = '" + reader1.GetValue(3).ToString().Trim() + "'"))
                                tempo = Convert.ToString(row["prod_nombre"].ToString().Trim());

                            DGV4.Rows.Add(reader1.GetValue(0).ToString().Trim(), reader1.GetValue(3).ToString().Trim(), tempo, reader1.GetValue(2).ToString().Trim(), reader1.GetValue(1).ToString().Trim());
                        }
                        reader1.Dispose();
                        thisConnection.Close();
                    }
                    if (estatus == "F")
                    {
                        btncancelrecibo.Enabled = false;
                        btnmodi.Enabled = false;
                        btnmodeva.Enabled = false;
                        //                        btnimpre.Visible = false;
                        MessageBox.Show("El recibo esta cancelado", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    if (Environment.MachineName == "OPERATOR-PC")
                        btncancelrecibo.Enabled = false;
                }
            }
        }

        private void btnimpre_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("¿Desea imprimir el recibo?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                PrintDialog printDialog1 = new PrintDialog();
                printDialog1.Document = printDocument1;
                DialogResult result = printDialog1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    PrintDocument printdocu = new PrintDocument();
                    printdocu.PrintPage += new PrintPageEventHandler(this.printDocument1_PrintPage);
                    printdocu.Print();
                    //printDocument1.Print();
                }
                //PrintDocument printdocu = new PrintDocument();
                //printdocu.PrintPage += new PrintPageEventHandler(this.printDocument1_PrintPage);
                //printdocu.Print();
                //PrintPreviewDialog VistaPrevia = new PrintPreviewDialog();
                //VistaPrevia.Document = printDocument1;
                //VistaPrevia.Show();
                MessageBox.Show("Se termino de mandar la impresión", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (MessageBox.Show("¿Desea imprimir las etiquetas verdes?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                etiqueta_verde etiver = new etiqueta_verde();
                num_viaje = Convert.ToInt32(txtnumviaje.Text);
                cletiqueta = lbcletiqueta.Text;
                FechaEla = lbfecha.Value;
                etiqueta_verde.linea = txtlin_clave.Text;
                etiqueta_verde.date = lbfecha.Value.ToString("dd/MM/yyyy");
                if (etiver.dt1.Columns.Count == 0 && etiver.dtoriginal.Columns.Count == 0)
                {
                    etiver.dt1.Columns.Add("prod_clave");
                    etiver.dt1.Columns.Add("prod_nombre");
                    etiver.dt1.Columns.Add("numero_tarimas");
                    etiver.dt1.Columns.Add("cajas_x_tarima");
                    etiver.dt1.Columns.Add("feccad");
                    etiver.dt1.Columns.Add("pais");
                    etiver.dt1.Columns.Add("items");

                    etiver.dtoriginal.Columns.Add("prod_clave");
                    etiver.dtoriginal.Columns.Add("prod_nombre");
                    etiver.dtoriginal.Columns.Add("numero_tarimas");
                    etiver.dtoriginal.Columns.Add("cajas_x_tarima");
                    etiver.dtoriginal.Columns.Add("fec_cad");
                    etiver.dtoriginal.Columns.Add("pais");
                    etiver.dtoriginal.Columns.Add("items");
                }

                for (int i = 0; i < DGV1.Rows.Count; i++)
                {
                    //thisConnection.Open();               
                    DataRow row = etiver.dt1.NewRow();
                    row[0] = DGV1.Rows[i].Cells[0].Value;
                    row[1] = DGV1.Rows[i].Cells[1].Value;
                    row[2] = DGV1.Rows[i].Cells[7].Value;
                    row[3] = Math.Truncate(Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value) / Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value));
                    row[4] = Convert.ToString(DGV1.Rows[i].Cells[10].Value);
                    row[5] = Convert.ToString(DGV1.Rows[i].Cells["prod_pais"].Value);
                    row[6] = Convert.ToString(DGV1.Rows[i].Cells["items"].Value);
                    etiver.dt1.Rows.Add(row);

                    DataRow rw = etiver.dtoriginal.NewRow();
                    rw[0] = DGV1.Rows[i].Cells[0].Value;
                    rw[1] = DGV1.Rows[i].Cells[1].Value;
                    rw[2] = DGV1.Rows[i].Cells[7].Value;
                    rw[3] = DGV1.Rows[i].Cells[6].Value;
                    rw[4] = Convert.ToString(DGV1.Rows[i].Cells[10].Value);
                    rw[5] = Convert.ToString(DGV1.Rows[i].Cells["prod_pais"].Value);
                    rw[6] = Convert.ToString(DGV1.Rows[i].Cells["items"].Value);
                    etiver.dtoriginal.Rows.Add(rw);

                }
                etiver.ShowDialog();
            }

            limpiarTextBoxes(this);
        }

        private void btncancelrecibo_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("¿Desea cancelar el recibo?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                //verifica que no se haya tomado producto de ese recibo de materia prima
                //FOX
                //if (txtrecibo.Text.Trim().Length < 6)
                //    txtrecibo.Text = "0" + txtrecibo.Text;

                //MyConnection.Open();
                //cmd1 = MyConnection.CreateCommand();
                //cmd1.CommandText = "select hrp_num_unidades, hrp_reman_unidades from tb_hist_recepcion where hrp_tipo_recepcion ='PTC' and hrp_recibo = '" + txtrecibo.Text + "'";
                //read1 = cmd1.ExecuteReader();
                //while (read1.Read())
                //{
                //    hrp_num_unidades = Convert.ToDecimal(read1.GetValue(0).ToString());
                //    hrp_reman_unidades = Convert.ToDecimal(read1.GetValue(1).ToString());
                //}
                //read1.Dispose();
                //MyConnection.Close();

                //SQL
                LblObsCan.Visible = true;
                TxtObsCan.Visible = true;
                TxtObsCan.Enabled = true;
                if (TxtObsCan.Text.Trim().Length < 10)
                {
                    MessageBox.Show("Debe Capturar las Observaciones de la Cancelación una vez Capturada vuelva a presionar el Boton Cancelar Recibo", "Cancelar Recibo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    TxtObsCan.Focus();
                    return;
                }
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "select prod_clave, sum(surtido) surtido, prod_nombre from tb_det_trazabilidad where recibo = '" + txtrecibo.Text + "' group by prod_clave, prod_nombre";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    if (Convert.ToDecimal(reader1.GetValue(1).ToString().Trim()) != 0)
                    {
                        MessageBox.Show("No se puede cancelar el recibo, porque del producto: " + reader1.GetValue(2).ToString().Trim() + " ya surtieron cajas", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        thisConnection.Close();
                        return;
                    }
                }
                //cmnd1.CommandText = "select hrp_num_unidades, hrp_reman_unidades from tb_hist_recepcion where hrp_tipo_recepcion ='PTC' and hrp_recibo = '" + txtrecibo.Text + "'";
                //reader1=cmnd1.ExecuteReader();
                //while(reader1.Read())
                //{
                //    hrp_num_unidades = Convert.ToDecimal(reader1.GetValue(0).ToString());
                //    hrp_reman_unidades =Convert.ToDecimal(reader1.GetValue(1).ToString());
                //}
                //reader1.Dispose();

                //if (hrp_num_unidades != hrp_reman_unidades)
                //{
                //    MessageBox.Show("El recibo " + txtrecibo.Text + " no se puede cancelar porque producción ya tomo producto", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //    thisConnection.Close();
                //    return;
                //}
                string body = "", mCorreo = "";
                #region SQL
                try
                {
                    //thisConnection.Open();
                    //cambia el estatus a F en tb_mstr_recepcion_mp
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "Select EMAIL_DEST from TB_MSTR_EMAIL WHERE CNTE_CLAVE = 'CANRECPT' AND EMAIL_MOV = 'RPT'";
                    mCorreo = Convert.ToString(cmnd1.ExecuteScalar());

                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "update tb_mstr_recepcion_pt set rpt_estatus = 'F', rpt_ObsCancelacion = '" + TxtObsCan.Text.Trim() + "' where rpt_recibo = '" + txtrecibo.Text + "'";
                    reader1 = cmnd1.ExecuteReader();
                    reader1.Dispose();

                    //cambia el estatus a C en tb_hist_recepcion
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "update tb_hist_recepcion set hrp_estatus = 'C' where hrp_recibo = '" + txtrecibo.Text + "' and hrp_tipo_recepcion ='PTC'";
                    reader1 = cmnd1.ExecuteReader();
                    reader1.Dispose();

                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select hrp_num_unidades, hrp_peso_neto, lin_clave, prod_clave from tb_hist_recepcion where hrp_recibo ='" + txtrecibo.Text + "' and hrp_tipo_recepcion = 'PTC'";
                    body = "";

                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        hrp_num_unidades = Convert.ToDecimal(reader1.GetValue(0).ToString());
                        hrp_peso_neto = Convert.ToDecimal(reader1.GetValue(1).ToString());
                        linea = reader1.GetValue(2).ToString();
                        producto = reader1.GetValue(3).ToString();

                        //afecta inventarios en la tb_mstr_inventario_pt
                        cmnd11 = thisConnection.CreateCommand();
                        cmnd11.CommandText = "select isnull (invpt_salidas_kg, 0), isnull (invpt_salidas_un, 0) from tb_mstr_inventario_pt where invpt_fecha = '" + System.DateTime.Now.ToShortDateString() + "' and lin_clave = '" + linea + "' and prod_clave = '" + producto + "'";
                        reader11 = cmnd11.ExecuteReader();
                        while (reader11.Read())
                        {
                            invpt_salidas_kg = reader11.GetSqlDecimal(0).Value;
                            invpt_salidas_un = reader11.GetSqlDecimal(1).Value;

                            invpt_salidas_kg = invpt_salidas_kg + hrp_peso_neto;
                            invpt_salidas_un = invpt_salidas_un + hrp_num_unidades;

                            cmnd111 = thisConnection.CreateCommand();
                            cmnd111.CommandText = "update tb_mstr_inventario_pt set invpt_salidas_kg = " + invpt_salidas_kg + ", invpt_salidas_un = " + invpt_salidas_un + " where lin_clave = '" + linea + "' and prod_clave = '" + producto + "'";
                            reader111 = cmnd111.ExecuteReader();
                            reader111.Dispose();
                        }
                        if (reader11.HasRows == false)
                        {
                            cmnd111 = thisConnection.CreateCommand();
                            cmnd111.CommandText = "select top 1 isnull (invpt_inicial_kg,0), isnull (invpt_inicial_un,0), isnull (invpt_entradas_kg,0), isnull (invpt_entradas_un,0), isnull (invpt_salidas_kg, 0), isnull (invpt_salidas_un, 0) from tb_mstr_inventario_pt where lin_clave ='" + linea + "' and prod_clave ='" + producto + "' order BY invpt_fecha desc";
                            reader111 = cmnd111.ExecuteReader();
                            while (reader111.Read())
                            {
                                invpt_inicial_kg = reader111.GetSqlDecimal(0).Value;
                                invpt_inicial_un = reader111.GetSqlDecimal(1).Value;
                                invpt_entradas_kg = reader111.GetSqlDecimal(2).Value;
                                invpt_entradas_un = reader111.GetSqlDecimal(3).Value;
                                invpt_salidas_kg = reader111.GetSqlDecimal(4).Value;
                                invpt_salidas_un = reader111.GetSqlDecimal(5).Value;

                                var_dec_inv_inicial_kg = invpt_inicial_kg + invpt_entradas_kg - invpt_salidas_kg;
                                var_dec_inv_inicial_un = invpt_inicial_un + invpt_entradas_un - invpt_salidas_un;

                                //inserto un nuevo registro
                                cmnd1111 = thisConnection.CreateCommand();
                                cmnd1111.CommandText = "insert into tb_mstr_inventario_pt (invpt_fecha, lin_clave, prod_clave, invpt_inicial_un, invpt_inicial_kg, invpt_entradas_un, invpt_entradas_kg, invpt_salidas_un, invpt_salidas_kg) " +
                                                   "values ('" + System.DateTime.Now.ToShortDateString() + "', '" + linea + "', '" + producto + "', " + var_dec_inv_inicial_un + ", " + var_dec_inv_inicial_kg + ", " + (invpt_salidas_un + hrp_num_unidades) + ", " + (invpt_salidas_kg + hrp_peso_neto) + ", 0, 0)";
                                reader1111 = cmnd1111.ExecuteReader();
                                reader1111.Dispose();
                            }
                            if (reader111.HasRows == false)
                            {
                                cmnd1111 = thisConnection.CreateCommand();
                                cmnd1111.CommandText = "insert into tb_mstr_inventario_pt (invpt_fecha, lin_clave, prod_clave, invpt_inicial_kg, invpt_inicial_un, invpt_entradas_kg, invpt_entradas_un, invpt_salidas_kg, invpt_salidas_un) " +
                                                    "values ('" + System.DateTime.Now.ToShortDateString() + "', '" + linea.Trim() + "', '" + producto + "', 0, 0, " + (invpt_salidas_kg + hrp_peso_neto) + ", " + (invpt_salidas_un + hrp_num_unidades) + ", 0, 0)";
                                reader1111 = cmnd1111.ExecuteReader();
                                reader1111.Dispose();
                            }
                            reader111.Dispose();
                        }
                        reader11.Dispose();
                    }
                    reader1.Dispose();

                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "update tb_det_trazabilidad set pti_estatus_sur = 'S' where recibo = '" + txtrecibo.Text + "' and tipo = 'PTC'";
                    reader1 = cmnd1.ExecuteReader();
                    reader1.Dispose();
                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "Select EMAIL_DEST from TB_MSTR_EMAIL WHERE CNTE_CLAVE = 'CANRECPT' AND EMAIL_MOV = 'RPT'";
                    //mCorreo = Convert.ToString(cmd1.ExecuteScalar());
                    thisConnection.Close();
                }
                catch (SqlException ex)
                {
                    thisConnection.Close();
                    //Utilerias.Class1.registrar_movimiento(DateTime.Now, Utilerias.Class1.Nombre_equipo, Utilerias.Class1.Usuario, "errorSQL", Utilerias.Class1.Formulario, txtrecibo.Text, ex.ToString());
                    Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.3", ex.ToString().Trim(), "MPFOX");
                    Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "Sistem@s2026$", ex.ToString().Trim());
                    MessageBox.Show(ex.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                catch (Exception ex1)
                {
                    thisConnection.Close();
                    Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "Sistem@s2026$", ex1.ToString().Trim());
                    MessageBox.Show(ex1.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                #endregion

                #region FOX
                /*try
                {
                    MyConnection.Open();
                    OleDbCommand dbCmdNull1 = MyConnection.CreateCommand();
                    dbCmdNull1.CommandText = "SET NULL OFF";
                    dbCmdNull1.ExecuteNonQuery();


                    if (txtrecibo.Text.Length < 6)
                        txtrecibo.Text = "0" + txtrecibo.Text;

                    //cambia el estatus a f en tb_mstr_recepcion_pt
                    cmd1 = MyConnection.CreateCommand();
                    cmd1.CommandText = "update tb_mstr_recepcion_pt set rpt_estatus = 'F' where rpt_recibo = '" + txtrecibo.Text + "'";
                    cmd1.Connection = MyConnection;
                    cmd1.ExecuteNonQuery();
                    cmd1.Dispose();

                    //cambia el estatus a C en tb_hist_recepcion
                    cmd1 = MyConnection.CreateCommand();
                    cmd1.CommandText = "update tb_hist_recepcion set hrp_estatus = 'C' where hrp_recibo = '" + txtrecibo.Text + "' and hrp_tipo_recepcion = 'PTC'";
                    cmd1.Connection = MyConnection;
                    cmd1.ExecuteNonQuery();
                    cmd1.Dispose();

                    cmd1 = MyConnection.CreateCommand();
                    cmd1.CommandText = "select hrp_num_unidades, hrp_peso_neto, lin_clave, prod_clave from tb_hist_recepcion where hrp_recibo ='" + txtrecibo.Text + "' and hrp_tipo_recepcion ='PTC'";
                    read1 = cmd1.ExecuteReader();
                    while (read1.Read())
                    {
                        hrp_num_unidades = Convert.ToDecimal(read1.GetValue(0).ToString());
                        hrp_peso_neto = Convert.ToDecimal(read1.GetValue(1).ToString());
                        linea = read1.GetValue(2).ToString();
                        producto = read1.GetValue(3).ToString();

                        //read1.Dispose();

                        OleDbCommand cmd11 = MyConnection.CreateCommand();
                        cmd11 = MyConnection.CreateCommand();
                        cmd11.CommandText = "select invpt_salidas_kg, invpt_salidas_un from tb_mstr_inventario_pt where invpt_fecha = ? and lin_clave = ? and prod_clave = ?";
                        cmd11.Parameters.Add("@Nombre1", OleDbType.Date).Value = System.DateTime.Now.ToShortDateString();
                        cmd11.Parameters.Add("@Nombre2", OleDbType.Char).Value = linea;
                        cmd11.Parameters.Add("@Nombre3", OleDbType.Char).Value = producto;
                        cmd11.Connection = MyConnection;
                        read11 = cmd11.ExecuteReader();
                        while (read11.Read())
                        {
                            invpt_salidas_kg = Convert.ToDecimal(read11.GetValue(0).ToString());
                            invpt_salidas_un = Convert.ToDecimal(read11.GetValue(1).ToString());

                            invpt_salidas_kg = invpt_salidas_kg + hrp_peso_neto;
                            invpt_salidas_un = invpt_salidas_un + hrp_num_unidades;

                            OleDbCommand cmd111 = MyConnection.CreateCommand();
                            cmd111.CommandText = "update tb_mstr_inventario_pt set invpt_salidas_kg = " + invpt_salidas_kg + ", invpt_salidas_un = " + invpt_salidas_un + " where lin_clave = '" + linea + "' and prod_clave = '" + producto + "'";
                            read111 = cmd111.ExecuteReader();
                            read111.Dispose();
                        }
                        if (read11.HasRows == false)
                        {
                            OleDbCommand cmd111 = MyConnection.CreateCommand();
                            cmd111.CommandText = "select top 1 invpt_inicial_kg, invpt_inicial_un, invpt_entradas_kg, invpt_entradas_un, invpt_salidas_kg, invpt_salidas_un from tb_mstr_inventario_pt where lin_clave ='" + linea + "' and prod_clave ='" + producto + "' order BY invpt_fecha desc";
                            read111 = cmd111.ExecuteReader();
                            while (read111.Read())
                            {
                                invpt_inicial_kg = Convert.ToDecimal(read111.GetValue(0).ToString());
                                invpt_inicial_un = Convert.ToDecimal(read111.GetValue(1).ToString());
                                invpt_entradas_kg = Convert.ToDecimal(read111.GetValue(2).ToString());
                                invpt_entradas_un = Convert.ToDecimal(read111.GetValue(3).ToString());
                                invpt_salidas_kg = Convert.ToDecimal(read111.GetValue(4).ToString());
                                invpt_salidas_un = Convert.ToDecimal(read111.GetValue(5).ToString());
                                //read1.Dispose();

                                var_dec_inv_inicial_kg = invpt_inicial_kg + invpt_entradas_kg - invpt_salidas_kg;
                                var_dec_inv_inicial_un = invpt_inicial_un + invpt_entradas_un - invpt_salidas_un;

                                OleDbCommand cmd1111 = MyConnection.CreateCommand();
                                cmd1111.CommandText = "insert into tb_mstr_inventario_pt (invpt_fecha, lin_clave, prod_clave, invpt_inicial_un, invpt_inicial_kg, invpt_entradas_un, invpt_entradas_kg, invpt_salidas_un, invpt_salidas_kg) " +
                                                      "values (?, ?, ?, ?, ?, ?, ?, ?, ?)";
                                cmd1111.Parameters.Add("@Nombre1", OleDbType.Date).Value = System.DateTime.Now.ToShortDateString();
                                cmd1111.Parameters.Add("@Nombre2", OleDbType.Char).Value = linea;
                                cmd1111.Parameters.Add("@Nombre3", OleDbType.Char).Value = producto;
                                cmd1111.Parameters.Add("@Nombre4", OleDbType.Numeric).Value = float.Parse(var_dec_inv_inicial_un.ToString());
                                cmd1111.Parameters.Add("@Nombre5", OleDbType.Numeric).Value = float.Parse(var_dec_inv_inicial_kg.ToString());
                                cmd1111.Parameters.Add("@Nombre6", OleDbType.Numeric).Value = float.Parse((invpt_salidas_un + hrp_num_unidades).ToString());
                                cmd1111.Parameters.Add("@Nombre7", OleDbType.Numeric).Value = float.Parse((invpt_salidas_kg + hrp_peso_neto).ToString());
                                cmd1111.Parameters.Add("@Nombre8", OleDbType.Numeric).Value = 0;
                                cmd1111.Parameters.Add("@Nombre9", OleDbType.Numeric).Value = 0;
                                read1111 = cmd1111.ExecuteReader();
                                read1111.Dispose();
                            }
                            if (read111.HasRows == false)
                            {
                                //read1.Dispose();
                                OleDbCommand cmd11111 = MyConnection.CreateCommand();
                                cmd11111.CommandText = "insert into tb_mstr_inventario_pt (invpt_fecha, lin_clave, prod_clave, invpt_inicial_kg, invpt_inicial_un, invpt_entradas_kg, invpt_entradas_un, invpt_salidas_kg, invpt_salidas_un) " +
                                                       "values (?, ?, ?, ?, ?, ?, ?, ?,? ,?)";
                                cmd11111.Parameters.Add("@Nombre1", OleDbType.Date).Value = System.DateTime.Now.ToShortDateString();
                                cmd11111.Parameters.Add("@Nombre2", OleDbType.Char).Value = linea;
                                cmd11111.Parameters.Add("@Nombre3", OleDbType.Char).Value = producto;
                                cmd11111.Parameters.Add("@Nombre4", OleDbType.Numeric).Value = 0;
                                cmd11111.Parameters.Add("@Nombre5", OleDbType.Numeric).Value = 0;
                                cmd11111.Parameters.Add("@Nombre6", OleDbType.Numeric).Value = float.Parse((invpt_salidas_un + hrp_num_unidades).ToString());
                                cmd11111.Parameters.Add("@Nombre7", OleDbType.Numeric).Value = float.Parse((invpt_salidas_kg + hrp_peso_neto).ToString());
                                cmd11111.Parameters.Add("@Nombre8", OleDbType.Numeric).Value = 0;
                                cmd11111.Parameters.Add("@Nombre9", OleDbType.Numeric).Value = 0;
                                read1111 = cmd11111.ExecuteReader();
                                read1111.Dispose();
                            }
                            read111.Dispose();
                        }
                        read11.Dispose();
                    }
                    read1.Dispose();
                    MyConnection.Close();
                }
                catch (OleDbException ex)
                {
                    MyConnection.Close();
                    Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.3", ex.ToString().Trim(), "MPFOX");
                    Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString().Trim());
                    //Utilerias.Class1.registrar_movimiento(DateTime.Now, Utilerias.Class1.Nombre_equipo, Utilerias.Class1.Usuario, "errorSQL", Utilerias.Class1.Formulario, txtrecibo.Text, ex.ToString());
                    MessageBox.Show(ex.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                catch (Exception ex1)
                {
                    MyConnection.Close();
                    Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString().Trim());
                    MessageBox.Show(ex1.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }*/
                #endregion
                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "B", "2.3", txtrecibo.Text, "CANCELACION RECIBO DE PT " + txtrecibo.Text, "SIPGAB");
                MessageBox.Show("El recibo : " + txtrecibo.Text + " ha sido cancelado con éxito", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                body = "Hola Buen día" + "<br><br>" + "<H3><p style='color: blue'>CANCELACION DEL RECIBO " + txtrecibo.Text + "  Flete: " + txtflete.Text.Trim() + "</p></H3><br><br>" +
                       "<table border=\"1\"> " +
                       " <tr>   " +
                       " <th scope=\"col\">PRODUCTO</strong></th> " +
                       " <th scope=\"col\">CAJAS </strong></th> " +
                       " <th scope=\"col\">TARIMAS</strong></th> " +
                       " <th scope=\"col\">PESO UNITARIO</strong></th> " +
                       " </tr>";
                for (int i = 0; i < DGV1.Rows.Count; i++)
                {
                    body += "<tr>" +
                           "<td ALING=LEFT>" + DGV1.Rows[i].Cells["Nomb"].Value.ToString().Trim() + "</td>" +
                           "<td ALING=CENTER>" + Convert.ToDecimal(DGV1.Rows[i].Cells["Canti"].Value).ToString("#,##0") + "</td>" +
                           "<td ALING=CENTER>" + Convert.ToDecimal(DGV1.Rows[i].Cells["Tarim"].Value).ToString("#,##0") + "</td>" +
                           "<td ALING=CENTER>" + Convert.ToDecimal(DGV1.Rows[i].Cells["pes_uni"].Value).ToString("#,##0.00") + " Kg</td>" +
                           "</tr>";
                }
                body += "</table><br><br>Obs Cancelación: " + TxtObsCan.Text.Trim() + "<br>Computadora: " + System.Environment.MachineName.ToString();
                if (mCorreo.Trim().Length > 0)
                    SendMail(mCorreo, "", body, "Cancelación Del Recibo " + txtrecibo.Text + "  Del día: " + lbfecha.Value.ToShortDateString());
                //Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString().Trim());
                limpiarTextBoxes(this);
                btncancelrecibo.Enabled = false;
                btnimpre.Visible = false;
            }
            else
            { return; }
        }

        private void cbvariedad_SelectionChangeCommitted(object sender, EventArgs e)
        {
            foreach (DataRow row in variedades.Select("lin_clave= '" + txtlin_clave.Text + "' and vari_nombre = '" + cbvariedad.SelectedItem.ToString() + "'"))
            {
                txtvariedad.Text = Convert.ToString(row["vari_clave"].ToString().Trim());
                varieda = Convert.ToString(row["vari_clave"].ToString().Trim());
            }
            /*thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "select vari_clave from tb_cat_variedad where lin_clave ='" + txtlin_clave.Text + "' and vari_nombre = '" + cbvariedad.SelectedItem.ToString() + "'";
            reader1 = cmnd1.ExecuteReader();
            while (reader1.Read())
            {
                txtvariedad.Text = reader1.GetValue(0).ToString().Trim();
                varieda = txtvariedad.Text.Trim();
            }
            reader1.Dispose();
            thisConnection.Close();*/
        }

        private void txtvariedad_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtvariedad.Text.Trim().Length > 0)
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    foreach (DataRow row in variedades.Select("lin_clave= '" + txtlin_clave.Text + "' and vari_clave = '" + txtvariedad.Text + "'"))
                        cbvariedad.SelectedItem = Convert.ToString(row["vari_nombre"].ToString().Trim());


                    if (cbvariedad.SelectedIndex == -1)
                    {
                        MessageBox.Show("No existe ninguna variedad con esa clave. Favor de checar la clave", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    //thisConnection.Open();
                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "select vari_nombre from tb_cat_variedad where lin_clave ='" + txtlin_clave.Text + "' and vari_clave ='" + txtvariedad.Text + "'";
                    //reader1 = cmnd1.ExecuteReader();
                    //while (reader1.Read())
                    //{
                    //    cbvariedad.SelectedItem = reader1.GetValue(0).ToString().Trim();
                    //}
                    //if (reader1.HasRows == false)
                    //{
                    //    MessageBox.Show("No existe ninguna variedad con esa clave. Favor de checar la clave", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    thisConnection.Close();
                    //    return;
                    //}
                    //reader1.Dispose();
                    //thisConnection.Close();
                }
            }
        }

        private void txtprov_clave_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (txtprov_clave.Text.Length > 0)
                {
                    foreach (DataRow row in proveedores.Select("prov_clave ='" + txtprov_clave.Text + "'"))
                        txtprov.Text = Convert.ToString(row["prov_nombre"].ToString().Trim());

                    if (txtprov.Text.Trim() == "")
                    {
                        MessageBox.Show("No existe ningún proveedor con esa clave, favor de intentar con otra clave", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.None);
                        return;
                    }
                    /*thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select prov_nombre from tb_cat_proveedor where prov_clave = '" + txtprov_clave.Text + "'";
                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            txtprov.Text = reader1.GetValue(0).ToString().Trim();
                        }
                        reader1.Dispose();
                        thisConnection.Close();
                        proveedor = txtprov.Text.Trim();
                    }
                    else
                    {
                        thisConnection.Close();
                        MessageBox.Show("No existe ningún proveedor con esa clave, favor de intentar con otra clave", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.None);
                        return;
                    }*/
                }
            }
        }

        private void txtrch_clave_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (txtrch_clave.Text.Length > 0)
                {
                    foreach (DataRow row in ranchos.Select("prov_clave ='" + txtprov_clave.Text + "' and rch_clave = '" + txtrch_clave.Text + "'"))
                        txtrancho.Text = Convert.ToString(row["rch_nombre"].ToString().Trim());

                    if (txtrancho.Text.Trim() == "")
                    {
                        MessageBox.Show("No existe ningún rancho con esa clave, favor de intentar con otra clave", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.None);
                        return;
                    }
                    /*thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select rch_nombre from tb_cat_ranchos where prov_clave = '" + txtprov_clave.Text + "' and rch_clave = '" + txtrch_clave.Text + "'";
                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            txtrancho.Text = reader1.GetValue(0).ToString().Trim();
                        }
                        reader1.Dispose();
                        thisConnection.Close();
                    }
                    else
                    {
                        thisConnection.Close();
                        MessageBox.Show("No existe ningún rancho con esa clave, favor de intentar con otra clave", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.None);
                        return;
                    }*/
                }
            }
        }

        private void txttbl_clave_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (txttbl_clave.Text.Length > 0)
                {
                    foreach (DataRow row in tablas.Select("prov_clave = '" + txtprov_clave.Text + "' and rch_clave= '" + txtrch_clave.Text + "' and tbl_clave = '" + txttbl_clave.Text + "'"))
                        txttabla.Text = Convert.ToString(row["tbl_nombre"].ToString().Trim());

                    if (txttabla.Text.Trim() != "")
                    {
                        CBcodigo.Items.Clear();
                        foreach (DataRow row in subtablas.Select("prov_clave = '" + txtprov_clave.Text + "' and rch_clave= '" + txtrch_clave.Text + "' and tbl_clave='" + txttbl_clave.Text + "'"))
                            CBcodigo.Items.Add(Convert.ToString(row["tbl_codigo"].ToString().Trim()));
                    }
                    else
                    {
                        MessageBox.Show("No existe ninguna tabla con esa clave, favor de intentar con otra clave", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.None);
                        return;
                    }
                    /*thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "select tbl_nombre from tb_cat_tablas where prov_clave = '" + txtprov_clave.Text + "' and rch_clave = '" + txtrch_clave.Text + "' and tbl_clave = '" + txttbl_clave.Text + "'";
                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            txttabla.Text = reader1.GetValue(0).ToString().Trim();
                        }
                        reader1.Dispose();

                        CBcodigo.Items.Clear();
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "select tbl_codigo from tb_cat_subtablas where prov_clave ='" + txtprov_clave.Text + "' and rch_clave = '" + txtrch_clave.Text + "' and tbl_clave = '" + txttbl_clave.Text + "'";
                        reader1 = cmnd1.ExecuteReader();
                        while (reader1.Read())
                        {
                            CBcodigo.Items.Add(reader1.GetValue(0).ToString());
                        }
                        reader1.Dispose();
                        thisConnection.Close();
                        rancho_nom = txttabla.Text;
                    }
                    else
                    {
                        thisConnection.Close();
                        MessageBox.Show("No existe ninguna tabla con esa clave, favor de intentar con otra clave", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.None);
                        return;
                    }*/
                }
            }
        }

        private void txthora_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back)
            {
                e.Handled = false;
            }
            else if (Char.IsControl(e.KeyChar))
            {
                e.Handled = false;
            }
            if (txthora.Text.Trim().Length == 2)
            {
                string hour = txthora.Text;
                txthora.Text = hour + ":";
                txthora.Select(3, 0);
            }
        }

        private void DGV1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                if (DGV1.CurrentCell.ColumnIndex == 0)
                {
                    int fill = Convert.ToInt32(DGV1.CurrentCell.RowIndex.ToString());
                    foreach (DataRow row in peso_x_tarimas.pesotarima.Select("produc = '" + DGV1.Rows[fill].Cells["produc"].Value + "'"))
                        row.Delete();

                    DGV1.Rows.RemoveAt(DGV1.CurrentCell.RowIndex);
                }
            }
        }

        private void btnminimizar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void btnimpeti_Click(object sender, EventArgs e)
        {
            Imp_eti_tar etitar = new Imp_eti_tar();
            etitar.ShowDialog();
        }

        private void btnemail_Click(object sender, EventArgs e)
        {
            //if (MessageBox.Show("¿Desea enviar el recibo por correo?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            //{
            //if (email.Trim() == "")
            //{
            //    MessageBox.Show("El proveedor no tiene cuenta de correo registrada, favor de agregarla", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}
            //if (txtprov.Text.Contains("AGUILARES"))
            //{
            SentToOutlook();
            //    return;
            //}

            label7.Text = "ENVIANDO CORREO . . .";
            label7.Visible = true;
            label7.Update();

            PrintDocument printdocu = new PrintDocument();
            printdocu.PrintPage += new PrintPageEventHandler(this.Generar_PDF_PrintPage);
            printdocu.PrinterSettings.PrinterName = "Foxit Reader PDF Printer";
            printdocu.Print();
            string body = string.Format("Hola Buen día" + "<br><br>" + "Envio información del detalle del recibo " + txtrecibo.Text + "<br><br>" +
                        "Del Día " + lbfecha.Text + "<br><br>" +
                        "Atentamente, " + "<br><br>" +
                        "Comercializadora GAB S.A. de C.V." + "<br>" +
                        "www.mrlucky.com.mx");
            var inv_pt1 = @"C:\\Reportes\document.pdf";

            if (Utilerias.Class1.ConnectionString.Contains("GABIRASQL"))
            {
                var inv_pt2 = @"\\gabira1\Recibos_proveedores\PT_" + txtrecibo.Text + ".pdf";
                if (!File.Exists(inv_pt1))
                {
                    using (FileStream fs = File.Create(inv_pt1)) { }
                }
                if (File.Exists(inv_pt2))
                    File.Delete(inv_pt2);

                File.Move(inv_pt1, inv_pt2);
                if (email.Trim() != "")
                    SendMail(email, inv_pt2.ToString(), body, "Envio información del recibo " + txtrecibo.Text);
            }
            else
            {
                var inv_pt3 = @"C:\\Reportes\PT_" + txtrecibo.Text + ".pdf";
                if (File.Exists(inv_pt3))
                    File.Delete(inv_pt3);

                File.Move(inv_pt1, inv_pt3);
                if (email.Trim() != "")
                    SendMail(email, inv_pt3.ToString(), body, "Envio información del recibo " + txtrecibo.Text);
            }
            MessageBox.Show("El archivo ha sido enviado al proveedor", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            label7.Visible = false;
            limpiarTextBoxes(this);
            LblGra1.Text = "0";
            //}
        }

        public void SendMail(string Dest, string Archivo, string mBody, string mAsunto)
        {
            MailMessage msg = new MailMessage();
            MailMessage email = new MailMessage();

            string[] destinatarios = Dest.Split(';');
            foreach (string destinos in destinatarios)
            {
                email.To.Add(new MailAddress(destinos));
            }
            //email.To.Add(new MailAddress("gcamacho@mrlucky.com.mx"));

            email.From = new MailAddress("mprima@mrlucky.com.mx"); //
            email.Subject = mAsunto; //"Mensaje de Prueba";
            email.Body = mBody;  //"Información de la factura";
            email.IsBodyHtml = true;
            email.Priority = MailPriority.Normal;

            //string archivo = @"C:\Reportes\factura_informativa.txt";
            if (Archivo.Trim().Length > 0)
                if (File.Exists(Archivo))
                    email.Attachments.Add(new Attachment(Archivo));

            SmtpClient smtp = new SmtpClient();
            smtp.Host = "mail1.mrlucky.com.mx";
            smtp.Port = 587;
            smtp.EnableSsl = true;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential("mprima", "F7udjVm1928M$123");

            try
            {
                smtp.Send(email);
                email.Dispose();
                //MessageBox.Show("correo enviado", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("correo no enviado\r\n" + ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void Generar_PDF_PrintPage(object sender, PrintPageEventArgs e)
        {
            SolidBrush color = new SolidBrush(Color.Black);
            Font fuente = new Font("Courier", 10);
            Font fuente2 = new Font("Courier", 8);
            Font fuente3 = new Font("Courier", 6);

            Image newImage1 = Image.FromFile(@"C:/sisgabweb/PT.jpg");
            e.Graphics.DrawImage(newImage1, 0, 0);

            //recibo
            Point recib = new Point(660, 130);
            e.Graphics.DrawString(txtrecibo.Text, fuente, color, recib);


            //flete
            Point fle = new Point(710, 185);
            e.Graphics.DrawString(txtflete.Text, fuente, color, fle);

            //ticket bascula            
            //Point tic = new Point(670, 210);
            //e.Graphics.DrawString(ticket_bascula.SelectedItem.ToString().Trim(), fuente, color, tic);

            //nombre proveedor
            Point prove = new Point(25, 200);
            e.Graphics.DrawString(txtprov.Text, fuente, color, prove);

            //rancho
            Point rch = new Point(25, 270);
            e.Graphics.DrawString(txtrancho.Text, fuente, color, rch);

            //tabla
            Point tbl = new Point(300, 270);
            e.Graphics.DrawString(txttabla.Text, fuente, color, tbl);

            //fecha
            Point day = new Point(635, 275);
            e.Graphics.DrawString(lbfecha.Text, fuente, color, day);

            int cont = 340;
            coor_x = 0;
            //grid de los productos
            for (int i = 0; i < DGV1.Rows.Count; i++)
            {
                Point produc = new Point(25, cont);
                string pr = "";
                if (Convert.ToString(DGV1.Rows[i].Cells[1].Value).Length > 32)
                    pr = Convert.ToString(DGV1.Rows[i].Cells[1].Value).Substring(0, 31);
                else
                    pr = Convert.ToString(DGV1.Rows[i].Cells[1].Value);

                e.Graphics.DrawString(pr, fuente, color, produc);

                Point tick = new Point(300, cont);
                e.Graphics.DrawString(Convert.ToString(DGV1.Rows[i].Cells[2].Value), fuente, color, tick);

                Point envas = new Point(370, cont);
                if (Convert.ToString(DGV1.Rows[i].Cells[5].Value).Length > 10)
                    e.Graphics.DrawString(Convert.ToString(DGV1.Rows[i].Cells[5].Value).Substring(0, 10), fuente, color, envas);
                else
                    e.Graphics.DrawString(Convert.ToString(DGV1.Rows[i].Cells[5].Value), fuente, color, envas);

                coordenada_x(Convert.ToString(DGV1.Rows[i].Cells[6].Value).Length, "unidades");
                Point cant = new Point(coor_x, cont);
                e.Graphics.DrawString(Convert.ToString(DGV1.Rows[i].Cells[6].Value), fuente, color, cant);

                coordenada_x(Convert.ToString(DGV1.Rows[i].Cells[8].Value).Length, "peso_unitario");
                if (coor_x <= 535)
                    coor_x = 630;
                Point pes_uni = new Point(coor_x, cont);
                e.Graphics.DrawString(Convert.ToDecimal(DGV1.Rows[i].Cells[8].Value).ToString("###,###,###.00"), fuente, color, pes_uni);

                coordenada_x(Convert.ToString(DGV1.Rows[i].Cells[9].Value).Length, "peso_total");
                Point pes_tot = new Point(coor_x, cont);
                e.Graphics.DrawString(Convert.ToDecimal(DGV1.Rows[i].Cells[9].Value).ToString("###,###,###.00"), fuente, color, pes_tot);

                cont = cont + 13;
            }
            // SE IMPRIME LA INFORMACION DEL GRADO 1
            cont = 710;
            Point numpiezas = new Point(25, cont);
            e.Graphics.DrawString("Piezas de la Evaluacion: " + TxtPieza.Text, fuente, color, numpiezas);
            numpiezas = new Point(400, cont);
            e.Graphics.DrawString("% Grado 1: " + LblGra1.Text + " %", fuente, color, numpiezas);

            #region imprime defectos
            if (DGV2.Rows.Count > 1)
            {
                cont = 762;
                DGV2.Sort(DGV2.Columns[1], ListSortDirection.Descending);
                for (int i = 0; i < DGV2.Rows.Count; i++)
                {
                    //if (Convert.ToInt32(DGV2.Rows[i].Cells[1].Value) <= 0)
                    //    continue;
                    Point nom_defecto = new Point(25, cont);
                    e.Graphics.DrawString(Convert.ToString(DGV2.Rows[i].Cells[0].Value), fuente3, color, nom_defecto);

                    Point cantidad_defecto = new Point(250, cont);
                    e.Graphics.DrawString(Convert.ToString(DGV2.Rows[i].Cells[1].Value), fuente3, color, cantidad_defecto);
                    cantidad_defecto = new Point(350, cont);
                    e.Graphics.DrawString(Convert.ToDecimal(DGV2.Rows[i].Cells[3].Value).ToString("##0.#0"), fuente3, color, cantidad_defecto);
                    cont = cont + 10;
                }
            }
            #endregion

            #region imprime procesos
            if (DGV3.Rows.Count > 1)
            {
                //cont = 905;
                cont += 13;
                DGV3.Sort(DGV3.Columns[1], ListSortDirection.Descending);
                for (int i = 0; i < DGV3.Rows.Count; i++)
                {
                    //if (Convert.ToInt32(DGV3.Rows[i].Cells[1].Value) <= 0)
                    //    continue;
                    Point nom_proceso = new Point(25, cont);
                    e.Graphics.DrawString(Convert.ToString(DGV3.Rows[i].Cells[0].Value), fuente3, color, nom_proceso);

                    Point cantidad_proceso = new Point(250, cont);
                    e.Graphics.DrawString(Convert.ToString(DGV3.Rows[i].Cells[1].Value), fuente3, color, cantidad_proceso);
                    cantidad_proceso = new Point(350, cont);
                    e.Graphics.DrawString(Convert.ToDecimal(DGV3.Rows[i].Cells[3].Value).ToString("##0.#0"), fuente3, color, cantidad_proceso);
                    cont = cont + 10;
                }
            }
            #endregion

            //tipo de recepción
            Point tipo_recep = new Point(25, 625);
            string type = cbtipo.SelectedItem.ToString().Trim();
            int lon_typy = (type.Length) - 6;
            type = type.Substring(6, lon_typy);
            e.Graphics.DrawString(type, fuente, color, tipo_recep);

            //observaciones                        
            Point obser = new Point(440, 780);
            e.Graphics.DrawString(txtobs.Text, fuente, color, obser);

            //pesador
            //e.Graphics.RotateTransform(270);            
            Point pesador = new Point(185, 625);
            //Point pesador = new Point(-500, 200);            
            e.Graphics.DrawString(txtpesador.Text, fuente, color, pesador);

            //evaluador 
            Point evaluadro = new Point(430, 625);
            e.Graphics.DrawString(txtevaluador.Text, fuente, color, evaluadro);

            if (cbevaluacion.SelectedIndex == 0)
                e.Graphics.DrawString("X", fuente, color, 620, 625);

            if (cbevaluacion.SelectedIndex == 1)
                e.Graphics.DrawString("X", fuente, color, 620, 655);

            if (cbevaluacion.SelectedIndex == 2)
                e.Graphics.DrawString("X", fuente, color, 620, 685);

            //variedad
            if (cbvariedad.SelectedIndex != -1)
            {
                Point varieda = new Point(440, 800);
                e.Graphics.DrawString("Variedad: " + cbvariedad.SelectedItem.ToString(), fuente, color, varieda);
            }
        }

        private void txtfolio_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtfolio.Text.Trim().Length >= 5)
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (tipo_recepcion.Trim() == "")
                    {
                        MessageBox.Show("Favor de seleccionar de donde proviene el recibo. Campo o Planta", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    txtclaveprod.Text = DGV1.Rows[DGV1.CurrentRow.Index].Cells["produc"].Value.ToString().Trim();
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    if (rbcampo.Checked == true)
                        //cmnd1.CommandText = "select top 1 * from tb_det_trazabilidad where recibo = '" + txtfolio.Text + "' and prod_clave = '" + txtclaveprod.Text + "'";
                        cmnd1.CommandText = "select count(*) from tb_det_trazabilidad where recibo = '" + txtfolio.Text + "' and prod_clave = '" + txtclaveprod.Text + "'";
                    if (rbplanta.Checked == true)
                        //cmnd1.CommandText = "select top 1 * from tb_det_eti_final where folio = '" + txtfolio.Text + "' and cve_prod = '" + txtclaveprod.Text + "'";
                        cmnd1.CommandText = "select count(*) from tb_det_eti_final where folio = '" + txtfolio.Text + "' and cve_prod = '" + txtclaveprod.Text + "'";

                    int hay = Convert.ToInt32(cmnd1.ExecuteScalar());
                    thisConnection.Close();
                    if (hay == 0)
                    {
                        MessageBox.Show("El recibo no se encontro, favor de verificarlo", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    txttar.Focus();
                    txttar.SelectAll();
                }
            }
        }

        private void txttar_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back)
            {
                e.Handled = false;
            }
            else if (Char.IsControl(e.KeyChar))
            {
                e.Handled = false;
            }
            if (e.KeyChar == (char)Keys.Enter)
            {
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                if (rbcampo.Checked == true)
                    //cmnd1.CommandText = "select top 1 * from tb_det_trazabilidad where recibo = '" + txtfolio.Text + "' and prod_clave = '" + txtclaveprod.Text + "' and tarima = '" + txttar.Text + "'";
                    cmnd1.CommandText = "select count(*) from tb_det_trazabilidad where recibo = '" + txtfolio.Text + "' and prod_clave = '" + txtclaveprod.Text + "' and tarima = '" + txttar.Text + "'";
                if (rbplanta.Checked == true)
                    //cmnd1.CommandText = "select top 1 * from tb_det_eti_final where folio = '" + txtfolio.Text + "' and cve_prod = '" + txtclaveprod.Text + "' and tarima = '" + txttar.Text + "'";
                    cmnd1.CommandText = "select count(*) from tb_det_eti_final where folio = '" + txtfolio.Text + "' and cve_prod = '" + txtclaveprod.Text + "' and tarima = '" + txttar.Text + "'";

                int hay = Convert.ToInt32(cmnd1.ExecuteScalar());

                if (hay == 0)
                {
                    MessageBox.Show("No se encontro ninguna tarima, favor de verificarlo", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    thisConnection.Close();
                    return;
                }
                else
                {
                    cmnd1 = thisConnection.CreateCommand();
                    if (rbcampo.Checked == true)
                        //cmnd1.CommandText = "select top 1 * from tb_det_trazabilidad where recibo = '" + txtfolio.Text + "' and prod_clave = '" + txtclaveprod.Text + "' and tarima = '" + txttar.Text + "'";
                        cmnd1.CommandText = "select etiqueta from tb_det_trazabilidad where recibo = '" + txtfolio.Text + "' and prod_clave = '" + txtclaveprod.Text + "' and tarima = '" + txttar.Text + "'";
                    if (rbplanta.Checked == true)
                        //cmnd1.CommandText = "select top 1 * from tb_det_eti_final where folio = '" + txtfolio.Text + "' and cve_prod = '" + txtclaveprod.Text + "' and tarima = '" + txttar.Text + "'";
                        cmnd1.CommandText = "select num_cajas from tb_det_eti_final where folio = '" + txtfolio.Text + "' and cve_prod = '" + txtclaveprod.Text + "' and tarima = '" + txttar.Text + "'";

                    txtnumcjs.Text = Convert.ToString(cmnd1.ExecuteScalar()).Trim();
                }
                thisConnection.Close();
                txtnumcjs.Focus();
                txtnumcjs.SelectAll();
            }
            //if (txttar.Text.Trim().Length == 2)
            //{
            //    string tarim = txttar.Text;
            //    txttar.Text = tarim + "/";
            //    txttar.Select(3, 0);
            //}
            //if (txttar.Text.Trim().Length == 4)
            //    txtnumcjs.Focus();
        }

        //int count = 0;
        private void txtnumcjs_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                if (rbcampo.Checked == true)
                {
                    cmnd1.CommandText = "select etiqueta from tb_det_trazabilidad where recibo = '" + txtfolio.Text + "' and prod_clave = '" + txtclaveprod.Text + "' and tarima = '" + txttar.Text + "'";
                    tipo_recepcion = "PTC";
                }
                if (rbplanta.Checked == true)
                {
                    cmnd1.CommandText = "select num_cajas from tb_det_eti_final where folio = '" + txtfolio.Text + "' and cve_prod = '" + txtclaveprod.Text + "' and tarima = '" + txttar.Text + "'";
                    tipo_recepcion = "PTP";
                }
                int hay = Convert.ToInt32(cmnd1.ExecuteScalar());
                thisConnection.Close();
                if (Convert.ToInt32(txtnumcjs.Text) > hay)
                {
                    MessageBox.Show("El número de cajas es mayor al número de cajas por tarima", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtnumcjs.SelectAll();
                    return;
                }

                bool find = false;
                int con = 0, count = 0;

                if (DGV4.Rows.Count != 0)
                    con = DGV4.Rows.Count;

                string prod_name = "";
                foreach (DataRow rw in productos.Select("prod_clave= '" + txtclaveprod.Text + "'"))
                    prod_name = Convert.ToString(rw["prod_nombre"].ToString().Trim());

                //valida que la tarima no haya sido agregada ya
                if (con == 0)
                    DGV4.Rows.Insert(con, txtfolio.Text.Trim(), txtclaveprod.Text.Trim(), prod_name, txttar.Text, txtnumcjs.Text, tipo_recepcion);
                else
                {
                    for (int i = 0; i < DGV4.Rows.Count; i++)
                    {
                        if (txtfolio.Text.Trim() == Convert.ToString(DGV4.Rows[i].Cells["rec_folio"].Value).Trim())
                        {
                            if (Convert.ToString(DGV4.Rows[i].Cells["pallet"].Value).Trim() == txttar.Text.Trim())
                            {
                                MessageBox.Show("La tarima ya fue agregada", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }
                            else
                            {
                                DGV4.Rows.Insert(con, txtfolio.Text.Trim(), txtclaveprod.Text.Trim(), prod_name, txttar.Text, txtnumcjs.Text, tipo_recepcion);
                                find = true;
                                break;
                            }
                        }
                    }
                    if (find == false)
                        DGV4.Rows.Insert(con, txtfolio.Text.Trim(), txtclaveprod.Text.Trim(), prod_name, txttar.Text, txtnumcjs.Text, tipo_recepcion);
                }

                DGV4.Visible = true;

                if (DGV4.Rows.Count > 0)
                {
                    for (int a = 0; a < DGV4.Rows.Count; a++)
                    {
                        if (a == 0)
                            count = Convert.ToInt32(DGV4.Rows[a].Cells["boxes"].Value);
                        else
                        {
                            //si el producto cambia
                            if (Convert.ToString(DGV4.Rows[a].Cells["cve_prod"].Value).Trim() != Convert.ToString(DGV4.Rows[a - 1].Cells["cve_prod"].Value))
                            {
                                foreach (DataRow row in not_cre.Select("producto = '" + Convert.ToString(DGV4.Rows[a - 1].Cells["cve_prod"].Value) + "'"))
                                {
                                    if (count > Convert.ToInt32(row["cantidad"].ToString()))
                                    {
                                        MessageBox.Show("No se puede agregar más cajas, sobre pasa la cantidad", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        DGV4.Rows.RemoveAt(a - 1);
                                        return;
                                    }
                                    count = 0;
                                }
                            }
                            else
                            {
                                count += Convert.ToInt32(DGV4.Rows[a].Cells["boxes"].Value);
                                foreach (DataRow row in not_cre.Select("producto = '" + Convert.ToString(DGV4.Rows[a - 1].Cells["cve_prod"].Value) + "'"))
                                {
                                    if (count > Convert.ToInt32(row["cantidad"].ToString()))
                                    {
                                        MessageBox.Show("No se puede agregar más cajas, sobre pasa la cantidad", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        DGV4.Rows.RemoveAt(a - 1);
                                        return;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void DGV1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            txtclaveprod.Text = DGV1.Rows[DGV1.CurrentRow.Index].Cells["produc"].Value.ToString().Trim();
        }

        private void DGV4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                if (DGV4.CurrentCell.ColumnIndex == 0)
                {
                    DGV4.Rows.RemoveAt(DGV4.CurrentCell.RowIndex);
                }
            }
        }


        private void txtflete_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtflete.Text.Trim().Contains("DEV"))
            {
                DGV2.Visible = false;
                DGV3.Visible = false;
                label8.Visible = true;
                txtfolio.Visible = true;
                lbguion1.Visible = true;
                txtclaveprod.Visible = true;
                lbguion2.Visible = true;
                txttar.Visible = true;
                lbnumcjs.Visible = true;
                txtnumcjs.Visible = true;
                groupBox1.Visible = true;
                PnlProdDev.Visible = true;
                DGV4.Visible = true;
            }
        }

        private void cbevaluacion_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbevaluacion.SelectedIndex == 0 || cbevaluacion.SelectedIndex == 1)
            {
                situacion = "S";
                CBIngProd_fis.Checked = true;
            }
            if (cbevaluacion.SelectedIndex == 2 || cbevaluacion.SelectedIndex == 1)
            {
                if (cbevaluacion.SelectedIndex == 2)
                {
                    situacion = "N";
                    CBIngProd_fis.Checked = false;
                }
                if (cbevaluacion.SelectedIndex == 1)
                {
                    situacion = "S";
                    CBIngProd_fis.Checked = true;
                }


                No_conformidad.recibo = txtrecibo.Text;
                No_conformidad.nombre_proveedor = txtprov.Text + " - " + txttabla.Text;
                No_conformidad.evaluacion = cbevaluacion.SelectedItem.ToString().Trim();
                No_conformidad.fecha = lbfecha.Text;
                No_conformidad.danos = "";
                for (int i = 0; i < DGV1.Rows.Count; i++)
                {
                    if (Convert.ToDecimal(DGV1.Rows[i].Cells[2].Value) > 0)
                        No_conformidad.danos = No_conformidad.danos + Convert.ToString(DGV1.Rows[i].Cells[0].Value).Trim() + " " + Convert.ToString(DGV1.Rows[i].Cells[2].Value).Trim() + "%, ";
                }
                No_conformidad.danos = No_conformidad.danos.Substring(0, No_conformidad.danos.Length - 2);
                No_conformidad nocon = new No_conformidad();
                No_conformidad.productos.Rows.Clear();

                foreach (DataGridViewRow row in DGV1.Rows)
                {
                    DataRow rw = No_conformidad.productos.NewRow();
                    rw["producto"] = Convert.ToString(row.Cells["Nomb"].Value.ToString().Trim());
                    No_conformidad.productos.Rows.Add(rw);
                }

                nocon.ShowDialog();
                btnCancel_Click(sender, e);
                //No_conformidad.recibo = txtrecibo.Text;
                //No_conformidad.nombre_proveedor = txtprov.Text;
                //No_conformidad.evaluacion = cbevaluacion.SelectedItem.ToString().Trim();
                //No_conformidad nocon = new No_conformidad();                                
                //nocon.ShowDialog();
            }
        }

        private void btnpdf_Click(object sender, EventArgs e)
        {

            string archivo = @"\\gabira1\No_conformidad\" + txtfolnocon.Text + "PT.pdf";
            if (File.Exists(archivo))
            {
                ProcessStartInfo proces = new ProcessStartInfo(archivo);
                Process.Start(proces);
            }
            else
                MessageBox.Show("El archivo no existe: " + archivo, "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnaddfile_Click(object sender, EventArgs e)
        {
            var pdf_servidor = @"\\gabira1\No_conformidad\ " + txtfolnocon.Text + "PT.pdf";
            CargarPDF = new OpenFileDialog();
            CargarPDF.Filter = "Pdf Files|*.pdf";
            CargarPDF.InitialDirectory = "C:\\no_conformidad";
            if (CargarPDF.ShowDialog() == DialogResult.OK)
                ruta_pdf = CargarPDF.FileName;

            if (File.Exists(pdf_servidor))
            {
                if (MessageBox.Show("El archivo ya existe. ¿Desea reemplazarlo?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    File.Copy(ruta_pdf, pdf_servidor, true);
                }
            }
            else
            {
                if (File.Exists(ruta_pdf))
                    File.Move(ruta_pdf, pdf_servidor);
                else
                    MessageBox.Show("El archivo no existe: " + pdf_servidor, "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        bool find = false;
        private void btnpesoxtar_Click(object sender, EventArgs e)
        {
            peso_x_tarimas tar_peso = new peso_x_tarimas();
            if (opcion == 3)
            {
                for (int i = 0; i < DGV1.Rows.Count; i++)
                {
                    find = false;
                    foreach (DataRow row1 in peso_x_tarimas.pesotarima.Select("produc = '" + Convert.ToString(DGV1.Rows[i].Cells["produc"].Value).Trim() + "'"))
                    {
                        find = true;
                        row1["pes_tot"] = Convert.ToDecimal(DGV1.Rows[i].Cells["pes_tot"].Value);
                    }
                    if (find == false)
                    {
                        DataRow row = peso_x_tarimas.pesotarima.NewRow();
                        row["produc"] = Convert.ToString(DGV1.Rows[i].Cells["produc"].Value).Trim();
                        row["nomb"] = Convert.ToString(DGV1.Rows[i].Cells["nomb"].Value).Trim();
                        row["enva"] = Convert.ToString(DGV1.Rows[i].Cells["enva"].Value).Trim();
                        row["canti"] = Convert.ToInt32(DGV1.Rows[i].Cells["canti"].Value);
                        row["env_peso"] = Convert.ToDecimal(DGV1.Rows[i].Cells["env_peso"].Value);
                        row["pes_tot"] = Convert.ToDecimal(DGV1.Rows[i].Cells["pes_tot"].Value);
                        peso_x_tarimas.pesotarima.Rows.Add(row);
                    }
                }
            }
            if (opcion != 3)
            {
                for (int i = 0; i < DGV1.Rows.Count; i++)
                {
                    find = false;
                    foreach (DataRow row1 in peso_x_tarimas.pesotarima.Select("produc = '" + Convert.ToString(DGV1.Rows[i].Cells["produc"].Value).Trim() + "'"))
                    {
                        find = true;
                        row1["pes_tot"] = Convert.ToDecimal(DGV1.Rows[i].Cells["pes_tot"].Value);
                    }
                    if (find == false)
                    {
                        DataRow row = peso_x_tarimas.pesotarima.NewRow();
                        row["produc"] = Convert.ToString(DGV1.Rows[i].Cells["produc"].Value).Trim();
                        row["nomb"] = Convert.ToString(DGV1.Rows[i].Cells["nomb"].Value).Trim();
                        row["enva"] = Convert.ToString(DGV1.Rows[i].Cells["enva"].Value).Trim();
                        row["canti"] = Math.Truncate(Convert.ToDecimal(DGV1.Rows[i].Cells["canti"].Value) / Convert.ToDecimal(DGV1.Rows[i].Cells["tarim"].Value));
                        row["env_peso"] = Convert.ToDecimal(DGV1.Rows[i].Cells["env_peso"].Value);
                        row["pes_tot"] = Convert.ToDecimal(DGV1.Rows[i].Cells["pes_bru"].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells["ta"].Value);
                        peso_x_tarimas.pesotarima.Rows.Add(row);
                    }
                }
            }
            tar_peso.ShowDialog();

            if (opcion == 3)
            {
                for (int i = 0; i < DGV1.Rows.Count; i++)
                {
                    foreach (DataRow row in peso_x_tarimas.pesotarima.Select("produc = '" + Convert.ToString(DGV1.Rows[i].Cells["produc"].Value).Trim() + "'"))
                    {
                        DGV1.Rows[i].Cells["pes_tot"].Value = Convert.ToDecimal(row["pes_uni"].ToString().Trim()) * Convert.ToDecimal(DGV1.Rows[i].Cells["canti"].Value);
                        DGV1.Rows[i].Cells["pes_uni"].Value = Convert.ToDecimal(row["pes_uni"].ToString().Trim());
                        DGV1.Rows[i].Cells["pes_bru"].Value = (20 * Convert.ToInt32(DGV1.Rows[i].Cells["tarim"].Value)) +
                                                          (Convert.ToDecimal(DGV1.Rows[i].Cells["env_peso"].Value) * Convert.ToInt32(DGV1.Rows[i].Cells["canti"].Value)) +
                                                          (Convert.ToDecimal(DGV1.Rows[i].Cells["canti"].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells["pes_uni"].Value));

                        DGV1.Rows[i].Cells["ta"].Value = 0;
                    }
                }
            }
        }

        private void btnmodi_Click(object sender, EventArgs e)
        {
            //btnpesoxtar.Enabled = true;
            //btnGuardar.Enabled = true;
            btnCancel.Enabled = true;
            btncancelrecibo.Enabled = false;
            btnConsulta.Enabled = false;
            //btnpesoxtar.Visible = true;
            opcion = 3;
        }

        private void btnmodeva_Click(object sender, EventArgs e)
        {
            cbevaluacion.Enabled = true;
            btnpdf.Visible = true;
            btnaddfile.Visible = true;
        }

        private void rbcampo_CheckedChanged(object sender, EventArgs e)
        {
            if (rbcampo.Checked == true)
                tipo_recepcion = "PTC";
        }

        private void rbplanta_CheckedChanged(object sender, EventArgs e)
        {
            if (rbplanta.Checked == true)
                tipo_recepcion = "PTP";
        }

        private void cbreccamp_SelectionChangeCommitted(object sender, EventArgs e)
        {
            opcion = 4;
            string[] dividir = cbreccamp.SelectedItem.ToString().Split('~');
            txtrecibo.Text = dividir[0].ToString().Trim();
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "select rpt_fecha, prov_clave, rch_clave, tbl_clave, rpt_codigo, lin_clave, rpt_tipo, vari_clave, rpt_cve_fecha, rpt_inventario " +
                                "from tb_mstr_recepcion_pt where rpt_recibo = '" + dividir[0].ToString().Trim() + "'";
            reader1 = cmnd1.ExecuteReader();
            while (reader1.Read())
            {
                lbfecha.Value = Convert.ToDateTime(reader1.GetValue(0).ToString());
                txtprov_clave.Text = reader1.GetValue(1).ToString().Trim();
                txtrch_clave.Text = reader1.GetValue(2).ToString().Trim();
                txttbl_clave.Text = reader1.GetValue(3).ToString().Trim();
                txtcodigo_clave.Text = reader1.GetValue(4).ToString().Trim();
                txtlin_clave.Text = reader1.GetValue(5).ToString().Trim();
                tipo = reader1.GetValue(6).ToString().Trim();
                txtvariedad.Text = reader1.GetValue(7).ToString().Trim();
                lbcletiqueta.Text = reader1.GetValue(8).ToString().Trim();
                situacion = reader1.GetValue(9).ToString().Trim();
            }
            reader1.Dispose();

            if (situacion.Trim() == "S")
                CBIngProd_fis.Checked = true;
            if (situacion.Trim() == "N")
                CBIngProd_fis.Checked = false;

            if (tipo.Trim() == "CM")
                cbtipo.SelectedIndex = 0;
            if (tipo.Trim() == "TR")
                cbtipo.SelectedIndex = 1;
            if (tipo.Trim() == "MA")
                cbtipo.SelectedIndex = 2;

            txtprov_clave_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
            txtrch_clave_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
            txttbl_clave_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
            foreach (DataRow row in subtablas.Select("prov_clave = '" + txtprov_clave.Text + "' and rch_clave = '" + txtrch_clave.Text + "' and tbl_clave = '" + txttbl_clave.Text + "'"))
                CBcodigo.Items.Add(Convert.ToString(row["tbl_codigo"].ToString().Trim()));

            foreach (DataRow row in lineas.Select("lin_clave = '" + txtlin_clave.Text + "'"))
                CBlin_nombre.SelectedItem = Convert.ToString(row["lin_nombre"].ToString().Trim());

            //trae variedad                        
            foreach (DataRow row in variedades.Select("vari_clave= '" + txtvariedad.Text + "' and lin_clave = '" + txtlin_clave.Text + "'"))
                cbvariedad.Items.Add(Convert.ToString(row["vari_nombre"].ToString().Trim()));

            txtvariedad_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));

            if (txtvariedad.Text.Trim() == "")
            {
                cbvariedad.Items.Clear();
                foreach (DataRow row1 in variedades.Select("lin_clave = '" + txtlin_clave.Text + "'"))
                    cbvariedad.Items.Add(Convert.ToString(row1["vari_nombre"].ToString().Trim()));
            }
            DGV2.Rows.Clear();
            //trae los defectos de acuerdo a la linea
            foreach (DataRow row in danos.Select("lin_clave = '" + txtlin_clave.Text + "'"))
                DGV2.Rows.Add(Convert.ToString(row["dno_nombre"].ToString().Trim()), 0, Convert.ToString(row["dno_clave"].ToString().Trim()));

            DGV3.Rows.Clear();
            //trae los procesos de acuerdo a la linea
            foreach (DataRow row in procesos.Select("lin_clave = '" + txtlin_clave.Text + "'"))
                DGV3.Rows.Add(Convert.ToString(row["proc_nombre"].ToString().Trim()), 0, Convert.ToString(row["proc_clave"].ToString().Trim()));

            DGV1.Rows.Clear();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "select b.prod_clave, c.prod_nombre, a.id_ticket, a.peso_bruto, a.tara, a.envase, b.rptd_cantidad, b.rptd_tarimas, b.rptd_fechacad, a.peso_neto, " +
                                "a.flete, a.num_prod from tb_det_recepcion_bascula a, tb_det_recepcion_pt b, tb_cat_producto c where a.rpt_recibo = '" + dividir[0].ToString().Trim() + "' and " +
                                "b.rpt_recibo = a.rpt_recibo and  a.estatus = 'P' and b.prod_clave = a.prod_clave and c.prod_clave = a.prod_clave";
            reader1 = cmnd1.ExecuteReader();
            while (reader1.Read())
            {
                tar = Convert.ToDecimal(reader1.GetValue(4).ToString());
                b_p = Convert.ToDecimal(reader1.GetValue(3).ToString());
                p_t = Convert.ToDecimal(reader1.GetValue(9).ToString());
                cantidad = Convert.ToInt32(reader1.GetValue(6).ToString());
                tarm = Convert.ToInt32(reader1.GetValue(7).ToString());
                //p_t = b_p - tar - (cantidad - peso_env) - (tarm * 20);
                p_t = b_p - tar - (cantidad * peso_env) - (tarm * 20);

                if (cantidad > 0)
                    p_u = p_t / cantidad;
                else
                    p_u = 0;

                foreach (DataRow row in envases.Select("env_nombre = '" + reader1.GetValue(5).ToString().Trim() + "'"))
                    DGV1.Rows.Add(reader1.GetValue(0).ToString().Trim(), reader1.GetValue(1).ToString().Trim(), reader1.GetValue(2).ToString().Trim(), b_p.ToString("###,###,##0.00"),
                                  tar.ToString("###,###,##0.00"), reader1.GetValue(5).ToString().Trim(), cantidad, tarm, p_u.ToString("###,###,##0.00"), p_t.ToString("###,###,##0.00"),
                                  reader1.GetValue(8).ToString().Trim(), "", row["env_clave"].ToString().Trim(), "", "", reader1.GetValue(11).ToString().Trim());

                //txtflete.Text = reader1.GetValue(10).ToString().Trim();
                if ((reader1.GetValue(10).ToString().Trim() != ""))
                {
                    if (reader1.GetValue(10).ToString().Trim() != "0")
                        flt = reader1.GetValue(10).ToString().Trim();
                }
                else
                    flt = "";

                txtflete.Text = flt;

            }
            reader1.Dispose();

            //tb_mstr_recepcion_bascula
            //trae los datos de tb_mstr_recepcion_bascula
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "select folio, recibio_bascula, hora_peso from tb_mstr_recepcion_bascula where id_ticket = " + dividir[1].ToString().Trim();
            reader1 = cmnd1.ExecuteReader();
            while (reader1.Read())
            {
                //txtflete.Text = reader1.GetValue(0).ToString().Trim();
                txthora.Text = reader1.GetValue(2).ToString().Trim();
                cmnd11 = thisConnection.CreateCommand();
                cmnd11.CommandText = "select bas_nombre from tb_cat_basculeros where bas_clave = '" + reader1.GetValue(1).ToString().Trim() + "'";
                reader11 = cmnd11.ExecuteReader();
                while (reader11.Read())
                {
                    txtpesador.Text = reader11.GetValue(0).ToString().Trim();
                }
                reader11.Dispose();
            }
            reader1.Dispose();
            thisConnection.Close();
            DGV1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            txtvariedad.ReadOnly = false;
            cbvariedad.Enabled = true;
            txtnumviaje.ReadOnly = false;
            txtflete.ReadOnly = false;
            CBMoneda.SelectedIndex = 0;
            CBMoneda.Enabled = true;
            CBcodigo.Enabled = true;
            CBlin_nombre.Enabled = true;
            txtcodigo_clave.ReadOnly = false;
            txtlin_clave.ReadOnly = false;
            txtnumviaje.ReadOnly = false;
            CBIngProd_fis.Enabled = true;
            txtevaluador.ReadOnly = false;
            txtobs.ReadOnly = false;
            txthora.ReadOnly = false;
            DGV2.Columns[1].ReadOnly = false;
            DGV3.Columns[1].ReadOnly = false;
            env.SelectedItem = envas;
            cbevaluacion.Enabled = true;
        }

        private void txtnumped_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back)
            {
                e.Handled = false;
            }
            else if (Char.IsControl(e.KeyChar))
            {
                e.Handled = false;
            }
            if (txtnumped.Text.Trim().Length == 2)
            {
                string pedimento = txtnumped.Text;
                txtnumped.Text = pedimento + "  ";
                txtnumped.Select(4, 0);
            }
            if (txtnumped.Text.Trim().Length == 6)
            {
                string pedimento = txtnumped.Text;
                txtnumped.Text = pedimento + "  ";
                txtnumped.Select(8, 0);
            }
            if (txtnumped.Text.Trim().Length == 12)
            {
                string pedimento = txtnumped.Text;
                txtnumped.Text = pedimento + "  ";
                txtnumped.Select(14, 0);
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            //string impr_blancas = "^XA^HH^XZ";
            string imprverde = "^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597180^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD099 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P099119^FS^FO30,880,^A0N,40,30^FD17:15:24^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P099119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P099119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597181^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD100 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P100119^FS^FO30,880,^A0N,40,30^FD17:15:24^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P100119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P100119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597182^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD101 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P101119^FS^FO30,880,^A0N,40,30^FD17:15:24^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P101119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P101119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597183^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD102 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P102119^FS^FO30,880,^A0N,40,30^FD17:15:24^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P102119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P102119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597184^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD103 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P103119^FS^FO30,880,^A0N,40,30^FD17:15:24^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P103119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P103119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597185^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD104 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P104119^FS^FO30,880,^A0N,40,30^FD17:15:24^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P104119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P104119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597186^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD105 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P105119^FS^FO30,880,^A0N,40,30^FD17:15:24^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P105119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P105119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597187^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD106 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P106119^FS^FO30,880,^A0N,40,30^FD17:15:24^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P106119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P106119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597188^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD107 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P107119^FS^FO30,880,^A0N,40,30^FD17:15:24^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P107119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P107119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597189^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD108 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P108119^FS^FO30,880,^A0N,40,30^FD17:15:24^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P108119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P108119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597190^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD109 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P109119^FS^FO30,880,^A0N,40,30^FD17:15:25^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P109119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P109119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597191^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD110 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P110119^FS^FO30,880,^A0N,40,30^FD17:15:25^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P110119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P110119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597192^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD111 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P111119^FS^FO30,880,^A0N,40,30^FD17:15:25^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P111119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P111119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597193^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD112 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P112119^FS^FO30,880,^A0N,40,30^FD17:15:25^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P112119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P112119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597194^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD113 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P113119^FS^FO30,880,^A0N,40,30^FD17:15:25^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P113119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P113119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597195^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD114 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P114119^FS^FO30,880,^A0N,40,30^FD17:15:25^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P114119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P114119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597196^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD115 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P115119^FS^FO30,880,^A0N,40,30^FD17:15:25^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P115119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P115119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597197^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD116 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P116119^FS^FO30,880,^A0N,40,30^FD17:15:25^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P116119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P116119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597198^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD117 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P117119^FS^FO30,880,^A0N,40,30^FD17:15:25^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P117119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P117119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD31^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597199^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD118 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P118119^FS^FO30,880,^A0N,40,30^FD17:15:25^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P118119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P118119^FS^XZ^XA~SD20^PW832^FO100,25,^A0N,48,35,^FDCLAVE: 4-VIE^FS^FO564,10,^GB150,100,7^FS^FO580,50,^A2N,28,28^FD75^FS^FO30,100,^A0N,48,35^FDSource^FS^FO30,150,^A0N,48,35^FDAGUILARES SPR DE RL.^FS^FO30,230,^A0N,48,35^FDField^FS^FO030,280^A0N,48,35,^FDGARAMBULLO-2             ^FS^FO30,360,^A0N,48,35^FDProduct^FS^FO30,410,^A0N,48,35^FDLECHUGA MR.LUCKY 2 PZAS. ^FS^FO30,450,^A0N,48,35^FD                         ^FS^FO250,500,^A0N,48,35,^FDTrazability Lot #: 314439^FS^FO250,550,^A0N,48,35,^FDLoad #: 1^FS^FO250,600,^A0N,48,35,^FDTime in ^FS^FO180,650,^GB500,100,7^FS^FO15,950^ABN,25,17^FDSSCC^FS^FO150,930,^BY3,^BCN,60,Y,N,Y,N^FD>;>80000796631000597200^FS^FO700,1050^A0N,20,18^FDF-200-37^FS^FO700,1075^A0N,20,18^FDREV: 04^FS^FO196,690,^A2N,33,23^FD119 of 119 Pallets^FS^FO340,750^BQN,2,7^FDLA,31443905005ML2P119119^FS^FO30,880,^A0N,40,30^FD17:15:25^FS^FO580,880,^A0N,40,30^FD12/02/2023^FS^FO120,1060,^BY2,^BCN,120,N,N,N^FD31443905005ML2P119119^FS^FO140,1190,^A2N,35,20,^FD31443905005ML2P119119^FS^XZ";
            RawPrinterHelper.SendStringToPrinter("etiqueta_blanca", imprverde);
            string impr_blancas = "zpl.label_length";

            PrintDialog pd = new PrintDialog();
            pd.PrinterSettings = new PrinterSettings();
            if (DialogResult.OK == pd.ShowDialog(this))
            {
                // Send a printer-specific to the printer.
                RawPrinterHelper.SendStringToPrinter(pd.PrinterSettings.PrinterName, impr_blancas);
                MessageBox.Show("Impresora a Configuracion Default Correctamente", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            thisConnection.Open();
            for (int i = 0; i < DGV1.Rows.Count; i++)
            {
                decimal hrp_peso_neto = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                decimal hrp_clase1 = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                decimal hrp_peso_util = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                decimal hrp_reman_kg = Convert.ToDecimal(DGV1.Rows[i].Cells[3].Value) - Convert.ToDecimal(DGV1.Rows[i].Cells[4].Value) - (Convert.ToDecimal(DGV1.Rows[i].Cells[14].Value) * Convert.ToDecimal(DGV1.Rows[i].Cells[6].Value)) - (Convert.ToDecimal(DGV1.Rows[i].Cells[7].Value) * 20);
                int hrp_num_unidades = Convert.ToInt32(DGV1.Rows[i].Cells[6].Value);
                int hrp_reman_unidades = Convert.ToInt32(DGV1.Rows[i].Cells[6].Value);

                cmnd1 = thisConnection.CreateCommand();
                if (cbevaluacion.SelectedIndex == 0 || cbevaluacion.SelectedIndex == 1)
                    cmnd1.CommandText = "INSERT INTO tb_hist_recepcion (hrp_recibo, hrp_fecha, lin_clave, hrp_tipo_recepcion, hrp_estatus, prod_clave, hrp_peso_neto, hrp_clase1, " +
                                        "hrp_peso_util, " +
                                        "hrp_num_unidades, hrp_reman_kg, hrp_reman_unidades, hrp_situacion, hrp_clase2, hrp_proceso, hrp_surtido, hrp_liquidado, hrp_numliq) values " +
                                        "(" + txtrecibo.Text + ", '" + lbfecha.Text + "', '" + txtlin_clave.Text + "', 'PTC', 'T', " +
                                        "'" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "'," + hrp_peso_neto + ", " + hrp_clase1 + ", " + hrp_peso_util + "," + hrp_num_unidades + ", " + hrp_reman_kg + ", " +
                                        "" + hrp_reman_unidades + ", '" + tipo + "', 0, 0, '" + hrp_peso_neto + "', '', '')";
                if (cbevaluacion.SelectedIndex == 2)
                    cmnd1.CommandText = "INSERT INTO tb_hist_recepcion (hrp_recibo, hrp_fecha, lin_clave, hrp_tipo_recepcion, hrp_estatus, prod_clave, hrp_peso_neto, hrp_clase1, " +
                                        "hrp_peso_util, " +
                                        "hrp_num_unidades, hrp_reman_kg, hrp_reman_unidades, hrp_situacion, hrp_clase2, hrp_proceso, hrp_surtido, hrp_liquidado, hrp_numliq) " +
                                        "values (" + txtrecibo.Text + ", " +
                                        "'" + lbfecha.Text + "', '" + txtlin_clave.Text + "', 'PTC', 'C', " +
                                        "'" + Convert.ToString(DGV1.Rows[i].Cells[0].Value) + "'," + hrp_peso_neto + ", " + hrp_clase1 + ", " + hrp_peso_util + ", " +
                                        "" + hrp_num_unidades + ", " + hrp_reman_kg + ", " +
                                        "" + hrp_reman_unidades + ", '" + tipo + "', 0, 0, '" + hrp_peso_neto + "', '', '')";
                reader1 = cmnd1.ExecuteReader();
                reader1.Dispose();
            }
            thisConnection.Close();
            MessageBox.Show("Historico de Recepcion Actualizado");
        }

        private void BtnFecCad_Click(object sender, EventArgs e)
        {
            FrmFecCad FrmFec = new FrmFecCad();
            FrmFec.ShowDialog(this);
        }

        private void DGV2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (DGV2.CurrentCell.ColumnIndex == 1) // Cantidad
            {
                decimal TotMtra = decimal.Parse(TxtPieza.Text);
                if (TotMtra > 0)
                {
                    decimal TotPorce = 0, Porce = 0;
                    //if (DGV2.CurrentRow.Cells["TipCalidad"].Value.ToString().Trim() == "GRADO 1")
                    //{
                    Porce = (Convert.ToDecimal(DGV2.CurrentRow.Cells[1].Value) * 100) / TotMtra;
                    DGV2.CurrentRow.Cells[3].Value = Porce.ToString("##0.#0");
                    //}
                    for (int i = 0; i < DGV2.Rows.Count; i++) // Daños
                    {
                        if (DGV2.Rows[i].Cells["TipCalidad"].Value.ToString().Trim() == "GRADO 1")
                            TotPorce += Convert.ToDecimal(DGV2.Rows[i].Cells[3].Value);
                    }
                    for (int i = 0; i < DGV3.Rows.Count; i++) // Procesos
                        TotPorce += Convert.ToDecimal(DGV3.Rows[i].Cells[3].Value);
                    LblGra1.Text = (Convert.ToDecimal(100) - TotPorce).ToString("##0.#0");
                }
            }
        }

        private void DGV3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (DGV3.CurrentCell.ColumnIndex == 1) // Cantidad
            {
                decimal TotMtra = decimal.Parse(TxtPieza.Text);
                if (TotMtra > 0)
                {
                    decimal TotPorce = 0;
                    decimal Porce = (Convert.ToDecimal(DGV3.CurrentRow.Cells[1].Value) * 100) / TotMtra;
                    DGV3.CurrentRow.Cells[3].Value = Porce.ToString("##0.#0");
                    for (int i = 0; i < DGV2.Rows.Count; i++)
                        if (DGV2.Rows[i].Cells["TipCalidad"].Value.ToString().Trim() == "GRADO 1")
                            TotPorce += Convert.ToDecimal(DGV2.Rows[i].Cells[3].Value);
                    for (int i = 0; i < DGV3.Rows.Count; i++)
                        TotPorce += Convert.ToDecimal(DGV3.Rows[i].Cells[3].Value);
                    LblGra1.Text = (Convert.ToDecimal(100) - TotPorce).ToString("##0.#0");
                }
            }
        }

        private void SentToOutlook()
        {
            string destinatarios = "mdelrio@mrlucky.com.mx;mrhernandez@mrlucky.com.mx;musabiaga@mrlucky.com.mx;msamano@mrlucky.com.mx;armandog@mrlucky.com.mx;arturonieto@grupou.mx;omaraleman@mrlucky.com.mx;ochavez@grupou.mx;cmoreno@mrlucky.com.mx;comprasmp@mrlucky.com.mx;jantonio@grupou.mx;estebanm@grupou.mx;mprima@mrlucky.com.mx;armandog@mrlucky.com.mx";
            string subject = "Línea. " + CBlin_nombre.Text.Trim() + "   Recibo: " + txtrecibo.Text + " Del Día:  " + lbfecha.Value.ToShortDateString();
            string body = //"<H3><p style='color: blue'>Envio Información Del Flete de Personal " + FolioFle.ToString() + " </p></H3>" + 
                        "Hola Buen día" + "<br><br>" + "<H3><p style='color: blue'>Envío información de la EVALUACION del RECIBO " + txtrecibo.Text + "  Flete: " + txtflete.Text.Trim() + "</p></H3><br><br>" +
                        "<table border=\"1\"> " +
                       "<tr> <TH bgcolor='lightblue' COLSPAN=4>PRODUCTOS</TH></TR> " +
                       //"<tr> <TH COLSPAN=3>TABLA: " + txttabla.Text.Trim() + "           VARIEDAD:  " + CBvariedad.Text  +" </TH></TR> " +
                       "<tr> <td ALING=LEFT bgcolor='lightgreen'>TABLA: " + txttabla.Text.Trim() + "</td> " +
                            "<td ALING=LEFT>           </td> " +
                            "<td ALING=LEFT bgcolor='lightgreen'>VARIEDAD:  " + cbvariedad.Text + " </td></tr> " +
                       //"<tr> <td ALING=LEFT bgcolor='lightgreen'>LLEGARON: " + txtrmpnumenvase.Text.ToString() + " " + txtenv.Text.Trim() + "</td> " +
                       //     "<td ALING=LEFT>          </td> " +
                       //     "<td ALING=LEFT bgcolor='lightgreen'>PESO DE " + (Convert.ToDecimal(txtpesoutil.Text) / Convert.ToDecimal(txtrmpnumenvase.Text)).ToString("#,###.00") + " KG" + " </td></TR> " +
                       " <tr>   " +
                       " <th scope=\"col\">PRODUCTO</strong></th> " +
                       " <th scope=\"col\">CAJAS </strong></th> " +
                       " <th scope=\"col\">TARIMAS</strong></th> " +
                       " <th scope=\"col\">PESO UNITARIO</strong></th> " +
                       " </tr>";
            string cap = "";
            for (int i = 0; i < DGV1.Rows.Count; i++)
            {
                cap += "<tr>" +
                       "<td ALING=LEFT>" + DGV1.Rows[i].Cells["Nomb"].Value.ToString().Trim() + "</td>" +
                       "<td ALING=CENTER>" + Convert.ToDecimal(DGV1.Rows[i].Cells["Canti"].Value).ToString("#,##0") + "</td>" +
                       "<td ALING=CENTER>" + Convert.ToDecimal(DGV1.Rows[i].Cells["Tarim"].Value).ToString("#,##0") + "</td>" +
                       "<td ALING=CENTER>" + Convert.ToDecimal(DGV1.Rows[i].Cells["pes_uni"].Value).ToString("#,##0.00") + " Kg</td>" +
                       "</tr>";
            }
            cap += "<tr>" +
                   "<TH bgcolor='yellow' COLSPAN=4>No. de Piezas: " + TxtPieza.Text + "      GRADO 1:   " + Convert.ToDecimal(LblGra1.Text).ToString("#,###.00") + " %</strong></td>" +
                   "</tr></table><br><br>";
            if (DGV2.Rows.Count > 0)
            {
                cap += "<table border=\"1\"> " +
                       "<tr> <TH COLSPAN=3> DAÑOS </TH></TR> " +
                       "<tr>" +
                       " <th scope=\"col\">DAÑOS</strong></th> " +
                       " <th scope=\"col\">No. de Piezas: " + TxtPieza.Text + " </strong></th> " +
                       " <th scope=\"col\">% DAÑO</strong></th> " +
                       " </tr>";

                //<tr><td>NOMBRE UNO</td>10<td></td>5.00 %<td></td></tr> ";

                for (int i = 0; i < DGV2.Rows.Count; i++)
                {
                    if (Convert.ToDecimal(DGV2.Rows[i].Cells[1].Value.ToString()) > 0 && DGV2.Rows[i].Cells[4].Value.ToString().Contains("GRADO 1"))
                    {
                        cap += "<tr>" +
                               "<td ALING=LEFT>" + DGV2.Rows[i].Cells[0].Value.ToString().Trim() + "</td>" +
                               "<td ALING=CENTER>" + Convert.ToDecimal(DGV2.Rows[i].Cells[1].Value).ToString("#,##0.00") + " </td>" +
                               "<td ALING=CENTER>" + Convert.ToDecimal(DGV2.Rows[i].Cells[3].Value).ToString("#,##0.00") + " %</td>" +
                               "</tr>";
                    }
                }
                cap += "<tr> <TH COLSPAN=3></TH></TR> " +
                       "<tr> <TH COLSPAN=3></TH></TR> ";
                int j = 0;
                for (int i = 0; i < DGV2.Rows.Count; i++)
                {
                    if (Convert.ToDecimal(DGV2.Rows[i].Cells[1].Value.ToString()) > 0 && DGV2.Rows[i].Cells[4].Value.ToString().Contains("PARAM"))
                    {
                        if (j == 0)
                        {
                            cap += "<tr> <TH COLSPAN=3> PARAMETROS </TH></TR> ";
                            j++;
                        }
                        cap += "<tr>" +
                               "<td ALING=LEFT>" + DGV2.Rows[i].Cells[0].Value.ToString().Trim() + "</td>" +
                               "<td ALING=CENTER>" + Convert.ToDecimal(DGV2.Rows[i].Cells[1].Value).ToString("#,##0.00") + " </td>" +
                               "<td ALING=CENTER>" + Convert.ToDecimal(DGV2.Rows[i].Cells[3].Value).ToString("#,##0.00") + " %</td>" +
                               "</tr>";
                    }
                }
                cap += "</tr></table><br><br>";
            }
            if (DGV3.Rows.Count > 0)
            {
                cap += "<table border=\"1\"> " +
                       "<tr> <TH COLSPAN=3> PROCESO </TH></TR> " +
                       "<tr>" +
                       " <th scope=\"col\">PROCESO</strong></th> " +
                       " <th scope=\"col\">No. de Piezas: " + TxtPieza.Text + " KG</strong></th> " +
                       " <th scope=\"col\">% DAÑO</strong></th> " +
                       " </tr>";

                //<tr><td>NOMBRE UNO</td>10<td></td>5.00 %<td></td></tr> ";

                for (int i = 0; i < DGV3.Rows.Count; i++)
                {
                    if (Convert.ToDecimal(DGV3.Rows[i].Cells[1].Value.ToString()) > 0)
                    {
                        cap += "<tr>" +
                               "<td ALING=LEFT>" + DGV3.Rows[i].Cells[0].Value.ToString().Trim() + "</td>" +
                               "<td ALING=CENTER>" + Convert.ToDecimal(DGV3.Rows[i].Cells[1].Value).ToString("#,##0.00") + " </td>" +
                               "<td ALING=CENTER>" + Convert.ToDecimal(DGV3.Rows[i].Cells[3].Value).ToString("#,##0.00") + " %</td>" +
                               "</tr>";
                    }
                }
                cap += "</tr></table><br><br>";
            }
            //cap += "<table border=\"1\"> " +
            //       "<tr> <TH COLSPAN=1> INSECTOS </TH></TR> " +
            //       "<tr>" +
            //       "<td ALING=CENTER>" + txtdescrip.Text.Trim() + "</td>" +
            //       "</tr></table>";
            body = body + cap;
            //276132
            // Inicializar la aplicación de Outlook
            Outlook.Application outlookApp = new Outlook.Application();

            // Crear un nuevo correo
            Outlook.MailItem mail = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
            mail.Subject = subject;
            mail.To = destinatarios;
            mail.BCC = "ricardo.cortes@mrlucky.com.mx";
            mail.HTMLBody = body;

            // Mostrar el correo para que el usuario lo edite
            mail.Display(false);
        }
    }
}
