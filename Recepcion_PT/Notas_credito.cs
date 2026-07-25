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

namespace Recepcion_PT
{
    public partial class Notas_credito : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlCommand cmnd1 = new SqlCommand();
        SqlCommand cmnd11 = new SqlCommand();        
        SqlDataReader reader1, reader11;
        public DataTable ntc = new DataTable();
        public string rptrecibo = "", rpttiporecepcion = "", prod = "", prodnom = "", envnombre, moti, line = "", line1 = "", nomline = "", proveclave, provenom, factura = "";
        public bool estado = false;

        public Notas_credito()
        {
            InitializeComponent();
            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
        }

        private void Notas_credito_Load(object sender, EventArgs e)
        {
            try
            {
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "select A.rpt_recibo, A.rpt_fecha, A.prov_clave, B.prov_nombre, C.lin_nombre, A.tipo_nota, A.rpt_Tiporecibo, A.fcn_folio, A.lin_clave " +
                                    "from tb_tmp_mstr_nota as A, tb_cat_proveedor as B, tb_cat_linea as C " +
                                    "where A.autorizado = 'S' and A.afectado = '' and B.prov_clave = A.prov_clave and C.lin_clave = A.lin_clave";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    DGV6.Rows.Add(reader1.GetValue(0).ToString().Trim(), reader1.GetValue(1).ToString().Substring(0, 9), reader1.GetValue(2).ToString().Trim(), reader1.GetValue(3).ToString().Trim(), reader1.GetValue(4).ToString().Trim(), reader1.GetValue(5).ToString().Trim(), reader1.GetValue(6).ToString().Trim(), reader1.GetValue(7).ToString().Trim(), reader1.GetValue(8).ToString().Trim());
                    Notas_credito nc = new Notas_credito();
                    nc.Focus();
                }
                reader1.Dispose();
                thisConnection.Close();
                DGV6.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                estado = true;
                thisConnection.Close();
                ntc.Columns.Add("produc");
                ntc.Columns.Add("nomb");
                ntc.Columns.Add("enva");
                ntc.Columns.Add("canti");
                ntc.Columns.Add("linea");
                ntc.Columns.Add("lin");
            }
            catch (SqlException EX)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", EX.ToString().Trim());
                MessageBox.Show(EX.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void DGV6_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (estado == true)
            {
                //if (DGV6.Rows.Count > 1)
                //{
                DGV7.Rows.Clear();
                rptrecibo = Convert.ToString(DGV6.Rows[e.RowIndex].Cells[0].Value);
                rpttiporecepcion = Convert.ToString(DGV6.Rows[e.RowIndex].Cells[6].Value);
                factura = Convert.ToString(DGV6.Rows[e.RowIndex].Cells[7].Value);
                if (thisConnection.State != ConnectionState.Open)
                    thisConnection.Open();

                cmnd11 = thisConnection.CreateCommand();
                cmnd11.CommandText = "select A.prod_clave, B.prod_nombre, A.rptd_cantidad, A.afectado, A.rpt_Tiporecibo, B.lin_clave, D.lin_nombre " +
                                     "from tb_tmp_det_nota A, tb_cat_producto B, tb_cat_linea D where A.rpt_recibo = '" + rptrecibo + "' and A.rpt_Tiporecibo = '" + rpttiporecepcion + "' and " +
                                    "B.prod_clave = A.prod_clave and D.lin_clave = B.lin_clave and A.afectado <> 'S' and A.fcn_folio = '" + factura + "'";
                reader11 = cmnd11.ExecuteReader();
                while (reader11.Read())
                {
                    DGV7.Rows.Add(reader11.GetValue(0).ToString().Trim(), reader11.GetValue(1).ToString(), reader11.GetValue(2).ToString(), reader11.GetValue(3).ToString(), reader11.GetValue(4).ToString().Trim(), reader11.GetValue(5).ToString(), reader11.GetValue(6).ToString(), "", "", rptrecibo, "");
                }
                reader11.Dispose();
                thisConnection.Close();
                DGV7.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                //}
            }
        }

        private void DGV7_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                for (int i = 0; i < DGV7.Rows.Count; i++)
                {
                    if (Convert.ToString(DGV7.Rows[i].Cells[3].Value).Trim() != "S")
                    {
                        line = DGV7.Rows[i].Cells[6].Value.ToString().Trim();
                        //line = DGV7.Rows[i].Cells[0].Value.ToString().Substring(0, 2);
                        if (i == 0)
                        {
                            DataRow rw = ntc.NewRow();
                            rw[0] = DGV7.Rows[i].Cells[0].Value.ToString();
                            rw[1] = DGV7.Rows[i].Cells[1].Value.ToString();
                            rw[2] = DGV7.Rows[i].Cells[4].Value.ToString();
                            rw[3] = DGV7.Rows[i].Cells[2].Value.ToString();                            
                            rw[4] = DGV7.Rows[i].Cells[6].Value.ToString();
                            rw[5] = DGV7.Rows[i].Cells[5].Value.ToString();
                            ntc.Rows.Add(rw);
                        }
                        else
                        {
                            //line = DGV7.Rows[i].Cells[6].Value.ToString().Trim();
                            if (ntc.Rows.Count > 0)
                            {
                                line1 = ntc.Rows[ntc.Rows.Count - 1]["linea"].ToString().Trim();

                                if (line == line1)
                                {
                                    DataRow rw = ntc.NewRow();
                                    rw[0] = DGV7.Rows[i].Cells[0].Value.ToString();
                                    rw[1] = DGV7.Rows[i].Cells[1].Value.ToString();
                                    rw[2] = DGV7.Rows[i].Cells[4].Value.ToString();
                                    rw[3] = DGV7.Rows[i].Cells[2].Value.ToString();                                    
                                    rw[4] = DGV7.Rows[i].Cells[6].Value.ToString();
                                    rw[5] = DGV7.Rows[i].Cells[5].Value.ToString();
                                    ntc.Rows.Add(rw);
                                }
                            }
                            else
                            {
                                DataRow rw = ntc.NewRow();
                                rw[0] = DGV7.Rows[i].Cells[0].Value.ToString();
                                rw[1] = DGV7.Rows[i].Cells[1].Value.ToString();
                                rw[2] = DGV7.Rows[i].Cells[4].Value.ToString();
                                rw[3] = DGV7.Rows[i].Cells[2].Value.ToString();                                
                                rw[4] = DGV7.Rows[i].Cells[6].Value.ToString();
                                rw[5] = DGV7.Rows[i].Cells[5].Value.ToString();
                                ntc.Rows.Add(rw);
                            }
                        }
                    }
                }
                proveclave = DGV6.Rows[e.RowIndex].Cells[2].Value.ToString();
                provenom = DGV6.Rows[e.RowIndex].Cells[3].Value.ToString();
                nomline = ntc.Rows[ntc.Rows.Count - 1]["linea"].ToString().Trim();
                line = ntc.Rows[ntc.Rows.Count - 1]["lin"].ToString().Trim();
                moti = "N. Crédito x " + Convert.ToString(DGV7.Rows[e.RowIndex].Cells[4].Value) + " de la factura " + Convert.ToString(DGV6.Rows[e.RowIndex].Cells[7].Value);
                
                Recepcion_PT.notacre = true;
                Recepcion_PT.tipo_notcre = Convert.ToString(DGV7.Rows[e.RowIndex].Cells["tiporecep"].Value);
                Recepcion_PT.fcn_folioNCR = Convert.ToString(DGV6.Rows[e.RowIndex].Cells["fcn_folio"].Value);
                this.Close();
            }
        }

        private void DGV6_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if (e.KeyChar == (char)Keys. || e.KeyChar == (char)Keys.Up)
            //{
            //    DGV7.Rows.Clear();
            //    //rptrecibo = Convert.ToString(DGV6.Rows[e.RowIndex].Cells[0].Value);
            //    //rpttiporecepcion = Convert.ToString(DGV6.Rows[e.RowIndex].Cells[6].Value);
            //    rptrecibo = Convert.ToString(DGV6.Rows[DGV6.CurrentRow.Index].Cells[0].Value);
            //    rpttiporecepcion = Convert.ToString(DGV6.Rows[DGV6.CurrentRow.Index].Cells[6].Value);
            //    if (thisConnection.State != ConnectionState.Open)
            //        thisConnection.Open();

            //    cmnd11 = thisConnection.CreateCommand();
            //    cmnd11.CommandText = "select A.prod_clave, B.prod_nombre, A.rptd_cantidad, A.afectado, A.rpt_Tiporecibo, A.lin_clave, D.lin_nombre " +
            //                         "from tb_tmp_det_nota A, tb_cat_producto B, tb_cat_linea D where A.rpt_recibo = '" + rptrecibo + "' and A.rpt_Tiporecibo = '" + rpttiporecepcion + "' and " +
            //                        "B.prod_clave = A.prod_clave and D.lin_clave = A.lin_clave";
            //    reader11 = cmnd11.ExecuteReader();
            //    while (reader11.Read())
            //    {
            //        DGV7.Rows.Add(reader11.GetValue(0).ToString().Trim(), reader11.GetValue(1).ToString(), reader11.GetValue(2).ToString(), reader11.GetValue(3).ToString(), reader11.GetValue(4).ToString().Trim(), reader11.GetValue(5).ToString(), reader11.GetValue(6).ToString(), "", "", rptrecibo, "");
            //    }
            //    reader11.Dispose();
            //    thisConnection.Close();
            //    DGV7.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            //}
        }      
    }
}
 
