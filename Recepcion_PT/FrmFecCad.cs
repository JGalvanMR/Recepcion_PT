using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Recepcion_PT
{
    public partial class FrmFecCad : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);

        public FrmFecCad()
        {
            InitializeComponent();
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(@"C:\SisGabWeb\fondo_formularios.jpg");
        }

        private void FrmFecCad_Load(object sender, EventArgs e)
        {

        }

        private void BtnGene_Click(object sender, EventArgs e)
        {
            thisConnection.Open();
            string Cadena = "Select prod_nombre, tarima, etiqueta, Fecha_Cad, pti_fecha From tb_det_trazabilidad WHERE recibo = '" + TxtRecibo.Text +"' AND prod_clave = '03001ML09 ' " +
                            "and pti_estatus_sur = ' ' ORDER BY prod_nombre, tarima";
            SqlCommand cmd = new SqlCommand(Cadena, thisConnection);
            SqlDataReader reader ;
            reader = cmd.ExecuteReader();
            DgDatos.Rows.Clear();
            DateTime FecProd = System.DateTime.Today; 
            while (reader.Read())
            {
                DgDatos.Rows.Add(reader["prod_nombre"].ToString(), reader["Tarima"].ToString(), reader["Etiqueta"].ToString(),reader["Fecha_Cad"].ToString(), false);
                FecProd = Convert.ToDateTime(reader["pti_fecha"]);
            }
            thisConnection.Close();
            if (DgDatos.Rows.Count > 0)
            {
                DtFC.Enabled = true;
                BtnSave.Enabled = true;
                DtFC.MinDate = FecProd.AddDays(10);
                DtFC.MaxDate = FecProd.AddDays(22);
                LblFeEla.Text = "Fecha de Elaboración: " + FecProd.ToString("dd/MMM/yyyy").Replace(".", "");
                LblFeEla.Visible = true;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Esta Seguro de Guardar la Información las Tarimas Asignadas se Grabaran con la siguiente Fecha de Caducidad "+ DtFC.Value.ToString("dd/MMM/yy").Replace(".","") , "Aviso Grabar Datos", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                return;
            thisConnection.Open();
            int i = 0, j= 0;
            string mTar = "";
            foreach (DataGridViewRow row in DgDatos.Rows)
            {
                if (Convert.ToBoolean(DgDatos.Rows[i].Cells["ASIGNAR"].Value) == true)
                {
                    string Cadena = "Update tb_det_trazabilidad set fecha_cad = '" + DtFC.Value.ToString("dd/MM/yyyy") + "' " +
                                    "WHERE recibo = '" + TxtRecibo.Text + "' AND prod_clave = '03001ML09' and tarima = '" + DgDatos.Rows[i].Cells["tarima"].Value.ToString() + "'";
                    SqlCommand cmd = new SqlCommand(Cadena, thisConnection);
                    cmd.ExecuteNonQuery();
                    mTar += DgDatos.Rows[i].Cells["tarima"].Value.ToString().Trim() +",";
                    j++;
                    Cadena = "Update tb_det_recepcion_pt set rptd_fechacad = '" + DtFC.Value.ToString("dd/MM/yyyy") + "' " +
                             "WHERE rpt_recibo = '" + TxtRecibo.Text + "' AND prod_clave = '03001ML09'";
                    cmd = new SqlCommand(Cadena, thisConnection);
                    cmd.ExecuteNonQuery();
                }
                i++;
            }
            thisConnection.Close();
            if (j > 0)
            {
                MessageBox.Show("La informacion se Grabo correctamente!!!", "Grbar información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "M", "2.3", TxtRecibo.Text, "Se Asigno Fecha de Caducidad " + DtFC.Value.ToString("dd/MM/yyyy") + " Tar: " + mTar , "SIPGAB");
                DtFC.Enabled = false;
                BtnSave.Enabled = false;
                DgDatos.Rows.Clear();
                LblFeEla.Visible = false;
            }
            
        }
    }
}
