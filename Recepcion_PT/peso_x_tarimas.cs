using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Recepcion_PT
{
    public partial class peso_x_tarimas : Form
    {
        public static DataTable pesotarima = new DataTable("pesotarima");

        public peso_x_tarimas()
        {
            InitializeComponent();
            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
            if (!(pesotarima.Columns.Contains("produc")))
            {
                pesotarima.Columns.Add("produc");
                pesotarima.Columns.Add("nomb");
                pesotarima.Columns.Add("enva");
                pesotarima.Columns.Add("canti");
                pesotarima.Columns.Add("env_peso");
                pesotarima.Columns.Add("pes_tot");
                pesotarima.Columns.Add("pes_uni");
            }
        }

        private void peso_x_tarimas_Load(object sender, EventArgs e)
        {
            if (Recepcion_PT.opcion == 3)
            {
                foreach (DataRow row in pesotarima.Rows)
                    DGV.Rows.Add(row["produc"].ToString().Trim(), row["nomb"].ToString().Trim(), row["pes_tot"].ToString().Trim(), 0, row["enva"].ToString().Trim(), row["env_peso"].ToString().Trim(), 0, 0);
            }
            if (Recepcion_PT.opcion != 3)
            {
                foreach (DataRow row in pesotarima.Rows)
                    DGV.Rows.Add(row["produc"].ToString().Trim(), row["nomb"].ToString().Trim(), row["pes_tot"].ToString().Trim(), row["canti"].ToString().Trim(), row["enva"].ToString().Trim(), row["env_peso"].ToString().Trim(), 0, 0);

                DGV.Rows[0].Cells[2].Selected = true;
                DGV.EndEdit();
                for (int i = 0; i < DGV.Rows.Count; i++)
                {
                    DGV.Rows[i].Cells["peso_total"].Value = Convert.ToDecimal(DGV.Rows[i].Cells["peso_tar"].Value) - 20
                                                                 - (Convert.ToDecimal(DGV.Rows[i].Cells["num_box"].Value) * Convert.ToDecimal(DGV.Rows[i].Cells["peso_env"].Value));

                    foreach (DataRow row in pesotarima.Select("produc = '" + Convert.ToString(DGV.Rows[i].Cells[0].Value) + "'"))
                    {
                        row["pes_tot"] = DGV.Rows[i].Cells["peso_total"].Value.ToString();

                        if (Convert.ToDecimal(DGV.Rows[i].Cells["num_box"].Value) > 0)
                        {
                            DGV.Rows[i].Cells["peso_uni"].Value = (Convert.ToDecimal(DGV.Rows[i].Cells["peso_total"].Value.ToString()) / Convert.ToDecimal(DGV.Rows[i].Cells["num_box"].Value)).ToString("##0.00");
                            row["pes_uni"] = Convert.ToDecimal(DGV.Rows[i].Cells["peso_uni"].Value).ToString("##0.00");
                        }
                    }
                }

            }
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        private void btnexit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void DGV_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (DGV.CurrentCell.ColumnIndex == 2)
            {
                DGV.Rows[e.RowIndex].Cells["peso_total"].Value = Convert.ToDecimal(DGV.Rows[e.RowIndex].Cells["peso_tar"].Value) - 20
                                                                 - (Convert.ToDecimal(DGV.Rows[e.RowIndex].Cells["num_box"].Value) * Convert.ToDecimal(DGV.Rows[e.RowIndex].Cells["peso_env"].Value));
                foreach (DataRow row in pesotarima.Select("produc = '" + Convert.ToString(DGV.Rows[e.RowIndex].Cells[0].Value) + "'"))
                {
                    row["pes_tot"] = DGV.Rows[e.RowIndex].Cells["peso_total"].Value.ToString();

                    if (Convert.ToDecimal(DGV.Rows[e.RowIndex].Cells["num_box"].Value) > 0)
                    {
                        DGV.Rows[e.RowIndex].Cells["peso_uni"].Value = (Convert.ToDecimal(DGV.Rows[e.RowIndex].Cells["peso_total"].Value.ToString()) / Convert.ToDecimal(DGV.Rows[e.RowIndex].Cells["num_box"].Value)).ToString("##0.00");
                        row["pes_uni"] = Convert.ToDecimal(DGV.Rows[e.RowIndex].Cells["peso_uni"].Value).ToString("##0.00");
                    }
                }
            }
            if (DGV.CurrentCell.ColumnIndex == 3)
            {
                DGV.Rows[e.RowIndex].Cells["peso_total"].Value = Convert.ToDecimal(DGV.Rows[e.RowIndex].Cells["peso_tar"].Value) - 20
                                                                 - (Convert.ToDecimal(DGV.Rows[e.RowIndex].Cells["num_box"].Value) * Convert.ToDecimal(DGV.Rows[e.RowIndex].Cells["peso_env"].Value));
                foreach (DataRow row in pesotarima.Select("produc = '" + Convert.ToString(DGV.Rows[e.RowIndex].Cells[0].Value) + "'"))
                {
                    row["pes_tot"] = DGV.Rows[e.RowIndex].Cells["peso_total"].Value.ToString();

                    if (Convert.ToDecimal(DGV.Rows[e.RowIndex].Cells["num_box"].Value) > 0)
                    {
                        DGV.Rows[e.RowIndex].Cells["peso_uni"].Value = (Convert.ToDecimal(DGV.Rows[e.RowIndex].Cells["peso_total"].Value.ToString()) / Convert.ToDecimal(DGV.Rows[e.RowIndex].Cells["num_box"].Value)).ToString("##0.00");
                        row["pes_uni"] = Convert.ToDecimal(DGV.Rows[e.RowIndex].Cells["peso_uni"].Value).ToString("##0.00");
                    }
                }
            }
        }

        private void btnguardar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
