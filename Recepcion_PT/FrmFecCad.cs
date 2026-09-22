using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Recepcion_PT
{
    public partial class FrmFecCad : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);

        public FrmFecCad()
        {
            InitializeComponent();
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(@"C:\SisGabWeb\fondo_formularios.jpg");

            // Agregar la nueva columna "Limpiar" en tiempo de ejecución (si no existe en el diseñador)
            if (!DgDatos.Columns.Contains("Limpiar"))
            {
                DataGridViewCheckBoxColumn colLimpiar = new DataGridViewCheckBoxColumn();
                colLimpiar.Name = "Limpiar";
                colLimpiar.HeaderText = "Limpiar";
                colLimpiar.Width = 60;
                colLimpiar.ReadOnly = false;
                // Insertar después de la columna "Fecha_Cad" (asumiendo que está en la posición 3)
                DgDatos.Columns.Insert(4, colLimpiar);
            }
        }

        private void FrmFecCad_Load(object sender, EventArgs e)
        {
            // Asegurar que el checkbox funcione con un solo clic
            DgDatos.CellContentClick += DgDatos_CellContentClick;
        }

        private void BtnGene_Click(object sender, EventArgs e)
        {
            thisConnection.Open();
            string Cadena = "Select prod_nombre, tarima, etiqueta, Fecha_Cad, pti_fecha From tb_det_trazabilidad WHERE recibo = '" + TxtRecibo.Text + "' AND prod_clave = '03001ML09 ' " +
                            "and pti_estatus_sur = ' ' ORDER BY prod_nombre, tarima";
            SqlCommand cmd = new SqlCommand(Cadena, thisConnection);
            SqlDataReader reader;
            reader = cmd.ExecuteReader();
            DgDatos.Rows.Clear();
            DateTime FecProd = System.DateTime.Today;
            while (reader.Read())
            {
                // Agregar fila con los 5 valores: nombre, tarima, etiqueta, fecha, y el checkbox "ASIGNAR" (false)
                DgDatos.Rows.Add(
                    reader["prod_nombre"].ToString(),
                    reader["Tarima"].ToString(),
                    reader["Etiqueta"].ToString(),
                    reader["Fecha_Cad"].ToString(),
                    false,    // ASIGNAR
                    false     // Limpiar (nuevo)
                );
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
            // Contar cuántas filas tienen "Limpiar" marcado y cuántas solo "ASIGNAR"
            int filasLimpiar = 0, filasAsignar = 0;
            string tarimasLimpiar = "", tarimasAsignar = "";

            foreach (DataGridViewRow row in DgDatos.Rows)
            {
                if (row.Cells["Limpiar"].Value != null && (bool)row.Cells["Limpiar"].Value == true)
                {
                    filasLimpiar++;
                    tarimasLimpiar += row.Cells["tarima"].Value.ToString().Trim() + ", ";
                }
                else if (row.Cells["ASIGNAR"].Value != null && (bool)row.Cells["ASIGNAR"].Value == true)
                {
                    filasAsignar++;
                    tarimasAsignar += row.Cells["tarima"].Value.ToString().Trim() + ", ";
                }
            }

            if (filasLimpiar == 0 && filasAsignar == 0)
            {
                MessageBox.Show("No ha seleccionado ninguna tarima para procesar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string mensaje = "¿Está seguro de guardar los cambios?\n";
            if (filasAsignar > 0)
                mensaje += $"• Asignar fecha {DtFC.Value.ToString("dd/MMM/yy")} a {filasAsignar} tarima(s).\n";
            if (filasLimpiar > 0)
                mensaje += $"• Limpiar fecha de {filasLimpiar} tarima(s).";

            if (MessageBox.Show(mensaje, "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            thisConnection.Open();

            // Procesar filas con "Limpiar" (se asigna NULL)
            if (filasLimpiar > 0)
            {
                foreach (DataGridViewRow row in DgDatos.Rows)
                {
                    if ((bool)row.Cells["Limpiar"].Value == true)
                    {
                        string tarima = row.Cells["tarima"].Value.ToString().Trim();
                        string sql = "UPDATE tb_det_trazabilidad SET fecha_cad = NULL " +
                                     "WHERE recibo = @recibo AND prod_clave = '03001ML09' AND tarima = @tarima";
                        using (SqlCommand cmd = new SqlCommand(sql, thisConnection))
                        {
                            cmd.Parameters.AddWithValue("@recibo", TxtRecibo.Text);
                            cmd.Parameters.AddWithValue("@tarima", tarima);
                            cmd.ExecuteNonQuery();
                        }

                        // Opcional: también limpiar en tb_det_recepcion_pt (pero solo si todas las tarimas se limpian, o se hace una vez)
                        // Como es lógica de negocio, prefiero dejarlo igual que en el código original: se actualiza siempre.
                        // Pero si se limpia al menos una, se puede poner NULL en la recepción.
                    }
                }

                // Actualizar tb_det_recepcion_pt con NULL si al menos una tarima se limpió
                string sqlRecepcion = "UPDATE tb_det_recepcion_pt SET rptd_fechacad = NULL " +
                                      "WHERE rpt_recibo = @recibo AND prod_clave = '03001ML09'";
                using (SqlCommand cmd = new SqlCommand(sqlRecepcion, thisConnection))
                {
                    cmd.Parameters.AddWithValue("@recibo", TxtRecibo.Text);
                    cmd.ExecuteNonQuery();
                }
            }

            // Procesar filas con "ASIGNAR" (se asigna la fecha del picker)
            if (filasAsignar > 0)
            {
                foreach (DataGridViewRow row in DgDatos.Rows)
                {
                    if ((bool)row.Cells["ASIGNAR"].Value == true &&
                        (row.Cells["Limpiar"].Value == null || (bool)row.Cells["Limpiar"].Value == false))
                    {
                        string tarima = row.Cells["tarima"].Value.ToString().Trim();
                        string sql = "UPDATE tb_det_trazabilidad SET fecha_cad = @fecha " +
                                     "WHERE recibo = @recibo AND prod_clave = '03001ML09' AND tarima = @tarima";
                        using (SqlCommand cmd = new SqlCommand(sql, thisConnection))
                        {
                            cmd.Parameters.AddWithValue("@fecha", DtFC.Value.ToString("dd/MM/yyyy"));
                            cmd.Parameters.AddWithValue("@recibo", TxtRecibo.Text);
                            cmd.Parameters.AddWithValue("@tarima", tarima);
                            cmd.ExecuteNonQuery();
                        }
                    }
                }

                // Actualizar tb_det_recepcion_pt con la fecha elegida (si hay al menos una asignación)
                string sqlRecepcionAsignar = "UPDATE tb_det_recepcion_pt SET rptd_fechacad = @fecha " +
                                             "WHERE rpt_recibo = @recibo AND prod_clave = '03001ML09'";
                using (SqlCommand cmd = new SqlCommand(sqlRecepcionAsignar, thisConnection))
                {
                    cmd.Parameters.AddWithValue("@fecha", DtFC.Value.ToString("dd/MM/yyyy"));
                    cmd.Parameters.AddWithValue("@recibo", TxtRecibo.Text);
                    cmd.ExecuteNonQuery();
                }
            }

            thisConnection.Close();

            // Mensaje final y registro
            string detalle = "";
            if (filasLimpiar > 0) detalle += $"Limpió {filasLimpiar} tarima(s): {tarimasLimpiar.TrimEnd(',', ' ')}. ";
            if (filasAsignar > 0) detalle += $"Asignó fecha {DtFC.Value.ToString("dd/MM/yyyy")} a {filasAsignar} tarima(s): {tarimasAsignar.TrimEnd(',', ' ')}.";

            MessageBox.Show($"Operación completada.\n{detalle}", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

            Utilerias.Class1.registrar_movimiento(
                DateTime.Now,
                Environment.MachineName,
                Utilerias.Class1.Usu_login,
                "M",
                "2.3",
                TxtRecibo.Text,
                detalle,
                "SIPGAB"
            );

            // Limpiar y deshabilitar controles
            DtFC.Enabled = false;
            BtnSave.Enabled = false;
            DgDatos.Rows.Clear();
            LblFeEla.Visible = false;
        }

        // Evento para que el checkbox se marque/desmarque con un clic
        private void DgDatos_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                // Verificar si se hizo clic en alguna de las columnas checkbox
                if (DgDatos.Columns[e.ColumnIndex] is DataGridViewCheckBoxColumn)
                {
                    // Cambiar el valor del checkbox
                    DataGridViewCheckBoxCell cell = DgDatos.Rows[e.RowIndex].Cells[e.ColumnIndex] as DataGridViewCheckBoxCell;
                    if (cell != null)
                    {
                        bool current = Convert.ToBoolean(cell.Value);
                        cell.Value = !current;
                        DgDatos.EndEdit();  // Forzar la actualización
                    }
                }
            }
        }
    }
}