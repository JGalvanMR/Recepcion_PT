using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Data.SqlClient;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;
using System.Diagnostics; 

namespace Recepcion_PT
{
    public partial class No_conformidad : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlDataReader reader;
        SqlCommand cmnd = new SqlCommand();

        public static string recibo = "", nombre_proveedor = "", mes = "", month = "", ruta_foto1 = "", ruta_foto2 = "", ruta_foto3 = "", prod = "", evaluacion = "", danos = "", 
                             fecha = "";
        public static DataTable productos = new DataTable();
        bool sel_todo = true;
        int consecutivo = 0, anio = 0;

        public No_conformidad()
        {
            InitializeComponent();
            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
            productos.Columns.Add("producto");
        }

        private void No_conformidad_Load(object sender, EventArgs e)
        {
            DTPFecNot.Value = Convert.ToDateTime(fecha);
            txtrecibo.Text = recibo;
            txtnomprov.Text = nombre_proveedor;
                      
            foreach (DataRow row in productos.Rows)
                ListProd.Items.Add(row["producto"].ToString().Trim());

            cbturno.SelectedIndex = 0;
            month = obtenerNombreMesNumero(Convert.ToInt32(DateTime.Now.Month.ToString()));
            anio = Convert.ToInt32(Convert.ToString(DateTime.Now.Year.ToString()).Substring(2));
            thisConnection.Open();
            cmnd = thisConnection.CreateCommand();
            cmnd.CommandText = "select count(folio) from tb_no_conformidad where tipo = 'PT' and mes = '" + month + "' and anio = '" + anio + "' "+
                               "group by anio, mes order by anio, mes";
            consecutivo = Convert.ToInt32(cmnd.ExecuteScalar()) + 1;
            thisConnection.Close();
            txtfolio.Text = month + "-" + anio + "-" + consecutivo;
            //txtenc.Text = danos;
        }

        private string obtenerNombreMesNumero(int numeroMes)
        {
            try
            {
                DateTimeFormatInfo formatoFecha = CultureInfo.CurrentCulture.DateTimeFormat;
                string nombreMes = formatoFecha.GetMonthName(numeroMes);
                mes = nombreMes.Substring(0, 3).ToUpper();
                return mes;
            }
            catch
            {
                return "Desconocido";
            }
        }

        private void btnsalir_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnselall_Click(object sender, EventArgs e)
        {           
            if (sel_todo == true)
            {
                for (int i = 0; i < ListProd.Items.Count; i++)
                    ListProd.SetItemChecked(i, true);

                sel_todo = false;
                return;
            }            
            if (sel_todo == false)
            {
                for (int i = 0; i < ListProd.Items.Count; i++)
                    ListProd.SetItemChecked(i, false);

                sel_todo = true;
                return;
            }
        }       
        private void btnfot1_Click(object sender, EventArgs e)
        {            
            Foto1 = new OpenFileDialog();
            Foto1.Filter = "JPG Files (*.jpg)|*.jpg|PNG Files (*.png)|*.png|JPEG Files (*.jpeg)|*.jpeg|GIF Files (*.gif)|*.gif";
            if (Foto1.ShowDialog() == DialogResult.OK)
            {
                ruta_foto1 = Foto1.FileName;
                pbfoto1.Image = System.Drawing.Bitmap.FromFile(ruta_foto1);
            }
        }

        private void btnfot2_Click(object sender, EventArgs e)
        {
            Foto2 = new OpenFileDialog();
            Foto2.Filter = "JPG Files (*.jpg)|*.jpg|PNG Files (*.png)|*.png|JPEG Files (*.jpeg)|*.jpeg|GIF Files (*.gif)|*.gif";
            if (Foto2.ShowDialog() == DialogResult.OK)
            {
                ruta_foto2 = Foto2.FileName;
                pbfoto2.Image = System.Drawing.Bitmap.FromFile(ruta_foto2);
            }
        }

        private void btnfot3_Click(object sender, EventArgs e)
        {
            Foto3 = new OpenFileDialog();
            Foto3.Filter = "JPG Files (*.jpg)|*.jpg|PNG Files (*.png)|*.png|JPEG Files (*.jpeg)|*.jpeg|GIF Files (*.gif)|*.gif";
            if (Foto3.ShowDialog() == DialogResult.OK)
            {
                ruta_foto3 = Foto3.FileName;
                pbfoto3.Image = System.Drawing.Bitmap.FromFile(ruta_foto3);
            }
        }

        private void btnguardar_Click(object sender, EventArgs e)
        {
            try
            {               
                for (int i = 0; i < ListProd.Items.Count; i++)
                {
                    if (ListProd.GetItemChecked(i))
                        prod = prod + ListProd.Items[i].ToString() + " ";

                }
              
                if (prod.Trim() == "")
                {
                    MessageBox.Show("Favor de seleccionar al menos un producto", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    ListProd.Focus();
                    return;
                }

                if (Convert.ToInt32(NUDCanti.Value) < 1)
                {
                    MessageBox.Show("Favor de elegir el número de cajas a devolver", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    NUDCanti.Focus();
                    return;
                }

                if (txtesp.Text.Trim() == "")
                {
                    MessageBox.Show("Favor de escribir en el campo de especificado", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtesp.Focus();
                    return;
                }

                if(txtenc.Text.Trim() == "")
                {
                    MessageBox.Show("Favor de escribir en el campo de encontrado", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtenc.Focus();
                    return;
                }

                if((ruta_foto1.Trim() == "") && (ruta_foto2.Trim() == "") && (ruta_foto3.Trim() == ""))
                {
                    MessageBox.Show("Favor de cargar al menos una foto", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    btnfot1.Focus();
                    return;
                }

                if (txtnumemp.Text.Trim() == "")
                {
                    MessageBox.Show("Favor de capturar lo del empaque", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtnumemp.Focus();
                    return;
                }

                if(txtcom.Text.Trim() == "")
                {
                    MessageBox.Show("Favor de escribir los comentarios", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtcom.Focus();
                    return;
                }

                thisConnection.Open();
                cmnd = thisConnection.CreateCommand();
                cmnd.CommandText = "insert into tb_no_conformidad (folio, fecha, recibo, mes, anio, tipo) values ('" + consecutivo + "', "+
                                   "'" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "', "+
                                   "'" + txtrecibo.Text + "', '" + month + "', '" + anio + "', 'PT')";
                reader = cmnd.ExecuteReader();
                reader.Dispose();


                cmnd = thisConnection.CreateCommand();
                cmnd.CommandText = "update tb_mstr_recepcion_pt set folio_no_conformidad = '" + txtfolio.Text + "' where rpt_recibo = '" + txtrecibo.Text + "'";
                reader = cmnd.ExecuteReader();
                reader.Dispose();
                thisConnection.Close();
                               
                foreach (Process proceso in Process.GetProcesses())
                {
                    if (proceso.ProcessName == "winword")
                        proceso.Kill();
                }
                File.Copy("C:\\sisgabweb\\PRODUCTO_NO_CONFORME.docx", "C:\\Reportes\\PRODUCTO_NO_CONFORME.docx", true);
                object missing = Missing.Value;
                //  create Word application object
                Word.Application wordApp = new Word.Application();
                //  create Word document object
                Word.Document aDoc = null;
                //  create & define filename object with temp.doc
                object filename = @"C:\Reportes\PRODUCTO_NO_CONFORME.docx";
                //  if temp.doc available
                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    //  make visible Word application
                    wordApp.Visible = false;
                    //  open Word document named temp.doc
                    aDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, 
                                                  ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
                    aDoc.Activate();
                    //  Call FindAndReplace()function for each change
                    this.FindAndReplace(wordApp, "<Date>", DTPFec.Text);
                    this.FindAndReplace(wordApp, "<Turno>", cbturno.SelectedItem.ToString().Trim());
                    this.FindAndReplace(wordApp, "<Recibo>", txtrecibo.Text);
                    this.FindAndReplace(wordApp, "<Folio>", txtfolio.Text);
                    this.FindAndReplace(wordApp, "<Producto>", prod);
                    this.FindAndReplace(wordApp, "<Cantidad>", NUDCanti.Value.ToString());
                    this.FindAndReplace(wordApp, "<Especificado>", txtesp.Text);
                    this.FindAndReplace(wordApp, "<Encontrado>", txtenc.Text);

                    if (ruta_foto1.Trim() != "")
                        aDoc.Shapes.AddPicture(ruta_foto1, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 10, 250, 150, 160);
                    if (ruta_foto2.Trim() != "")
                        aDoc.Shapes.AddPicture(ruta_foto2, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 160, 250, 150, 160);
                    if (ruta_foto3.Trim() != "")
                        aDoc.Shapes.AddPicture(ruta_foto3, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 310, 250, 150, 160);

                    this.FindAndReplace(wordApp, "<Empaque>", txtnumemp.Text);
                    this.FindAndReplace(wordApp, "<Comentarios>", txtcom.Text);
                    this.FindAndReplace(wordApp, "<Proveedor>", txtnomprov.Text);
                    this.FindAndReplace(wordApp, "<FechaNot>", DTPFecNot.Text);
                    this.FindAndReplace(wordApp, "<Contacto>", txtcontacto.Text);

                    //  save temp.doc after modified
                    aDoc.Save();
                    //foreach (Process proceso in Process.GetProcesses())
                    //{
                    //    if (proceso.ProcessName == "winword")
                    //        proceso.Kill();
                    //}
                    ProcessStartInfo proces = new ProcessStartInfo(@"C:\\reportes\PRODUCTO_NO_CONFORME.DOCX");
                    Process.Start(proces);

                    if (evaluacion.Trim() == "SOBRE INSPECCION")
                    {
                        Solicitud_desviacion.folio_no_conformidad = txtfolio.Text;
                        Solicitud_desviacion.turno = cbturno.SelectedItem.ToString().Trim();
                        Solicitud_desviacion.fecha = DTPFec.Value.ToShortDateString();
                        Solicitud_desviacion.producto = prod;
                        Solicitud_desviacion.cantidad = NUDCanti.Value.ToString();
                        Solicitud_desviacion.especificado = txtesp.Text;
                        Solicitud_desviacion sd = new Solicitud_desviacion();
                        sd.ShowDialog();
                    }
                }
                else
                {
                    MessageBox.Show("File does not exist.", "No File", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //foreach (Process proceso in Process.GetProcesses())
                    //{
                    //    if (proceso.ProcessName == "winword")
                    //        proceso.Kill();
                    //}
                }
                //killprocess("winword");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Internal Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                prod = "";
            }
        }

        private void FindAndReplace(Word.Application wordApp, object findText, object replaceText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            wordApp.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, 
                                           ref format, ref replaceText, ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }

        PictureBox picture = null;

        private void pbfoto1_Click(object sender, EventArgs e)
        {
            picture = pbfoto1;
            maximizar_foto();
        }

        private void pbfoto2_Click(object sender, EventArgs e)
        {
            picture = pbfoto2;
            maximizar_foto();
        }

        private void pbfoto3_Click(object sender, EventArgs e)
        {
            picture = pbfoto3;
            maximizar_foto();
        }

        public void maximizar_foto()
        {
            Imagen imagen = new Imagen();
            ((PictureBox)imagen.Controls["image"]).Image = this.picture.Image;
            imagen.ShowDialog();
        }
    }
}
