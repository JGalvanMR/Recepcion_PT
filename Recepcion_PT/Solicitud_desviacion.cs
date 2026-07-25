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
    public partial class Solicitud_desviacion : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlDataReader reader;
        SqlCommand cmnd = new SqlCommand();        
                
        int consecutivo = 0, anio = 0;
        public static string folio_no_conformidad = "", turno = "", fecha = "", producto = "", cantidad = "", especificado = "", mes =  "", month = "";

        public Solicitud_desviacion()
        {
            InitializeComponent();
            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
            txtds.Focus();
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

        private void Solicitud_desviacion_Load(object sender, EventArgs e)
        {
            txtnoconformidad.Text = folio_no_conformidad;
            cbturno.SelectedItem = turno;
            DTPFec.Value = Convert.ToDateTime(fecha);
            txtprod.Text = producto;
            NUDCanti.Value = Convert.ToInt32(cantidad);
            txtesp.Text = especificado;


            //cbturno.SelectedIndex = 0;
            month = obtenerNombreMesNumero(Convert.ToInt32(DateTime.Now.Month.ToString()));
            anio = Convert.ToInt32(Convert.ToString(DateTime.Now.Year.ToString()).Substring(2));
            thisConnection.Open();
            cmnd = thisConnection.CreateCommand();
            cmnd.CommandText = "select top 1 folio from tb_solicitud_desviacion where tipo = 'PT' and mes = '" + month + "' and anio = '" + anio + "' order by anio, mes";
            consecutivo = Convert.ToInt32(cmnd.ExecuteScalar()) + 1;
            thisConnection.Close();
            txtfolio.Text = month + "-" + anio + "-" + consecutivo;
        }

        private void btnguardar_Click(object sender, EventArgs e)
        {
            try 
            {
                if (txtesp.Text.Trim() == "")
                {
                    MessageBox.Show("Favor de capturar lo especificado", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtesp.Focus();
                    return;
                }
                if (txtds.Text.Trim() == "")
                {
                    MessageBox.Show("Favor de capturar la desviación solicitada", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtnc.Focus();
                    return;
                }
                if ((RBaceptada.Checked == false) && (RBrechazada.Checked == false))
                {
                    MessageBox.Show("Favor de seleccionar si la desviación es aceptada o rechazada", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    RBaceptada.Focus();
                    return;
                }
                thisConnection.Open();
                cmnd = thisConnection.CreateCommand();
                cmnd.CommandText = "insert into tb_solicitud_desviacion (folio, folio_no_conformidad, fecha, mes, anio, tipo) values ('" + consecutivo + "', " +
                                   "'" + txtnoconformidad.Text + "', '" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "', '" + month + "', '" + anio + "', 'PT')";
                reader = cmnd.ExecuteReader();
                reader.Dispose();
                thisConnection.Close();

                foreach (Process proceso in Process.GetProcesses())
                {
                    if (proceso.ProcessName == "winword")
                        proceso.Kill();
                }
                File.Copy("C:\\sisgabweb\\Solicitud_desviacion.docx", "C:\\Reportes\\Solicitud_desviacion.docx", true);
                object missing = Missing.Value;
                //  create Word application object
                Word.Application wordApp = new Word.Application();
                //  create Word document object
                Word.Document aDoc = null;
                //  create & define filename object with temp.doc
                object filename = @"C:\Reportes\Solicitud_desviacion.docx";
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
                    this.FindAndReplace(wordApp, "<Folio>", txtfolio.Text);
                    this.FindAndReplace(wordApp, "<Fecha>", DTPFec.Text);                    
                    this.FindAndReplace(wordApp, "<Turno>", cbturno.SelectedItem.ToString().Trim());
                    this.FindAndReplace(wordApp, "<Producto>", txtprod.Text);
                    this.FindAndReplace(wordApp, "<Folio_no_conformidad>", txtnoconformidad.Text);
                    this.FindAndReplace(wordApp, "<Cantidad>", NUDCanti.Value.ToString());
                    this.FindAndReplace(wordApp, "<Especificado>", txtesp.Text);
                    this.FindAndReplace(wordApp, "<No_conformidad>", txtnc.Text);
                    this.FindAndReplace(wordApp, "<Desviacion>", txtds.Text);
                    this.FindAndReplace(wordApp, "<Motivo>", txtmot.Text);                   
                    if(RBsi1.Checked == true)
                    {
                        this.FindAndReplace(wordApp, "<si>", "X");
                        this.FindAndReplace(wordApp, "<no>", " ");
                    }
                    if (RB1no.Checked == true)
                    {
                        this.FindAndReplace(wordApp, "<si>", " ");
                        this.FindAndReplace(wordApp, "<no>", "X");
                    }
                    this.FindAndReplace(wordApp, "<Representante>", txtreprecnte.Text);

                    if (RBaceptada.Checked == true)
                    {
                        this.FindAndReplace(wordApp, "<Ac>", "X");
                        this.FindAndReplace(wordApp, "<Re>", " ");
                    }
                    if (RBrechazado.Checked == true)
                    {
                        this.FindAndReplace(wordApp, "<Ac>", " ");
                        this.FindAndReplace(wordApp, "<Re>", "X");
                    }
                    this.FindAndReplace(wordApp, "<Fecha_cnte>", DTPcnte.Text);
                    this.FindAndReplace(wordApp, "<Evidencia>", txtevidencia.Text);
                    this.FindAndReplace(wordApp, "<Comentarios>", txtcom.Text);
                    if (RBaceptada.Checked == true)
                    {
                        this.FindAndReplace(wordApp, "<S>", "X");
                        this.FindAndReplace(wordApp, "<R>", " ");
                    }
                    if (RBrechazada.Checked == true)
                    {
                        this.FindAndReplace(wordApp, "<S>", " ");
                        this.FindAndReplace(wordApp, "<R>", "X");
                    }
                    //  save temp.doc after modified
                    aDoc.Save();
                    //foreach (Process proceso in Process.GetProcesses())
                    //{
                    //    if (proceso.ProcessName == "winword")
                    //        proceso.Kill();
                    //}
                    ProcessStartInfo proces = new ProcessStartInfo(@"C:\\reportes\Solicitud_desviacion.DOCX");
                    Process.Start(proces);                    
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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Internal Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void btnsalir_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        
    }
}
