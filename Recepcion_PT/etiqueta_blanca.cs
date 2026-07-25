using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Gma.QrCodeNet.Encoding;
using Gma.QrCodeNet.Encoding.Windows.Render;
using System.Diagnostics;
using System.IO;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Data.SqlClient;
using System.Globalization;
using System.Runtime.InteropServices;

namespace Recepcion_PT
{    
    public partial class etiqueta_blanca : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlCommand cmnd1 = new SqlCommand();
        public string rancho = "", mProducto = "";
        //SqlDataReader reader1;

        public DataTable trazabilidad = new DataTable();
        public string recibo;


        public etiqueta_blanca()
        {
            InitializeComponent();
            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
        }

        private void etiqueta_blanca_Load(object sender, EventArgs e)
        {
            //recibo = "C:/Reportes/" + recibo + "_1.txt";
            //recibo = recibo + "_1.txt";            
            //trazabilidad.Columns.Add("pti_clave");
            //trazabilidad.Columns.Add("prod_nombre_ingles");
            //trazabilidad.Columns.Add("gtin_clave");
            //trazabilidad.Columns.Add("prod_clave");
            //trazabilidad.Columns.Add("prod_nombre");
            //trazabilidad.Columns.Add("etiqueta");
            //trazabilidad.Columns.Add("fecha_cad");
            //trazabilidad.Columns.Add("tarima");
            //trazabilidad.Columns.Add("recibo");
        }
        Image newImage1 = null;
        private void btnImp_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("<Aceptar> para imprimir etiqueta 1", "AVISO", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                if (etiqueta_verde.impr.Length > 0)
                {
                    //MessageBox.Show("etiqueta_blanca");
                    //using (StreamWriter outputFile = new StreamWriter(@"C:\reportes\288073.txt"))
                    //{
                    //    outputFile.WriteLine(etiqueta_verde.impr);
                    //}
                    RawPrinterHelper.SendStringToPrinter("etiqueta_blanca", etiqueta_verde.impr);
                    //RawPrinterHelper.SendStringToPrinter("\\\\192.168.123.124\\etiqueta_blanca", etiqueta_verde.impr);                    
                    //RawPrinterHelper.SendStringToPrinter("ZDesigner ZT410-203dpi ZPL", etiqueta_verde.impr);
                }
                if (etiqueta_verde.impr2.Length > 0 || etiqueta_verde.imp100x50.Length > 0)
                {   
                    //4621468107                 
                    //MessageBox.Show("impresora_papel0");
                    //RawPrinterHelper.SendStringToPrinter("ZDesigner GX420t", etiqueta_verde.impr2);
                    if (etiqueta_verde.imp100x50.Length > 0)
                        if (Utilerias.Class1.Usu_login.Trim() == "N")
                            RawPrinterHelper.SendStringToPrinter("etiqueta_blanca", etiqueta_verde.imp100x50);
                        else
                        {
                            RawPrinterHelper.SendStringToPrinter("impresora_papel0", etiqueta_verde.imp100x50);
                            //MessageBox.Show("impresora_papel0");
                            //RawPrinterHelper.SendStringToPrinter("etiqueta_blanca", etiqueta_verde.imp100x50);                    
                        }
                    else
                        if (Utilerias.Class1.Usu_login.Trim() == "N")
                            RawPrinterHelper.SendStringToPrinter("etiqueta_blanca", etiqueta_verde.impr2);
                        else 
                            RawPrinterHelper.SendStringToPrinter("impresora_papel0", etiqueta_verde.impr2);
                    //RawPrinterHelper.SendStringToPrinter("etiqueta_blanca", etiqueta_verde.imp100x50);
                         //RawPrinterHelper.SendStringToPrinter("\\\\192.168.123.124\\etiqueta_blanca", etiqueta_verde.imp100x50);                    
                    //RawPrinterHelper.SendStringToPrinter("ZDesigner ZT410-203dpi ZPL", etiqueta_verde.impr2);
                    
                    /*for (int i = 0; i < trazabilidad.Rows.Count; i++)
                    {
                        if (Convert.ToString(trazabilidad.Rows[i]["pais"].ToString().Trim()) != "")
                        {
                            aux1 = Convert.ToString(trazabilidad.Rows[i]["pti_clave"].ToString()).Trim(); //pti_clave
                            aux2 = Convert.ToString(trazabilidad.Rows[i]["prod_nom_ingles"].ToString());//prod_nombre_ingles
                            aux3 = Convert.ToString(trazabilidad.Rows[i]["gtin_clave"].ToString());//gtin_clave
                            clave_gab = Convert.ToString(trazabilidad.Rows[i]["prod_clave"].ToString());//prod_clave
                            aux5 = Convert.ToString(trazabilidad.Rows[i]["prod_nombre"].ToString());//prod_nombre
                            aux5 = aux5.Replace("''", "");
                            aux5 = aux5.Replace("'", "");
                            auxn = Convert.ToDecimal(trazabilidad.Rows[i]["etiqueta"].ToString());//etiqueta                            
                            fecad = Convert.ToString(trazabilidad.Rows[i]["fecha_cad"].ToString());//fecha de caducidad 
                            pais = Convert.ToString(trazabilidad.Rows[i]["pais"].ToString());//pais origen del producto
                            auxtr_2 = clave_gab;//prod_clave          
                            if (i == 0)
                            {
                                BarcodeLib.Barcode.Linear code128 = new BarcodeLib.Barcode.Linear();
                                code128.Type = BarcodeLib.Barcode.BarcodeType.CODE128;
                                code128.Data = aux3;
                                code128.AddCheckSum = true;
                                code128.UOM = BarcodeLib.Barcode.UnitOfMeasure.PIXEL;
                                code128.BarHeight = 50;
                                code128.BarWidth = 2;
                                code128.drawBarcode("C:/sisgabweb/code128.jpg");
                                newImage1 = Image.FromFile(@"C:/sisgabweb/code128.jpg");
                                //byte[] code821 = code128.drawBarcodeAsBytes();
                                //ms1 = new MemoryStream(code821);
                            }
                            else
                            {
                                //si cambia el producto
                                if (clave_gab.Trim() != trazabilidad.Rows[i - 1]["prod_clave"].ToString().Trim())
                                {
                                    BarcodeLib.Barcode.Linear code128 = new BarcodeLib.Barcode.Linear();
                                    code128.Type = BarcodeLib.Barcode.BarcodeType.CODE128;
                                    code128.Data = aux3;
                                    code128.AddCheckSum = true;
                                    code128.UOM = BarcodeLib.Barcode.UnitOfMeasure.PIXEL;
                                    code128.BarHeight = 50;
                                    code128.BarWidth = 2;
                                    byte[] code821 = code128.drawBarcodeAsBytes();
                                    ms1 = new MemoryStream(code821);
                                }
                            }

                            string lin = auxtr_2.Substring(0, 2);//saber la linea
                            if ((auxtr_2.Trim() == "05003ML3P") || (auxtr_2.Trim() == "05005ML2P") || (auxtr_2.Trim() == "19"))
                            {
                                dia = fecad.Substring(0, 2);
                                fecad = fecad.Substring(3, 2);
                                obtenerNombreMesNumero(Convert.ToInt32(fecad));
                                mfeccad = "FC" + mes.ToUpper() + dia;
                            }

                            j = 1;
                            co = 0; renglon_selec = 0;
                            renglon_selec = Convert.ToInt32(auxn);
                            PaperSize ps;
                            PrintPreviewDialog pr = new PrintPreviewDialog();
                            pr.Document = eti_blanca;
                            //ps = new PaperSize("etiqueta_blanca", 425, 267);                           
                            ps = new PaperSize("etiqueta_blanca", 267, 425);
                            eti_blanca.DefaultPageSettings.PaperSize = ps;                            
                            eti_blanca.DefaultPageSettings.Landscape = true;
                            //eti_blanca.DefaultPageSettings.PrinterSettings.PrinterName = "Foxit Reader PDF Printer";
                            
                            eti_blanca.DefaultPageSettings.PrinterSettings.PrinterName = "ZDesigner GX420t";
                            //eti_blanca.DefaultPageSettings.PrinterSettings.PrinterName = "ZDesigner Z4M 203DPI (Copiar 2)";
                            eti_blanca.Print();
                        }//si tiene país el producto
                    }//for*/
                }

                //string s = "^XA^LH30,30\n^FO20,10^ADN,90,50^AD^FDHello World^FS\n^XZ";
                // Allow the user to select a printer.
                //PrintDialog pd = new PrintDialog();
                //pd.PrinterSettings = new PrinterSettings();
                //if (DialogResult.OK == pd.ShowDialog(this))
                //{                                                               
                    // Send a printer-specific to the printer.
                    //RawPrinterHelper.SendStringToPrinter(pd.PrinterSettings.PrinterName, etiqueta_verde.impr);                    
                //}
                //for (int i = 0; i < trazabilidad.Rows.Count; i++)
                //{
                //    aux1 = Convert.ToString(trazabilidad.Rows[i]["pti_clave"].ToString()).Trim(); //pti_clave
                //    aux2 = Convert.ToString(trazabilidad.Rows[i]["prod_nom_ingles"].ToString());//prod_nombre_ingles
                //    aux3 = Convert.ToString(trazabilidad.Rows[i]["gtin_clave"].ToString());//gtin_clave
                //    clave_gab = Convert.ToString(trazabilidad.Rows[i]["prod_clave"].ToString());//prod_clave
                //    aux5 = Convert.ToString(trazabilidad.Rows[i]["prod_nombre"].ToString());//prod_nombre
                //    aux5 = aux5.Replace("''", "");
                //    aux5 = aux5.Replace("'", "");
                //    auxn = Convert.ToDecimal(trazabilidad.Rows[i]["etiqueta"].ToString());//etiqueta                            
                //    fecad = Convert.ToString(trazabilidad.Rows[i]["fecha_cad"].ToString());//fecha de caducidad  
                //    auxtr_2 = clave_gab;//prod_clave          
                //    if (i == 0)
                //    {
                //        BarcodeLib.Barcode.Linear code128 = new BarcodeLib.Barcode.Linear();
                //        code128.Type = BarcodeLib.Barcode.BarcodeType.CODE128;
                //        code128.Data = aux3;
                //        code128.AddCheckSum = true;
                //        code128.UOM = BarcodeLib.Barcode.UnitOfMeasure.PIXEL;
                //        code128.BarHeight = 50;
                //        code128.BarWidth = 2;
                //        byte[] code821 = code128.drawBarcodeAsBytes();
                //        ms1 = new MemoryStream(code821);
                //    }
                //    else
                //    {
                //        //si cambia el producto
                //        if (clave_gab.Trim() != trazabilidad.Rows[i - 1]["prod_clave"].ToString().Trim())
                //        {
                //            BarcodeLib.Barcode.Linear code128 = new BarcodeLib.Barcode.Linear();
                //            code128.Type = BarcodeLib.Barcode.BarcodeType.CODE128;
                //            code128.Data = aux3;
                //            code128.AddCheckSum = true;
                //            code128.UOM = BarcodeLib.Barcode.UnitOfMeasure.PIXEL;
                //            code128.BarHeight = 50;
                //            code128.BarWidth = 2;
                //            byte[] code821 = code128.drawBarcodeAsBytes();
                //            ms1 = new MemoryStream(code821);
                //        }
                //    }

                //    string lin = auxtr_2.Substring(0, 2);//saber la linea
                //    if ((auxtr_2.Trim() == "05003ML3P") || (auxtr_2.Trim() == "05005ML2P") || (auxtr_2.Trim() == "19"))
                //    {
                //        dia = fecad.Substring(0, 2);
                //        fecad = fecad.Substring(3, 2);
                //        obtenerNombreMesNumero(Convert.ToInt32(fecad));
                //        mfeccad = "FC" + mes.ToUpper() + dia;
                //    }

                //    j = 1;
                //    co = 0; renglon_selec = 0;
                //    renglon_selec = Convert.ToInt32(auxn);

                //    PaperSize ps;
                //    PrintPreviewDialog pr = new PrintPreviewDialog();
                //    pr.Document = eti_blanca;
                //    ps = new PaperSize("etiqueta_blanca", 400, 400);
                //    eti_blanca.DefaultPageSettings.PaperSize = ps;
                //    eti_blanca.Print();
                //    //string imp = @"C:\Reportes\" + recibo + "_1.txt /d:LPT2";
                //    ////string imp = @"C:\Reportes\" + recibo + "_1.txt /d:lpt1"; 
                //    //Process myProcess = new Process();                
                //    //ProcessStartInfo myProcessStartInfo = new ProcessStartInfo("print", imp);
                //    //myProcess.StartInfo.CreateNoWindow = true;
                //    //myProcessStartInfo.UseShellExecute = false;
                //    //myProcessStartInfo.RedirectStandardOutput = true;
                //    //myProcess.StartInfo = myProcessStartInfo;
                //    //myProcess.Start();                
                //    //myProcess.Close();                
                //}
            }          
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            //if (Recepcion_PT.opcion != 4)
            //{
            //    thisConnection.Open();
            //    cmnd1 = thisConnection.CreateCommand();
            //    cmnd1.CommandText = "update tb_mstr_recepcion_pt set fin_captura = '" + DateTime.Now.ToString() + "', ind = 'R' where rpt_recibo = '" + recibo + "'";
            //    reader1 = cmnd1.ExecuteReader();
            //    thisConnection.Close();
            //}
            this.Close();
        }

        int j = 1, co = 0, renglon_selec = 0;
        string aux1 = "", aux2 = "", aux3 = "", clave_gab = "", aux3_texto = "", aux5 = "", /*fecad = "",*/ mfeccad = "";
        string auxc = "", auxd = "", aux4 = "", auxnum1 = "", auxnum2 = "", auxd_1 = "", aux4_1 = "";//auxtr_2 = "";
        string nutar1 = "", totar1 = "", nutar2 = "", totar2 = "", /*dia = "",*/ mes = "", pais = "";
        decimal auxn = 0, tot_cajas = 0;
        int y = 0, lon_cad1 = 0, lon_cad2 = 0;

        //Stream ms1;

        SolidBrush color=new SolidBrush(Color.Black);
        Font fuente_encabezados=new Font("Arial", 13);// encabezados                        
        Font nota=new Font("Impact", 12);//nombre del producto
        Font cadu=new Font("Impact", 8);//fecha caducidad            
        Font box = new Font("Impact", 7);//caja


        private void eti_blanca_PrintPage(object sender, PrintPageEventArgs e)
        {
            if (aux3 == "")
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
                //if ((renglon_selec % 2) != 0)
                //{
                //    QrEncoder qrEncoder = new QrEncoder(ErrorCorrectionLevel.H);
                //    QrCode qrCode = new QrCode();
                //    qrEncoder.TryEncode("http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4, out qrCode); //qrcode etiqueta 1

                //    GraphicsRenderer renderer = new GraphicsRenderer(new FixedModuleSize(5, QuietZoneModules.Two), Brushes.Black, Brushes.White);
                //    using (MemoryStream ms = new MemoryStream())
                //    {
                //        renderer.WriteToStream(qrCode.Matrix, ImageFormat.Png, ms);
                //        var ima = new Bitmap(ms);
                //        var imag = new Bitmap(ima);

                //        Image newImage = Image.FromStream(ms);
                //        e.Graphics.DrawImage(newImage, 128, 3, 60, 60);//qrcode etiqueta 1
                //        ms.Close();
                //        ms.Dispose();
                //        ima.Dispose();
                //        imag.Dispose();
                //    }

                //    e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar1 + "/" + totar1, fuente_encabezados, color, 0, 5);// primera linea etiqueta 1                                
                //    if (aux2.Length > 25)
                //    {
                //        e.Graphics.DrawString(aux2.Substring(0, 24), nota, color, 0, 15);//nombre en ingles parte 1 etiqueta 1
                //        e.Graphics.DrawString(aux2.Substring(25), nota, color, 0, 28);//nombre en ingles parte 2 etiqueta 1                                    
                //    }
                //    else
                //    {
                //        e.Graphics.DrawString(aux2, nota, color, 0, 15);//nombre en ingles completo etiqueta 1                                    
                //    }
                //    if (aux5.Length > 25)
                //    {
                //        e.Graphics.DrawString(aux5.Substring(0, 24), nota, color, 0, 41);//nombre en español parte 1 etiqueta 1
                //        e.Graphics.DrawString(aux5.Substring(25), nota, color, 0, 54);//nombre en español parte 2 etiqueta 1                                    
                //    }
                //    else
                //    {
                //        e.Graphics.DrawString(aux5, nota, color, 0, 41);//nombre en español completo etiqueta 1                                    
                //    }

                //    //BarcodeLib.Barcode.Linear code128 = new BarcodeLib.Barcode.Linear();
                //    //code128.Type = BarcodeLib.Barcode.BarcodeType.CODE128;
                //    //code128.Data = aux3;
                //    //code128.AddCheckSum = true;
                //    //code128.UOM = BarcodeLib.Barcode.UnitOfMeasure.PIXEL;
                //    //code128.BarHeight = 50;
                //    //code128.BarWidth = 2;
                //    ////code128.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                //    ////code128.drawBarcode("C:/sisgabweb/code128.jpg");
                //    //byte[] code821 = code128.drawBarcodeAsBytes();
                //    //Stream ms1 = new MemoryStream(code821);
                //    ////Image newImage1 = Image.FromFile(@"C:/sisgabweb/code128.jpg");
                //    //Image newImage1 = Image.FromStream(ms1);
                //    //e.Graphics.DrawImage(newImage1, 0, 80, 150, 30);// codebar 128 etiqueta 1          
                //    //Image newImage1 = Image.FromStream(ms1);
                //    //e.Graphics.DrawImage(newImage1, 80, 80, 150, 50);// codebar 128 etiqueta 1  
                //    //e.Graphics.DrawImage(newImage1, 50, 120);// codebar 128 etiqueta 1  
                //    e.Graphics.DrawImage(newImage1, 30, 30);

                //    StringFormat drawFormat = new StringFormat();
                //    drawFormat.FormatFlags = StringFormatFlags.DirectionVertical;
                //    e.Graphics.DrawString(mfeccad, cadu, color, 170, 80, drawFormat);//fecha de caducidad etiqueta 1                                

                //    e.Graphics.DrawString("C: " + auxnum1, box, color, 150, 67);//número de caja etiqueta 1                                

                //    Pen whitepen = new Pen(Color.White, 5);
                //    int fin = 105;
                //    for (int m = 0; m < 10; m++)
                //    {
                //        Rectangle recta = new Rectangle(50, fin, 100, 10);
                //        e.Graphics.DrawRectangle(whitepen, recta);
                //        fin++;
                //    }
                //    e.Graphics.DrawString("(01)" + aux3, nota, color, 30, 105); //numero codigo de barras etiqueta 1
                //    j++;
                //}
                //else
                //{
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
                    //e.Graphics.DrawImage(newImage, 128, 3, 60, 60);//qrcode etiqueta 1
                    e.Graphics.DrawImage(newImage, 315, 5, 75, 75);//qrcode etiqueta 1
                    ms.Close();
                    imag.Dispose();
                }

                //QrEncoder qrEncoder2 = new QrEncoder(ErrorCorrectionLevel.H);
                //QrCode qrCode2 = new QrCode();
                //qrEncoder2.TryEncode("http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4_1, out qrCode); //qrcode etiqueta 2

                //GraphicsRenderer renderer2 = new GraphicsRenderer(new FixedModuleSize(5, QuietZoneModules.Two), Brushes.Black, Brushes.White);
                //using (MemoryStream ms2 = new MemoryStream())
                //{
                //    renderer2.WriteToStream(qrCode.Matrix, ImageFormat.Png, ms2);
                //    var ima2 = new Bitmap(ms2);
                //    var imag2 = new Bitmap(ima2);

                //    Image newImage2 = Image.FromStream(ms2);
                //    e.Graphics.DrawImage(newImage2, 333, 3, 60, 60);//qrcode etiqueta 2
                //    ms2.Close();
                //    imag2.Dispose();
                //}

                e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar1 + "/" + totar1, fuente_encabezados, color, 5, 5);// primera linea etiqueta 1
                //e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar2 + "/" + totar2, fuente_encabezados, color, 210, 5);// primera linea etiqueta 2
                //e.Graphics.DrawString(aux2, nota, color, 5, 30);//nombre en ingles
                //e.Graphics.DrawString(aux5, nota, color, 5, 50);//nombre en español

                if (aux2.Length > 35)
                {
                    e.Graphics.DrawString(aux2.Substring(0, 35), nota, color, 5, 30);//nombre en ingles parte 1 etiqueta 1
                    e.Graphics.DrawString(aux2.Substring(36), nota, color, 5, 45);//nombre en ingles parte 2 etiqueta 1

                    //e.Graphics.DrawString(aux2.Substring(0, 24), nota, color, 210, 15);//nombre en ingles parte 1 etiqueta 2
                    //e.Graphics.DrawString(aux2.Substring(25), nota, color, 210, 28);//nombre en ingles parte 2 etiqueta 2
                }
                else
                {
                    e.Graphics.DrawString(aux2, nota, color, 5, 30);//nombre en ingles completo etiqueta 1
                    //e.Graphics.DrawString(aux2, nota, color, 210, 15);//nombre en ingles completo etiqueta 2
                }
                if (aux5.Length > 35)
                {
                    e.Graphics.DrawString(aux5.Substring(0, 35), nota, color, 5, 60);//nombre en español parte 1 etiqueta 1
                    e.Graphics.DrawString(aux5.Substring(36), nota, color, 5, 75);//nombre en español parte 2 etiqueta 1

                    //e.Graphics.DrawString(aux5.Substring(0, 24), nota, color, 210, 41);//nombre en español parte 1 etiqueta 2
                    //e.Graphics.DrawString(aux5.Substring(25), nota, color, 210, 54);//nombre en español parte 2 etiqueta 2
                }
                else
                {
                    e.Graphics.DrawString(aux5, nota, color, 5, 60);//nombre en español completo etiqueta 1
                    //e.Graphics.DrawString(aux5, nota, color, 210, 41);//nombre en español completo etiqueta 2
                }
                //if (aux2.Length > 25)
                //{
                //    e.Graphics.DrawString(aux2.Substring(0, 24), nota, color, 5, 30);//nombre en ingles parte 1 etiqueta 1
                //    e.Graphics.DrawString(aux2.Substring(25), nota, color, 5, 45);//nombre en ingles parte 2 etiqueta 1

                //    //e.Graphics.DrawString(aux2.Substring(0, 24), nota, color, 210, 15);//nombre en ingles parte 1 etiqueta 2
                //    //e.Graphics.DrawString(aux2.Substring(25), nota, color, 210, 28);//nombre en ingles parte 2 etiqueta 2
                //}
                //else
                //{
                //    e.Graphics.DrawString(aux2, nota, color, 5, 30);//nombre en ingles completo etiqueta 1
                //    //e.Graphics.DrawString(aux2, nota, color, 210, 15);//nombre en ingles completo etiqueta 2
                //}
                //if (aux5.Length > 25)
                //{
                //    e.Graphics.DrawString(aux5.Substring(0, 24), nota, color, 5, 60);//nombre en español parte 1 etiqueta 1
                //    e.Graphics.DrawString(aux5.Substring(25), nota, color, 5, 75);//nombre en español parte 2 etiqueta 1

                //    //e.Graphics.DrawString(aux5.Substring(0, 24), nota, color, 210, 41);//nombre en español parte 1 etiqueta 2
                //    //e.Graphics.DrawString(aux5.Substring(25), nota, color, 210, 54);//nombre en español parte 2 etiqueta 2
                //}
                //else
                //{
                //    e.Graphics.DrawString(aux5, nota, color, 5, 60);//nombre en español completo etiqueta 1
                //    //e.Graphics.DrawString(aux5, nota, color, 210, 41);//nombre en español completo etiqueta 2
                //}

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
                //Image newImage1 = Image.FromStream(ms1);
                //e.Graphics.DrawImage(newImage1, 0, 80, 150, 30);
                e.Graphics.DrawImage(newImage1, 30, 115);
                //BarcodeLib.Barcode.Linear code1282 = new BarcodeLib.Barcode.Linear();
                //code1282.Type = BarcodeLib.Barcode.BarcodeType.CODE128;
                //code1282.Data = aux3;
                //code1282.AddCheckSum = true;
                //code1282.UOM = BarcodeLib.Barcode.UnitOMfeasure.PIXEL;
                //code1282.BarHeight = 50;
                //code1282.BarWidth = 2;
                //code1282.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                //code1282.drawBarcode("C:/sisgabweb/code1282.jpg");
                //Image newImage3 = Image.FromFile(@"C:/sisgabweb/code1282.jpg");
                //e.Graphics.DrawImage(newImage1, 210, 80, 150, 30);// codebar 128 etiqueta 2

                StringFormat drawFormat = new StringFormat();
                drawFormat.FormatFlags = StringFormatFlags.DirectionVertical;
                e.Graphics.DrawString(mfeccad, cadu, color, 170, 80, drawFormat);//fecha de caducidad etiqueta 1

                //StringFormat drawFormat2 = new StringFormat();
                //drawFormat2.FormatFlags = StringFormatFlags.DirectionVertical;
                //e.Graphics.DrawString(mfeccad, cadu, color, 370, 80, drawFormat2);//fecha de caducidad etiqueta 2

                e.Graphics.DrawString("PRODUCT OF " + pais, box, color, 5, 97);//país
                
                
                e.Graphics.DrawString("C: " + auxnum1, box, color, 350, 77);//número de caja etiqueta 1
                //e.Graphics.DrawString("C: " + auxnum2, box, color, 350, 67);//número de caja etiqueta 2

                Pen whitepen = new Pen(Color.White, 5);
                //int fin = 180;
                int fin = 170;
                for (int m = 0; m < 10; m++)
                {
                    //Rectangle recta = new Rectangle(50, fin, 180, 10);
                    Rectangle recta = new Rectangle(103, fin, 180, 10);
                    e.Graphics.DrawRectangle(whitepen, recta);
                    fin++;
                }
                e.Graphics.DrawString("(01)" + aux3, nota, color, 103, 170); //numero codigo de barras etiqueta 1
                //e.Graphics.DrawString("US #1", nota, color, 180, 220);
                e.Graphics.DrawString("HEB      US #1", nota, color, 315, 97);
                e.Graphics.DrawString("PACKED BY: COMERCIALIZADORA GAB, S.A. DE C.V. CARRETERA PANAMERICANA KM. 291 - 1", box, color, 5, 190);
                e.Graphics.DrawString("COL. LA FORTALEZA CORTAZAR, GTO. C.P. 38300 MEXICO R.F.C. CGA-960614-2C5", box, color, 5, 203);                

                //e.Graphics.DrawString("PACKED BY:", box, color, 5, 190);
                //e.Graphics.DrawString("COMERCIALIZADORA GAB, S.A. DE C.V.", box, color, 5, 203);
                //e.Graphics.DrawString("CARRETERA PANAMERICANA KM. 291 - 1", box, color, 5, 216);
                //e.Graphics.DrawString("COL. LA FORTALEZA CORTAZAR, GTO.", box, color, 5, 229);
                //e.Graphics.DrawString("C.P. 38300 MEXICO", box, color, 5, 242);
                //e.Graphics.DrawString("R.F.C. CGA-960614-2C5", box, color, 5, 255);

                //e.Graphics.DrawString("US #1", nota, color, 180, 220);
                //e.Graphics.DrawString("HEB", nota, color, 300, 220);
                //Pen black = new Pen(Color.Black, 5);
                //fin = 105;
                //for (int m = 0; m < 10; m++)
                //{
                //    Rectangle recta = new Rectangle(230, fin, 100, 10);
                //    e.Graphics.DrawRectangle(whitepen, recta);
                //    fin++;
                //}
                //e.Graphics.DrawString("(01)" + aux3, nota, color, 230, 105); //numero codigo de barras etiqueta 2  
                //j += 2;
                j++;
                //}
                //co = 1;
            }
            //else
            /*if (co == 1)
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
                    e.Graphics.DrawImage(newImage, 128, 3, 60, 60);//qrcode etiqueta 1
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
                    //e.Graphics.DrawImage(newImage2, 333, 3, 60, 60);//qrcode etiqueta 2
                    ms2.Close();
                    imag2.Dispose();
                }

                e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar1 + "/" + totar1, fuente_encabezados, color, 3, 5);// primera linea etiqueta 1
                e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar2 + "/" + totar2, fuente_encabezados, color, 210, 5);// primera linea etiqueta 2
                if (aux2.Length > 25)
                {
                    e.Graphics.DrawString(aux2.Substring(0, 24), nota, color, 0, 15);//nombre en ingles parte 1 etiqueta 1
                    e.Graphics.DrawString(aux2.Substring(25), nota, color, 0, 28);//nombre en ingles parte 2 etiqueta 1

                    e.Graphics.DrawString(aux2.Substring(0, 24), nota, color, 210, 15);//nombre en ingles parte 1 etiqueta 2
                    e.Graphics.DrawString(aux2.Substring(25), nota, color, 210, 28);//nombre en ingles parte 2 etiqueta 2
                }
                else
                {
                    e.Graphics.DrawString(aux2, nota, color, 0, 15);//nombre en ingles completo etiqueta 1
                    e.Graphics.DrawString(aux2, nota, color, 210, 15);//nombre en ingles completo etiqueta 2
                }
                if (aux5.Length > 25)
                {
                    e.Graphics.DrawString(aux5.Substring(0, 24), nota, color, 0, 41);//nombre en español parte 1 etiqueta 1
                    e.Graphics.DrawString(aux5.Substring(25), nota, color, 0, 54);//nombre en español parte 2 etiqueta 1

                    e.Graphics.DrawString(aux5.Substring(0, 24), nota, color, 210, 41);//nombre en español parte 1 etiqueta 2
                    e.Graphics.DrawString(aux5.Substring(25), nota, color, 210, 54);//nombre en español parte 2 etiqueta 2
                }
                else
                {
                    e.Graphics.DrawString(aux5, nota, color, 0, 41);//nombre en español completo etiqueta 1
                    e.Graphics.DrawString(aux5, nota, color, 210, 41);//nombre en español completo etiqueta 2
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
                //Image newImage1 = Image.FromStream(ms1);
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
                e.Graphics.DrawImage(newImage1, 210, 80, 150, 30);// codebar 128 etiqueta 2

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
            }*/

            if (co == 2)
            {
                co = 0;
                e.Graphics.DrawString("|   |   |   |   |   |   |   |   |   |", nota, color, 230, 50); //numero codigo de barras etiqueta 1
                //e.Graphics.DrawString("|   |   |   |   |   |   |   |   |   |", nota, color, 3, 50); //numero codigo de barras etiqueta 2  
            }

            if (j > renglon_selec)
            {
                if (co == 0)
                {
                    e.HasMorePages = false;
                    j = 0;
                    co = 2;
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
                co = 0;
            }
        }

        private string obtenerNombreMesNumero(int numeroMes)
        {
            try
            {
                DateTimeFormatInfo formatoFecha = CultureInfo.CurrentCulture.DateTimeFormat;
                string nombreMes = formatoFecha.GetMonthName(numeroMes);
                mes = nombreMes.Substring(0, 3);
                return mes;
            }
            catch
            {
                return "Desconocido";
            }
        }
    }

    public class RawPrinterHelper
    {
        // Structure and API declarions:
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public class DOCINFOA
        {
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDocName;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pOutputFile;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDataType;
        }
        [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

        [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool EndDocPrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool StartPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool EndPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);

        // SendBytesToPrinter()
        // When the function is given a printer name and an unmanaged array
        // of bytes, the function sends those bytes to the print queue.
        // Returns true on success, false on failure.
        public static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, Int32 dwCount)
        {
            Int32 dwError = 0, dwWritten = 0;
            IntPtr hPrinter = new IntPtr(0);
            DOCINFOA di = new DOCINFOA();
            bool bSuccess = false; // Assume failure unless you specifically succeed.

            di.pDocName = "My C#.NET RAW Document";
            di.pDataType = "RAW";

            // Open the printer.
            if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
            {
                // Start a document.
                if (StartDocPrinter(hPrinter, 1, di))
                {
                    // Start a page.
                    if (StartPagePrinter(hPrinter))
                    {
                        // Write your bytes.
                        bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                        EndPagePrinter(hPrinter);
                    }
                    EndDocPrinter(hPrinter);
                }
                ClosePrinter(hPrinter);
            }
            // If you did not succeed, GetLastError may give more information
            // about why not.
            if (bSuccess == false)
            {
                dwError = Marshal.GetLastWin32Error();
            }
            return bSuccess;
        }

        public static bool SendFileToPrinter(string szPrinterName, string szFileName)
        {
            // Open the file.
            FileStream fs = new FileStream(szFileName, FileMode.Open);
            // Create a BinaryReader on the file.
            BinaryReader br = new BinaryReader(fs);
            // Dim an array of bytes big enough to hold the file's contents.
            Byte[] bytes = new Byte[fs.Length];
            bool bSuccess = false;
            // Your unmanaged pointer.
            IntPtr pUnmanagedBytes = new IntPtr(0);
            int nLength;

            nLength = Convert.ToInt32(fs.Length);
            // Read the contents of the file into the array.
            bytes = br.ReadBytes(nLength);
            // Allocate some unmanaged memory for those bytes.
            pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);
            // Copy the managed byte array into the unmanaged array.
            Marshal.Copy(bytes, 0, pUnmanagedBytes, nLength);
            // Send the unmanaged bytes to the printer.
            bSuccess = SendBytesToPrinter(szPrinterName, pUnmanagedBytes, nLength);
            // Free the unmanaged memory that you allocated earlier.
            Marshal.FreeCoTaskMem(pUnmanagedBytes);
            return bSuccess;
        }
        public static bool SendStringToPrinter(string szPrinterName, string szString)
        {
            IntPtr pBytes;
            Int32 dwCount;
            // How many characters are in the string?
            dwCount = szString.Length;
            // Assume that the printer is expecting ANSI text, and then convert
            // the string to ANSI text.
            pBytes = Marshal.StringToCoTaskMemAnsi(szString);
            // Send the converted ANSI string to the printer.
            SendBytesToPrinter(szPrinterName, pBytes, dwCount);
            Marshal.FreeCoTaskMem(pBytes);
            return true;
        }
    }

}
