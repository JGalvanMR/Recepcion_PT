using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Globalization;
using System.Drawing.Printing;
using Gma.QrCodeNet.Encoding;
using Gma.QrCodeNet.Encoding.Windows.Render;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.Net;
using System.Net.Sockets;

namespace Recepcion_PT
{
    public partial class Imp_eti_tar : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlCommand cmnd = new SqlCommand();
        SqlDataAdapter da;
        //SqlDataReader reader;
        DataSet ds = new DataSet();
        DataTable recibo = new DataTable("recibo");
        ComboBox prod_clave = new ComboBox();

        string aux1 = "", aux2 = "", aux3 = "", clave_gab = "", aux3_texto = "", aux5 = "", fecad = "", mfeccad = "", pais = "";
        string auxc = "", auxd = "", aux4 = "", auxnum1 = "", auxnum2 = "", auxd_1 = "", aux4_1 = "";
        string nutar1 = "", totar1 = "", nutar2 = "", totar2 = "", dia = "", mes = "", impr = "", fila = "", grado = "", num = "", impr2 = "", date = "";

        private void Imp_eti_tar_Load(object sender, EventArgs e)
        {

        }

        decimal auxn = 0, tot_cajas = 0;
        int y = 0, lon_cad1 = 0, lon_cad2 = 0;

        public Imp_eti_tar()
        {
            InitializeComponent();
            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
        }

        private void txtrecibo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtrecibo.Text.Trim().Length > 0)
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    thisConnection.Open();
                    string query = "select * from tb_det_trazabilidad where recibo = '" + txtrecibo.Text + "' order by prod_clave, tarima";
                    da = new SqlDataAdapter(query, thisConnection);
                    da.Fill(ds, "recibo");
                    recibo = ds.Tables["recibo"];
                    thisConnection.Close();

                    if (recibo.Rows.Count > 0)
                    {
                        lbselprod.Visible = true;
                        cbselprod.Visible = true;
                        for (int i = 0; i < recibo.Rows.Count; i++)
                        {
                            if (i == 0)
                            {
                                cbselprod.Items.Add(Convert.ToString(recibo.Rows[i]["prod_clave"].ToString()) + " - " + Convert.ToString(recibo.Rows[i]["prod_nombre"].ToString()));
                                prod_clave.Items.Add(Convert.ToString(recibo.Rows[i]["prod_clave"].ToString()).Trim());
                            }
                            else
                            {
                                if (Convert.ToString(recibo.Rows[i]["prod_clave"].ToString().Trim()) != Convert.ToString(recibo.Rows[i - 1]["prod_clave"].ToString()).Trim())
                                {
                                    cbselprod.Items.Add(Convert.ToString(recibo.Rows[i]["prod_clave"].ToString()) + " - " + Convert.ToString(recibo.Rows[i]["prod_nombre"].ToString()));
                                    prod_clave.Items.Add(Convert.ToString(recibo.Rows[i]["prod_clave"].ToString()).Trim());
                                }
                            }
                        }
                    }
                    else
                    {
                        lbselprod.Visible = false;
                        cbselprod.Visible = false;
                        MessageBox.Show("Este recibo no tiene etiquetas generadas, favor de verificarlo", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txtrecibo.SelectAll();
                        txtrecibo.Focus();
                        cbselprod.Items.Clear();
                        prod_clave.Items.Clear();
                        return;
                    }
                }
            }
        }

        private void btnsaveout_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cbselprod_SelectionChangeCommitted(object sender, EventArgs e)
        {
            thisConnection.Open();
            prod_clave.SelectedIndex = cbselprod.SelectedIndex;
            DGV.Rows.Clear();
            foreach (DataRow row in recibo.Select("prod_clave = '" + prod_clave.SelectedItem.ToString().Trim() + "'"))
            {
                cmnd = thisConnection.CreateCommand();
                cmnd.CommandText = "select prod_paisorigen from tb_cat_producto where prod_clave = '" + prod_clave.SelectedItem.ToString().Trim() + "'";
                pais = Convert.ToString(cmnd.ExecuteScalar()).Trim();
                DGV.Rows.Add(Convert.ToString(row["prod_clave"].ToString().Trim()), Convert.ToString(row["prod_nombre"].ToString().Trim()), Convert.ToString(row["tarima"].ToString().Trim()), Convert.ToString(row["etiqueta"].ToString().Trim()), 0, 0, false, "", pais);
                //DGV.Rows.Add(Convert.ToString(row["prod_clave"].ToString().Trim()), Convert.ToString(row["prod_nombre"].ToString().Trim()), Convert.ToString(row["etiqueta"].ToString().Trim()), Convert.ToString(row["tarima"].ToString().Trim()), 0, 0, false, "", pais);
                aux3 = Convert.ToString(row["gtin_clave"].ToString().Trim());
            }
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            btnvalidareti.Enabled = true;
            thisConnection.Close();
            //thisConnection.Open();
            //cmnd = thisConnection.CreateCommand();
            //cmnd.CommandText = "select prod_paisorigen, prod_codegtin from tb_cat_producto where prod_clave = '" + prod_clave.SelectedItem.ToString().Trim() + "'";
            //reader = cmnd.ExecuteReader();
            //while (reader.Read())
            //{
            //    //pais_origen = Convert.ToString(cmnd.ExecuteScalar());
            //    pais_origen = reader.GetValue(0).ToString().Trim();
            //}

            //thisConnection.Close();

            //BarcodeLib.Barcode.Linear code128 = new BarcodeLib.Barcode.Linear();
            //code128.Type = BarcodeLib.Barcode.BarcodeType.CODE128;
            //code128.Data = aux3;
            //code128.AddCheckSum = true;
            //code128.UOM = BarcodeLib.Barcode.UnitOfMeasure.PIXEL;
            //code128.BarHeight = 50;
            //code128.BarWidth = 2;
            //code128.ImageFormat = System.Drawing.Imaging.ImageFormat.Png;
            ////code128.drawBarcode("C:/sisgabweb/code128.png");
            ////code128.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
            ////code128.drawBarcode("C:/sisgabweb/code128.jpg");
            //byte[] code821 = code128.drawBarcodeAsBytes();
            //ms1 = new MemoryStream(code821);            
            //Image newImage1 = Image.FromFile(@"C:/sisgabweb/code128.jpg");
            //Image newImage1 = Image.FromStream(ms1);
            //e.Graphics.DrawImage(newImage1, 0, 80, 150, 30);// codebar 128 etiqueta 1            

            //foreach(byte b in code821)
            //{
            //    hexString += b.ToString("x2");
            //}
            //fileSize = code821.Length;
            //string bitmapFilePath = @"C:/sisgabweb/code128.bmp";
            //byte[] bitmapFileData = System.IO.File.ReadAllBytes(bitmapFilePath);

            //impr = "^XA" + "^MNN" + "^LL500" + "~DYE:code128,P,P," + code821.Length + ",," + zplImageData + "^XZ\n";
            //impr = "~DYE:code128,P,P," + code821.Length + ",," + zplImageData + "\n";
            //printImage = "^XA^FO115,50^IME:code128.PNG^FS^XZ";

            //fileSize = bitmapFileData.Length;
            //bitmapDataOffset = 62;
            //width = 255;
            //height = 255;
            //bistPerPixel = 1;
            //bitmapDataLength = 8160;
            //widthInBytes = Math.Ceiling(width / 8.0);
            //byte[] bitmap = new byte[bitmapDataLength];
            //Buffer.BlockCopy(bitmapFileData, bitmapDataOffset, bitmap, 0, bitmapDataLength);
            //for (int i = 0; i < bitmapDataLength; i++)
            //{
            //    bitmap[i] ^= 0xFF;
            //}

            //string ZPLImageDataString = BitConverter.ToString(bitmap);
            //ZPLImageDataString = ZPLImageDataString.Replace("-", string.Empty);            
        }
        //Stream ms1;
        Image newImage1 = null;
        private void btnvalidareti_Click(object sender, EventArgs e)
        {
            for (int a = 0; a < DGV.Rows.Count; a++)
            {
                if (Convert.ToString(DGV.Rows[a].Cells["imp"].Value).Trim() == "True")
                {
                    if (Convert.ToInt32(DGV.Rows[a].Cells["num_eti"].Value) == 0)
                    {
                        MessageBox.Show("El número de etiquetas a imprimir es igual a 0", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        DGV.Focus();
                        DGV.Rows[a].Cells["num_eti"].Selected = true;
                        return;
                    }

                    if (Convert.ToInt32(DGV.Rows[a].Cells["eti_fin"].Value) > Convert.ToInt32(DGV.Rows[a].Cells["eti"].Value))
                    {
                        MessageBox.Show("El número de etiquetas a imprimir es mayor al capturado en la tarima", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        DGV.Focus();
                        DGV.Rows[a].Cells["eti_fin"].Selected = true;
                        return;
                    }


                    if (Convert.ToInt32(DGV.Rows[a].Cells["num_eti"].Value) > Convert.ToInt32(DGV.Rows[a].Cells["eti_fin"].Value))
                    {
                        MessageBox.Show("El rango inicial es mayor al rango final", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        DGV.Focus();
                        DGV.Rows[a].Cells["eti_fin"].Selected = true;
                        return;
                    }

                    if ((Convert.ToString(DGV.Rows[a].Cells["clave_prod"].Value).Trim() == "18007JI56V") ||
                        (Convert.ToString(DGV.Rows[a].Cells["clave_prod"].Value).Trim() == "18JIBOML66") ||
                        (Convert.ToString(DGV.Rows[a].Cells["clave_prod"].Value).Trim() == "16001TO561") ||
                        (Convert.ToString(DGV.Rows[a].Cells["clave_prod"].Value).Trim() == "18007JI55M") ||
                        (Convert.ToString(DGV.Rows[a].Cells["clave_prod"].Value).Trim() == "18007JVM55") ||
                        (Convert.ToString(DGV.Rows[a].Cells["clave_prod"].Value).Trim() == "18007JVM66"))
                    {
                        if (cbgrado.SelectedIndex == -1)
                        {
                            MessageBox.Show("favor de seleccionar el tipo de grado", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            cbgrado.Visible = true;
                            lbgrado.Visible = true;
                            cbgrado.Focus();
                            fila = Convert.ToString(DGV.Rows[a].Cells["clave_prod"].Value);
                            //DGV5.Rows.Clear();
                            return;
                        }
                    }

                    #region sin pais
                    if (Convert.ToString(DGV.Rows[a].Cells["country"].Value).Trim() == "")
                    {
                        if ((Convert.ToInt32(DGV.Rows[a].Cells["eti"].Value) >= Convert.ToInt32(DGV.Rows[a].Cells["num_eti"].Value)))
                        {
                            foreach (DataRow row in recibo.Select("prod_clave = '" + Convert.ToString(DGV.Rows[a].Cells["clave_prod"].Value) + "' and tarima = '" + Convert.ToString(DGV.Rows[a].Cells["num_tar"].Value) + "'"))
                            {
                                aux1 = Convert.ToString(row["pti_clave"].ToString()).Trim(); //pti_clave
                                aux2 = Convert.ToString(row["prod_nom_ingles"].ToString());//prod_nombre_ingles
                                aux3 = Convert.ToString(row["gtin_clave"].ToString());//gtin_clave
                                clave_gab = Convert.ToString(row["prod_clave"].ToString());//prod_clave
                                aux5 = Convert.ToString(row["prod_nombre"].ToString());//prod_nombre
                                aux5 = aux5.Replace("''", "");
                                aux5 = aux5.Replace("'", "");
                                auxn = Convert.ToDecimal(row["etiqueta"].ToString());//etiqueta                            
                                fecad = Convert.ToString(row["fecha_cad"].ToString());
                                date = Convert.ToString(row["pti_fecha"].ToString());
                                string lin = clave_gab.Substring(0, 2);//saber la linea
                                if ((clave_gab == "05003ML3P") || (clave_gab == "05005ML2P") || (lin == "19"))
                                {
                                    dia = fecad.Substring(0, 2);
                                    fecad = fecad.Substring(3, 2);
                                    obtenerNombreMesNumero(Convert.ToInt32(fecad));
                                    mfeccad = "FC" + mes.ToUpper() + dia;
                                }
                                if (aux3 == "")
                                {
                                    aux3 = clave_gab;
                                    aux3_texto = "(01)" + aux3;
                                }
                                else
                                    aux3_texto = "(01)" + aux3;

                                tot_cajas = tot_cajas + auxn;

                                for (j = Convert.ToInt32(DGV.Rows[a].Cells["num_eti"].Value); j <= Convert.ToInt32(DGV.Rows[a].Cells["eti_fin"].Value); j++)
                                {
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
                                    if (j <= Convert.ToInt32(DGV.Rows[a].Cells["eti_fin"].Value))
                                    {
                                        if ((Convert.ToInt32(DGV.Rows[a].Cells["eti_fin"].Value)) % 2 == 0)
                                        {
                                            //impr = impr + "^MTT\n";// Sets the type to thermal transfer
                                            //impr = impr + "^XA\n^LH5,5\n";
                                            //impr = impr + "^BY2" + "^MNM\n";
                                            impr = impr + "^XA\n";
                                            impr = impr + "^FO270,25^BQN,2,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n"; //CODIGO 2D BQ DE CUADRO - PTI"
                                            impr = impr + "^FO700,25^BQN,2,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4_1 + "^FS\n";  // CODIGO 2D BQ DE CUADRO - PTI"
                                            impr = impr + "^FO-15,15^A2N,22,8,^FD" + txtrecibo.Text + "-" + clave_gab + "-" + nutar1 + "/" + totar1 + "^FS\n";//etiqueta 1
                                            impr = impr + "^FO450,15^A2N,22,8,^FD" + txtrecibo.Text + "-" + clave_gab + "-" + nutar2 + "/" + totar2 + "^FS\n";//etiqueta 2
                                            //nombre en ingles etiqueta 1
                                            if (aux2.Length > 25)
                                            {
                                                impr = impr + "^FO-15,55^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                impr = impr + "^FO-15,80^A0N,22,18^FD" + aux2.Substring(26, 24) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                            }
                                            else
                                                impr = impr + "^FO-15,55^A0N,22,18^FD" + aux2 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1

                                            //nombre en español etiqueta 1
                                            if (aux5.Length > 25)
                                            {
                                                impr = impr + "^FO-15,120^A0N,22,18^FD" + aux5.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                impr = impr + "^FO-15,143^A0N,22,18^FD" + aux5.Substring(26, 19) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                            }
                                            else
                                                impr = impr + "^FO-15,120^A0N,22,18^FD" + aux5 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1


                                            //nombre en ingles etiqueta 2
                                            if (aux2.Length > 25)
                                            {
                                                impr = impr + "^FO450,55^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2
                                                impr = impr + "^FO450,80^A0N,22,18^FD" + aux2.Substring(26, 24) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2
                                            }
                                            else
                                                impr = impr + "^FO450,55^A0N,22,18^FD" + aux2 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2

                                            //nombre en español etiqueta 2
                                            if (aux5.Length > 25)
                                            {
                                                impr = impr + "^FO450,120^A0N,22,18^FD" + aux5.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2
                                                impr = impr + "^FO450,143^A0N,22,18^FD" + aux5.Substring(26, 19) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2
                                            }
                                            else
                                                impr = impr + "^FO450,120^A0N,22,18^FD" + aux5 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 2


                                            if ((clave_gab.Trim() == "18007JI56V") || (clave_gab.Trim() == "18JIBOML66") || (clave_gab.Trim() == "16001TO561") ||
                                               (clave_gab.Trim() == "18007JI55M") || (clave_gab == "18007JVM55") || (clave_gab == "18007JVM66"))
                                            {
                                                //impr = impr + "^FO-15,185^A2B,30,25,^FD" + grado + "^FS\n"; // SE IMPRIME TIPO GRADO 1
                                                //impr = impr + "^FO430,185^A2B,30,25,^FD" + grado + "^FS\n"; // SE IMPRIME TIPO GRADO 2
                                                impr = impr + "^FO-15,175^A0B,17,18,^FD" + grado + "^FS\n"; // SE IMPRIME TIPO GRADO 1
                                                impr = impr + "^FO30,195^A1B,25,20,^FD" + num + "^FS\n"; // SE IMPRIME TIPO GRADO 1
                                                impr = impr + "^FO430,175^A0B,17,18,^FD" + grado + "^FS\n"; // SE IMPRIME TIPO GRADO 2
                                                impr = impr + "^FO450,195^A1B,25,20,^FD" + num + "^FS\n"; // SE IMPRIME TIPO GRADO 1
                                            }

                                            if (clave_gab.Trim() == "05006MLNA2")
                                            {
                                                impr = impr + "^FO-15,175^A0B,17,18,^FD" + Convert.ToDateTime(date).ToString("MM/dd/yy") + "^FS\n"; // fecha etiqueta 1                                                
                                                impr = impr + "^FO430,175^A0B,17,18,^FD" + Convert.ToDateTime(date).ToString("MM/dd/yy") + "^FS\n"; // fecha etiqueta 2                                                 
                                            }

                                            impr = impr + "^FO320,140^A0N,22,18^FD" + "C: " + auxnum1 + "^FS\n"; // CAJA ETIQUETA 1
                                            impr = impr + "^FO760,140^A0N,22,18^FD" + "C: " + auxnum2 + "^FS\n"; // CAJA ETIQUETA 2                                                                                                                                             
                                            impr = impr + "^FO90,170,^BY1,^BCN,60,N,N,N^MD5F^FD" + "01" + aux3 + "^FS\n"; // CODIGO GTIN COD DE BARRAS 128 1
                                            impr = impr + "^FO515,170,^BY1,^BCN,60,N,N,N^FD" + "01" + aux3 + "^FS\n"; // CODIGO GTIN COD DE BARRAS 128 2
                                            impr = impr + "^FO110,240^A0N,17,15,^FD" + aux3_texto + "^FS\n"; // DESCRIPCION DEL PRODUCTO EN EL CODIGO BQ LINEA 1                            
                                            impr = impr + "^FO350,165^A0B,17,18,^FD" + mfeccad + "^FS\n"; // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 1
                                            impr = impr + "^FO780,165^A0B,17,18,^FD" + mfeccad + "^FS\n"; // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 2    
                                            impr = impr + "^FO530,240^A0N,17,17,^FD" + aux3_texto + "^FS\n"; // DESCRIPCION DEL PRODUCTO EN EL CODIGO BQ LINEA 2
                                            impr = impr + "^XZ\n";
                                        }
                                        else
                                        {
                                            impr = impr + "^XA\n";
                                            impr = impr + "^FO270,25^BQN,2,3^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n"; //CODIGO 2D BQ DE CUADRO - PTI"
                                            impr = impr + "^FO-15,15^A2N,22,8,^FD" + txtrecibo.Text + "-" + clave_gab + "-" + nutar1 + "/" + totar1 + "^FS\n";//etiqueta 1                                    

                                            //nombre en ingles etiqueta 1
                                            if (aux2.Length > 25)
                                            {
                                                impr = impr + "^FO-15,55^A0N,22,18^FD" + aux2.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                impr = impr + "^FO-15,80^A0N,22,18^FD" + aux2.Substring(26, 24) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                            }
                                            else
                                                impr = impr + "^FO-15,55^A0N,22,18^FD" + aux2 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1

                                            //nombre en español etiqueta 1
                                            if (aux5.Length > 25)
                                            {
                                                impr = impr + "^FO-15,120^A0N,22,18^FD" + aux5.Substring(0, 25) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                                impr = impr + "^FO-15,143^A0N,22,18^FD" + aux5.Substring(26, 19) + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1
                                            }
                                            else
                                                impr = impr + "^FO-15,120^A0N,22,18^FD" + aux5 + "^FS\n"; // DESCRIPCION DEL PRODUCTO LINEA 1

                                            if ((clave_gab.Trim() == "18007JI56V") || (clave_gab.Trim() == "18JIBOML66") || (clave_gab.Trim() == "16001TO561") ||
                                                (clave_gab.Trim() == "18007JI55M") || (clave_gab == "18007JVM55") || (clave_gab == "18007JVM66"))
                                            {
                                                //impr = impr + "^FO-15,185^A0B,30,25,^FD" + grado + "^FS\n"; // SE IMPRIME TIPO GRADO 1                                            
                                                impr = impr + "^FO-15,175^A0B,17,18,^FD" + grado + "^FS\n"; // SE IMPRIME TIPO GRADO 1
                                                impr = impr + "^FO30,195^A1B,25,20,^FD" + num + "^FS\n"; // SE IMPRIME TIPO GRADO 1
                                            }

                                            if (clave_gab.Trim() == "05006MLNA2")
                                            {
                                                impr = impr + "^FO-15,175^A0B,17,18,^FD" + Convert.ToDateTime(date).ToString("MM/dd/yy") + "^FS\n"; // fecha etiqueta 1                                                                                            
                                            }

                                            impr = impr + "^FO320,140^A0N,22,18^FD" + "C: " + auxnum1 + "^FS\n"; // CAJA ETIQUETA 1                                            
                                            impr = impr + "^FO90,170,^BY1,^BCN,60,N,N,N^MD5F^FD" + "01" + aux3 + "^FS\n"; // CODIGO GTIN COD DE BARRAS 128 1                                            
                                            impr = impr + "^FO110,240^A0N,17,15,^FD" + aux3_texto + "^FS\n"; // DESCRIPCION DEL PRODUCTO EN EL CODIGO BQ LINEA 1                            
                                            impr = impr + "^FO350,165^A0B,17,18,^FD" + mfeccad + "^FS\n"; // SE IMPRIME LA FECHA DE CADUCIDAD LINEA 1                                            
                                            impr = impr + "^XZ\n";
                                        }
                                    }
                                    j++;
                                }//for         
                                impr = impr + "^XA\n";
                                impr = impr + "^FO80,30^A0N,45,20^FD -     -      -    -     -   -^FS\n";
                                impr = impr + "^FO460,30^A0N,45,20^FD -     -      -    -     -   -^FS\n";
                                impr = impr + "^FO60,30,^BY3,^BCN,80,N,N,N^FD00000^XZ\n";
                                //j = Convert.ToInt32(DGV.Rows[a].Cells["num_eti"].Value);
                                //co = 0; renglon_selec = 0;
                                //renglon_selec = Convert.ToInt32(DGV.Rows[a].Cells["eti_fin"].Value);
                                //j = Convert.ToInt32(DGV.Rows[a].Cells["num_eti"].Value);
                                //PaperSize ps;
                                //PrintDocument pd1 = new System.Drawing.Printing.PrintDocument();
                                //ps = new PaperSize("etiqueta_blanca", 400, 400);
                                //pd1.BeginPrint += new PrintEventHandler(this.eti_blanca_BeginPrint);
                                //pd1.PrintPage += new PrintPageEventHandler(this.eti_blanca_PrintPage);
                                //pd1.EndPrint += new PrintEventHandler(this.eti_blanca_EndPrint);
                                //pd1.DefaultPageSettings.PaperSize = ps;
                                //pd1.Print();
                                //ThreadStart delegado = new ThreadStart(CorrerProceso);
                                //Thread hilo = new Thread(delegado);
                                //hilo.Start();    
                                //Thread.Sleep(1000);  
                                //PaperSize ps;
                                //PrintPreviewDialog pr = new PrintPreviewDialog();
                                //pr.Document = eti_blanca;
                                //ps = new PaperSize("etiqueta_blanca", 400, 400);
                                //eti_blanca.DefaultPageSettings.PaperSize = ps;
                                //eti_blanca.Print();
                            }//foreach

                        }//if que valida que el rango inicial sea menor o igual al rango final 
                        else
                        {
                            MessageBox.Show("El número de etiquetas a imprimir sobrepasa el número de etiquetas que tiene la tarima", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            DGV.Focus();
                            DGV.Rows[a].Cells["num_eti"].Selected = true;
                            impr = "";
                            return;
                        }
                    }
                    #endregion                    

                    string auxtr_2 = "";
                    #region con pais
                    if (Convert.ToString(DGV.Rows[a].Cells["country"].Value).Trim() != "")
                    {
                        if ((Convert.ToInt32(DGV.Rows[a].Cells["eti_fin"].Value) >= Convert.ToInt32(DGV.Rows[a].Cells["num_eti"].Value)))
                        {
                            foreach (DataRow row in recibo.Select("prod_clave = '" + Convert.ToString(DGV.Rows[a].Cells["clave_prod"].Value) + "' and tarima = '" + Convert.ToString(DGV.Rows[a].Cells["num_tar"].Value) + "'"))
                            {
                                aux1 = Convert.ToString(row["pti_clave"].ToString()).Trim(); //pti_clave
                                aux2 = Convert.ToString(row["prod_nom_ingles"].ToString());//prod_nombre_ingles
                                aux3 = Convert.ToString(row["gtin_clave"].ToString());//gtin_clave
                                clave_gab = Convert.ToString(row["prod_clave"].ToString());//prod_clave
                                aux5 = Convert.ToString(row["prod_nombre"].ToString());//prod_nombre
                                aux5 = aux5.Replace("''", "");
                                aux5 = aux5.Replace("'", "");
                                auxn = Convert.ToDecimal(row["etiqueta"].ToString());//etiqueta                            
                                fecad = Convert.ToString(row["fecha_cad"].ToString());//fecha de caducidad 
                                pais = Convert.ToString(DGV.Rows[a].Cells["country"].Value);//pais origen del producto
                                auxtr_2 = clave_gab;//prod_clave      

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
                                //lon_cad2 = aux4_1.Length + 1;

                                nutar1 = aux4.Substring(lon_cad1 - 8, 2);
                                totar1 = aux4.Substring(lon_cad1 - 6, 2);
                                //nutar2 = aux4_1.Substring(lon_cad1 - 8, 2);
                                //totar2 = aux4_1.Substring(lon_cad1 - 6, 2);

                                for (int i = Convert.ToInt32(DGV.Rows[a].Cells["num_eti"].Value); i <= Convert.ToInt32(DGV.Rows[a].Cells["eti_fin"].Value); i++)
                                {
                                    impr2 += "^XA\n";
                                    impr2 += "^FO05,55^A2N,23,17,^FD" + txtrecibo.Text + "-" + clave_gab + " - " + nutar1 + "/" + totar1 + "^FS\n";
                                    impr2 += "^FO15,100^A0N,21,15^FDC:" + i.ToString() + "^FS\n";
                                    impr2 += "^FO200,100^BQN,4,4^FDLA,http://www.mrlucky.com.mx/tr/trazabilidad2_dmi.php?id_codigo=" + aux4 + "^FS\n";
                                    impr2 += "^FO15,250^A0N,48,24^FD" + aux2 + "^FS\n";
                                    impr2 += "^FO15,300^A0N,48,24^FD" + aux5 + "^FS\n";
                                    impr2 += "^FO15,350^A0N,25,25^FDPRODUCT OF " + Convert.ToString(DGV.Rows[a].Cells["country"].Value).Trim() + "^FS\n";
                                    impr2 += "^FO250,350^A0N,25,25^FDHEB US#1^FS\n";
                                    impr2 += "^FO50,380,^BY2,3,^BCN,90,N,N,N^FD01" + aux3 + "^FS\n";
                                    impr2 += "^FO150,480^A0N,25,25^FD" + aux3_texto + "^FS\n";
                                    impr2 += "^FO15,520^A0N,25,25^FDPACKED BY: COMERCIALIZADORA GAB, S.A. DE C.V.^FS\n";
                                    impr2 += "^FO15,550^A0N,25,25^FDCARRETERA PANAMERICANA KM. 291 - 1^FS\n";
                                    impr2 += "^FO15,580^A0N,25,25^FDCOL. LA FORTALEZA CORTAZAR, GTO. C.P. 38300^FS\n";
                                    impr2 += "^FO15,610^A0N,25,25^FDMEXICO R.F.C. CGA-960614-2C5^FS\n";
                                    impr2 += "^XZ\n";
                                }

                                //BarcodeLib.Barcode.Linear code128 = new BarcodeLib.Barcode.Linear();
                                //code128.Type = BarcodeLib.Barcode.BarcodeType.CODE128;
                                //code128.Data = aux3;
                                //code128.AddCheckSum = true;
                                //code128.UOM = BarcodeLib.Barcode.UnitOfMeasure.PIXEL;
                                //code128.BarHeight = 50;
                                //code128.BarWidth = 2;
                                //code128.drawBarcode("C:/sisgabweb/code128.jpg");
                                //newImage1 = Image.FromFile(@"C:/sisgabweb/code128.jpg");                                                          

                                //string lin = auxtr_2.Substring(0, 2);//saber la linea
                                //if ((auxtr_2.Trim() == "05003ML3P") || (auxtr_2.Trim() == "05005ML2P") || (auxtr_2.Trim() == "19"))
                                //{
                                //    dia = fecad.Substring(0, 2);
                                //    fecad = fecad.Substring(3, 2);
                                //    obtenerNombreMesNumero(Convert.ToInt32(fecad));
                                //    mfeccad = "FC" + mes.ToUpper() + dia;
                                //}
                                ////ultima = Convert.ToInt32(DGV.Rows[a].Cells["eti_fin"].Value);
                                //j = 1;
                                //co = 0; renglon_selec = 0;
                                ////renglon_selec = Convert.ToInt32(auxn);
                                //renglon_selec = Convert.ToInt32(DGV.Rows[a].Cells["eti_fin"].Value);
                                //PaperSize ps;
                                //PrintPreviewDialog pr = new PrintPreviewDialog();
                                //pr.Document = eti_blanca;
                                ////ps = new PaperSize("etiqueta_blanca", 425, 267);                           
                                //ps = new PaperSize("eti_blan", 267, 425);
                                //eti_blanca.DefaultPageSettings.PaperSize = ps;
                                //eti_blanca.DefaultPageSettings.Landscape = true;
                                //eti_blanca.Print();                                
                            }
                        }
                        else
                        {
                            MessageBox.Show("El número de etiquetas a imprimir sobrepasa el número de etiquetas que tiene la tarima", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            DGV.Focus();
                            DGV.Rows[a].Cells["num_eti"].Selected = true;
                            impr = "";
                            return;
                        }
                    }
                    #endregion
                    //}//if fila que esta seleccionada en el grid    

                    //ThreadStart delegado = new ThreadStart(CorrerProceso);
                    //Thread hilo = new Thread(delegado);
                    //hilo.Start();
                    //Thread.Sleep(10000);
                    //con++;
                    //mNUM_TAR = Convert.ToInt32(DGV.Rows[a].Cells["num_tar"].Value);
                    //clave_gab = Convert.ToString(DGV.Rows[a].Cells["clave_prod"].Value);
                    //fic = @"C:\\Reportes\\" + txtrecibo.Text + "_1.txt";

                    //if (File.Exists(fic))
                    //    File.Delete(fic);

                    //string ipaddress = "192.168.123.155";
                    //int puerto = 9100;
                    //byte[] bytes = System.IO.File.ReadAllBytes(@"C:\Reportes\document9.pdf");
                    //FileStream lpt1 = new FileStream()
                }

            }//que este seleccionada la tarima
            if (impr.Length > 0)
                RawPrinterHelper.SendStringToPrinter("etiqueta_blanca", impr);
            //RawPrinterHelper.SendStringToPrinter("ZDesigner ZT410-203dpi ZPL", impr);
            if (impr2.Length > 0)
                RawPrinterHelper.SendStringToPrinter("ZDesigner GX420t", impr2);
            //RawPrinterHelper.SendStringToPrinter("TSC MX240", impr);            


            //IPAddress addr = IPAddress.Parse("192.168.123.223");
            //EndPoint ep = new IPEndPoint(addr, 9100);
            //Socket sock = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            //sock.Connect(ep);
            //NetworkStream ns = new NetworkStream(sock);
            //String transferType = "^MTT\n";// Sets the type to thermal transfer
            //String ZPLString = "^LH5,5\n" + transferType +
            //                   "^BY2" + "^MNM\n" + impr;     
            //byte[] toSend = Encoding.ASCII.GetBytes(ZPLString);
            //ns.Write(toSend, 0, toSend.Length);
            aux1 = ""; aux2 = ""; aux3 = ""; clave_gab = ""; aux3_texto = ""; aux5 = ""; fecad = ""; mfeccad = "";
            auxc = ""; auxd = ""; aux4 = ""; auxnum1 = ""; auxnum2 = ""; auxd_1 = ""; aux4_1 = "";
            nutar1 = ""; totar1 = ""; nutar2 = ""; totar2 = ""; dia = ""; mes = ""; impr = "";
            auxn = 0; tot_cajas = 0;
            y = 0; lon_cad1 = 0; lon_cad2 = 0;



            //RawPrinterHelper.SendStringToPrinter("TSC MX240", ZPLString); 


            //for
            //ThreadStart delegado = new ThreadStart(CorrerProceso);
            //Thread hilo = new Thread(delegado);
            //hilo.Start();
            //Thread.Sleep(1000);
            //j = Convert.ToInt32(DGV.Rows[a].Cells["num_eti"].Value);
            //PaperSize ps;
            //PrintDocument pd1 = new System.Drawing.Printing.PrintDocument();
            //ps = new PaperSize("etiqueta_blanca", 400, 400);
            //pd1.BeginPrint += new PrintEventHandler(this.eti_blanca_BeginPrint);
            //pd1.PrintPage += new PrintPageEventHandler(this.eti_blanca_PrintPage);
            //pd1.EndPrint += new PrintEventHandler(this.eti_blanca_EndPrint);
            //pd1.DefaultPageSettings.PaperSize = ps;
            //pd1.Print();   
        }

        private void CorrerProceso()
        {
            PaperSize ps;
            PrintPreviewDialog pr = new PrintPreviewDialog();
            pr.Document = eti_blanca;
            ps = new PaperSize("etiqueta_blanca", 400, 400);
            eti_blanca.DefaultPageSettings.PaperSize = ps;
            eti_blanca.Print();
            //PrintDocument pd1 = new System.Drawing.Printing.PrintDocument();
            //pd1.BeginPrint += new PrintEventHandler(this.eti_blanca_BeginPrint);
            //pd1.PrintPage += new PrintPageEventHandler(this.eti_blanca_PrintPage);
            //pd1.EndPrint += new PrintEventHandler(this.eti_blanca_EndPrint);
            //pd1.DefaultPageSettings.PaperSize = ps;            
            //pd1.Print();
            //RawPrinterHelper.SendFileToPrinter(@"C:\Reportes\doc.pdf");            
            //MessageBox.Show("Proceso finalizado");
            //string pathToExecutable = "C:\\Program Files (x86)\\Adobe\\Acrobat Reader DC\\Reader\\acrord32.exe";
            //RunExecutable(pathToExecutable, @"/t ""C:\Reportes\doc.pdf"" ""My Windows PrinterName""");            
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

        int j = 1, co = 0, renglon_selec = 0;
        public void eti_blanca_PrintPage(object sender, PrintPageEventArgs e)
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

                e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar1 + "/" + totar1, fuente_encabezados, color, 5, 5);// primera linea etiqueta 1

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
                }
                else
                {
                    e.Graphics.DrawString(aux5, nota, color, 5, 60);//nombre en español completo etiqueta 1
                    //e.Graphics.DrawString(aux5, nota, color, 210, 41);//nombre en español completo etiqueta 2
                }

                e.Graphics.DrawImage(newImage1, 30, 115);

                StringFormat drawFormat = new StringFormat();
                drawFormat.FormatFlags = StringFormatFlags.DirectionVertical;
                e.Graphics.DrawString(mfeccad, cadu, color, 170, 80, drawFormat);//fecha de caducidad etiqueta 1


                e.Graphics.DrawString("PRODUCT OF " + pais, box, color, 5, 97);//país
                e.Graphics.DrawString("C: " + auxnum1, box, color, 350, 77);//número de caja etiqueta 1                

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

                e.Graphics.DrawString("PACKED BY:", box, color, 5, 190);
                e.Graphics.DrawString("COMERCIALIZADORA GAB, S.A. DE C.V.", box, color, 5, 203);
                e.Graphics.DrawString("CARRETERA PANAMERICANA KM. 291 - 1", box, color, 5, 216);
                e.Graphics.DrawString("COL. LA FORTALEZA CORTAZAR, GTO.", box, color, 5, 229);
                e.Graphics.DrawString("C.P. 38300 MEXICO", box, color, 5, 242);
                e.Graphics.DrawString("R.F.C. CGA-960614-2C5", box, color, 5, 255);

                e.Graphics.DrawString("US #1", nota, color, 180, 220);
                e.Graphics.DrawString("HEB", nota, color, 300, 220);

                j++;
            }

            if (co == 2)
            {
                co = 0;
                e.Graphics.DrawString("|   |   |   |   |   |   |   |   |   |", nota, color, 230, 50); //numero codigo de barras etiqueta 1                
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

            /*
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

                    e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar1 + "/" + totar1, fuente_encabezados, color, 3, 5);// primera linea etiqueta 1                                
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

                    //BarcodeLib.Barcode.Linear code128 = new BarcodeLib.Barcode.Linear();
                    //code128.Type = BarcodeLib.Barcode.BarcodeType.CODE128;
                    //code128.Data = aux3;
                    //code128.AddCheckSum = true;
                    //code128.UOM = BarcodeLib.Barcode.UnitOfMeasure.PIXEL;
                    //code128.BarHeight = 50;
                    //code128.BarWidth = 2;
                    ////code128.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                    ////code128.drawBarcode("C:/sisgabweb/code128.jpg");
                    //byte[] code821 = code128.drawBarcodeAsBytes();
                    //Stream ms1 = new MemoryStream(code821);
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
                        e.Graphics.DrawImage(newImage2, 333, 3, 60, 60);//qrcode etiqueta 2
                        ms2.Close();
                        imag2.Dispose();
                    }

                    e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar1 + "/" + totar1, fuente_encabezados, color, 3, 5);// primera linea etiqueta 1
                    e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar2 + "/" + totar2, fuente_encabezados, color, 210, 5);// primera linea etiqueta 2
                    if (aux2.Length > 25)
                    {
                        e.Graphics.DrawString(aux2.Substring(0, 24), nota, color, 3, 15);//nombre en ingles parte 1 etiqueta 1
                        e.Graphics.DrawString(aux2.Substring(25), nota, color, 3, 28);//nombre en ingles parte 2 etiqueta 1

                        e.Graphics.DrawString(aux2.Substring(0, 24), nota, color, 210, 15);//nombre en ingles parte 1 etiqueta 2
                        e.Graphics.DrawString(aux2.Substring(25), nota, color, 210, 28);//nombre en ingles parte 2 etiqueta 2
                    }
                    else
                    {
                        e.Graphics.DrawString(aux2, nota, color, 3, 15);//nombre en ingles completo etiqueta 1
                        e.Graphics.DrawString(aux2, nota, color, 210, 15);//nombre en ingles completo etiqueta 2
                    }
                    if (aux5.Length > 25)
                    {
                        e.Graphics.DrawString(aux5.Substring(0, 24), nota, color, 3, 41);//nombre en español parte 1 etiqueta 1
                        e.Graphics.DrawString(aux5.Substring(25), nota, color, 3, 54);//nombre en español parte 2 etiqueta 1

                        e.Graphics.DrawString(aux5.Substring(0, 24), nota, color, 210, 41);//nombre en español parte 1 etiqueta 2
                        e.Graphics.DrawString(aux5.Substring(25), nota, color, 210, 54);//nombre en español parte 2 etiqueta 2
                    }
                    else
                    {
                        e.Graphics.DrawString(aux5, nota, color, 3, 41);//nombre en español completo etiqueta 1
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
                //co++;
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
                    e.Graphics.DrawImage(newImage2, 333, 3, 60, 60);//qrcode etiqueta 2
                    ms2.Close();
                    imag2.Dispose();
                }

                e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar1 + "/" + totar1, fuente_encabezados, color, 3, 5);// primera linea etiqueta 1
                e.Graphics.DrawString(aux1.Substring(0, 6) + "-" + clave_gab + "-" + nutar2 + "/" + totar2, fuente_encabezados, color, 210, 5);// primera linea etiqueta 2
                if (aux2.Length > 25)
                {
                    e.Graphics.DrawString(aux2.Substring(0, 24), nota, color, 3, 15);//nombre en ingles parte 1 etiqueta 1
                    e.Graphics.DrawString(aux2.Substring(25), nota, color, 3, 28);//nombre en ingles parte 2 etiqueta 1

                    e.Graphics.DrawString(aux2.Substring(0, 24), nota, color, 210, 15);//nombre en ingles parte 1 etiqueta 2
                    e.Graphics.DrawString(aux2.Substring(25), nota, color, 210, 28);//nombre en ingles parte 2 etiqueta 2
                }
                else
                {
                    e.Graphics.DrawString(aux2, nota, color, 3, 15);//nombre en ingles completo etiqueta 1
                    e.Graphics.DrawString(aux2, nota, color, 210, 15);//nombre en ingles completo etiqueta 2
                }
                if (aux5.Length > 25)
                {
                    e.Graphics.DrawString(aux5.Substring(0, 24), nota, color, 3, 41);//nombre en español parte 1 etiqueta 1
                    e.Graphics.DrawString(aux5.Substring(25), nota, color, 3, 54);//nombre en español parte 2 etiqueta 1

                    e.Graphics.DrawString(aux5.Substring(0, 24), nota, color, 210, 41);//nombre en español parte 1 etiqueta 2
                    e.Graphics.DrawString(aux5.Substring(25), nota, color, 210, 54);//nombre en español parte 2 etiqueta 2
                }
                else
                {
                    e.Graphics.DrawString(aux5, nota, color, 3, 41);//nombre en español completo etiqueta 1
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
                e.Graphics.DrawImage(newImage1, 205, 8, 150, 30);// codebar 128 etiqueta 2

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
                e.Graphics.DrawString("|   |   |   |   |   |   |   |   |   |", nota, color, 230, 50); //numero codigo de barras etiqueta 1
                e.Graphics.DrawString("|   |   |   |   |   |   |   |   |   |", nota, color, 3, 50); //numero codigo de barras etiqueta 2
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
                co = 1;
            }   */
        }

        SolidBrush color;
        Font fuente_encabezados;
        Font nota;
        Font cadu;
        Font box;

        private void eti_blanca_BeginPrint(object sender, PrintEventArgs e)
        {
            color = new SolidBrush(Color.Black);
            fuente_encabezados = new Font("Arial", 7);// encabezados                        
            nota = new Font("Impact", 7);//nombre del producto
            cadu = new Font("Impact", 6);//fecha caducidad            
            box = new Font("Impact", 7);//caja
        }

        private void eti_blanca_EndPrint(object sender, PrintEventArgs e)
        {
            color.Dispose();
            fuente_encabezados.Dispose();
            nota.Dispose();
            cadu.Dispose();
            box.Dispose();
        }

        private void cbgrado_SelectionChangeCommitted(object sender, EventArgs e)
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

            for (int i = 0; i < DGV.Rows.Count; i++)
            {
                if (fila.Trim() == Convert.ToString(DGV.Rows[i].Cells["clave_prod"].Value).Trim())
                    DGV.Rows[i].Cells["gra"].Value = grado;
            }
            DGV.Refresh();
        }
    }
}
