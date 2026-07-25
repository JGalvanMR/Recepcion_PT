using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Net;
using System.Diagnostics;

namespace SIPGAB
{
    class Utilerias
    {
        //static string connectionString = "Data Source=201.120.60.21,2351;Initial Catalog=GAB_Irapuato;Connect Timeout=130;User ID=sa; MultipleActiveResultSets=True";
        //static string connectionString = "Data Source=gab.dyndns.org,2351;Initial Catalog=GAB_Irapuato;Connect Timeout=130;User ID=sa; MultipleActiveResultSets=True";
        //static string connectionString = "Data Source= GABIRA1\\SQL2005;Initial Catalog=GAB_Irapuato;Connect Timeout=130;User ID=sa;  MultipleActiveResultSets=True";
        static string connectionString, connectionStringFox;

        public static string validar_ip()
        {
            try
            {
                IPHostEntry host;
                string localIP = "";
                host = Dns.GetHostEntry(Dns.GetHostName());
                foreach (IPAddress ip in host.AddressList)
                {
                    if (ip.AddressFamily.ToString() == "InterNetwork")
                    {
                        localIP = ip.ToString();
                    }
                }

                int count = localIP.Length - 3;
                int cont = 0;
                string ips = "192.168.123.";
                for (int i = 0; i <= count - 1; i++)
                {
                    char caracter = localIP[i];
                    char caracter2 = ips[i];
                    if (caracter == caracter2)
                    {
                        cont++;
                    }
                    else
                    {
                        break;
                    }
                }

                if (count >= 12)
                {
                    ConnectionString = "Data Source= GABIRA1\\SQL2005;Initial Catalog=GAB_Irapuato;Connect Timeout=130;User ID=sa;  MultipleActiveResultSets=True";                    
                    //ConnectionString = "Data Source=201.120.60.21,2351;Initial Catalog=GAB_Irapuato;Connect Timeout=130;User ID=sa; MultipleActiveResultSets=True";                    
                }
                else
                {
                    ConnectionString = "Data Source=201.120.60.21,2351;Initial Catalog=GAB_Irapuato;Connect Timeout=130;User ID=sa; MultipleActiveResultSets=True";                    
                    //ConnectionString = "Data Source= gab.dyndns.org,2351;Initial Catalog=GAB_Irapuato;Connect Timeout=130;User ID=sa; MultipleActiveResultSets=True";                    
                }
                return ConnectionString;                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return ConnectionString = "error";
            }
        }

        public static string validar_ipfox()
        {
            try
            {
                IPHostEntry host;
                string localIP = "";
                host = Dns.GetHostEntry(Dns.GetHostName());
                foreach (IPAddress ip in host.AddressList)
                {
                    if (ip.AddressFamily.ToString() == "InterNetwork")
                    {
                        localIP = ip.ToString();
                    }
                }

                int count = localIP.Length - 3;
                int cont = 0;
                string ips = "192.168.123.";
                for (int i = 0; i <= count - 1; i++)
                {
                    char caracter = localIP[i];
                    char caracter2 = ips[i];
                    if (caracter == caracter2)
                    {
                        cont++;
                    }
                    else
                    {
                        break;
                    }
                }

                if (count <= 12)
                {                    
                    //connectionStringFox = @"Provider= VFPOLEDB.1;Data Source=gab.dyndns.org,2351;Collating Sequence=general;";
                    //connectionStringFox = @"Provider= VFPOLEDB.1;Data Source=\\gabira1\mr_lucky\base_de_datos;Collating Sequence=general;";
                    connectionStringFox = @"Provider= VFPOLEDB;Data Source=\\gabira1\mr_lucky\base_de_datos;Collating Sequence=general;";
                }
                else
                {                 
                    connectionStringFox = @"Provider= VFPOLEDB;Data Source=\\gabira1\mr_lucky\base_de_datos;Collating Sequence=general;";
                }
                return connectionStringFox;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return connectionStringFox = "error";
            }
        }

        public static string ConnectionString
        {
            get { return Utilerias.connectionString; }
            set { Utilerias.connectionString = value; }
        }

        public static string ConnectionStringFox
        {
            get { return Utilerias.connectionStringFox; }
            set { Utilerias.connectionStringFox = value; }
        }

        static SqlConnection cn = new SqlConnection(ConnectionString);
        static OleDbConnection cn1 = new OleDbConnection(connectionStringFox);

        private static string grupo;

        public static string Grupo
        {
            get { return Utilerias.grupo; }
            set { Utilerias.grupo = value; }
        }

        private static string usuario;

        public static string Usuario
        {
            get { return Utilerias.usuario; }
            set { Utilerias.usuario = value; }
        }

        private static string usu_login;

        public static string Usu_login
        {
            get { return Utilerias.usu_login; }
            set { Utilerias.usu_login = value; }
        }

        private static DateTime inicio_sesion;

        public static DateTime Inicio_sesion
        {
            get { return Utilerias.inicio_sesion; }
            set { Utilerias.inicio_sesion = value; }
        }

        private static DateTime fin_sesion;

        public static DateTime Fin_sesion
        {
            get { return Utilerias.fin_sesion; }
            set { Utilerias.fin_sesion = value; }
        }

        private static string nombre_equipo;

        public static string Nombre_equipo
        {
            get { return Utilerias.nombre_equipo; }
            set { Utilerias.nombre_equipo = value; }
        }

        private static string nombre_ingles;

        public static string Nombre_ingles
        {
            get { return Utilerias.nombre_ingles; }
            set { Utilerias.nombre_ingles = value; }
        }

        private static bool login;

        public static bool Login
        {
            get { return Utilerias.login; }
            set { Utilerias.login = value; }

        }

        private static string nombre_impresora;

        public static string Nombre_impresora
        {
            get { return Utilerias.nombre_impresora; }
            set { Utilerias.nombre_impresora = value; }
        }

        private static string formulario;

        public static string Formulario
        {
            get { return Utilerias.formulario; }
            set { Utilerias.formulario = value; }
        }

        //registra los eventos del usuario en el sistema
        public static void registrar_movimiento(DateTime fecha, string nom_compu, string nom_usu, string tipo_mov, string op_clave, string folio, string detalle)
        {
            SqlConnection thisConnection = new System.Data.SqlClient.SqlConnection(Utilerias.ConnectionString);
            try
            {
                SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
                SqlDataReader reader;
                thisConnection.Open();
                cmd = thisConnection.CreateCommand();
                cmd.CommandText = "insert into tb_registro_movimientos (fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, detalle) values " +
                                " ('" + fecha.ToString("s") + "', '" + nom_compu + "', '" + nom_usu + "', '" + tipo_mov + "', '" + op_clave + "', '" + folio + "', '" + detalle + "')";
                reader = cmd.ExecuteReader();
                reader.Dispose();
                thisConnection.Close();
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                MessageBox.Show(ex.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
