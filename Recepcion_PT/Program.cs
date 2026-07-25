using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Utilerias;
using System.Diagnostics;
using System.Reflection;

namespace Recepcion_PT
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        

        public static bool validar_login()
        {
            SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
            SqlCommand cmnd2;
            SqlDataReader reader2;

            //bool esta = false;
            //Utilerias.Login = false;
            //foreach (Process proceso in Process.GetProcesses())
            //{
            //    if (proceso.ProcessName == "SIPGAB")
            //    {
            //        esta = true;
            //    }
            //}
            //args = new string[1] { "N" };
            //if (esta == false)
            //{
            thisConnection.Open();
            cmnd2 = thisConnection.CreateCommand();
            cmnd2.CommandText = "select TOP 1 usu_login from tb_cat_historial_dia where nombre_maquina ='" + Environment.MachineName + "' AND fin_sesion IS NULL and sistema = 'SIPGAB' ORDER BY inicio_sesion DESC";
            reader2 = cmnd2.ExecuteReader();
            while (reader2.Read())
            {
                Utilerias.Class1.Usu_login = reader2.GetSqlString(0).ToString();
                Utilerias.Class1.Login = true;
            }
            if (reader2.HasRows == false)
            {
                MessageBox.Show("Error: Favor de iniciar sesion", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //System.Diagnostics.Process MiProceso = new System.Diagnostics.Process();
                //MiProceso.StartInfo.WorkingDirectory = "C:\\SisGabWeb"; // ubicacion donde esta el ejecutable
                //MiProceso.StartInfo.FileName = "SIPGAB.exe"; // nombre del archivo a ejecutar con su extension
                ////MiProceso.StartInfo.Arguments = Utilerias.Grupo.ToString(); // esto es opcional en caso que el ejecutablee reciba parametros
                //MiProceso.Start(); // inicia el ejecutable     
                ProcessStartInfo proces = new ProcessStartInfo(@"C:\\SisGabWeb\SIPGAB.exe");
                Process.Start(proces);
                Utilerias.Class1.Login = false;
                Application.Exit();
                //Application.Run(new OrdVenNal());                   
            }
            //}
            thisConnection.Close();
            return Utilerias.Class1.Login;
        }
        [STAThread]
        static void Main()
        {
            //Assembly miExtensionAssembly = Assembly.LoadFile(@"C:\SisGabWeb\Utilerias.dll");
            //Type miExtensionType = miExtensionAssembly.GetType("Utilerias.Class1");
            //object miExtensionObjeto = Activator.CreateInstance(miExtensionType);
            //miExtensionObjeto.GetType().InvokeMember("ConnectionString", BindingFlags.InvokeMethod, Type.DefaultBinder, miExtensionObjeto, null);
            //ITipoQueSea objetoFinal = (ITipoQueSea)miExtensionObjeto;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Utilerias.Class1.validar_ip();
            Utilerias.Class1.validar_ipfox();
            validar_login();
            if (Utilerias.Class1.Login == true)
                Application.Run(new Recepcion_PT());

            if (Utilerias.Class1.Login == false)
                Application.Exit();        
            
        }
    }
}
