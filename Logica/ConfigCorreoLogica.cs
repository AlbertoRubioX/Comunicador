using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Datos;

namespace Logica
{
    public class ConfigCorreoLogica
    {
        public string Correo { get; set; }
        public string Servidor { get; set; }
        public int Puerto { get; set; }
        public string Usuario { get; set; }
        public string Password { get; set; }
        public string Ssl { get; set; }
        public string Html { get; set; }

        public static int Guardar(ConfigCorreoLogica config)
        {
            string[] parametros = { "@Correo", "@Servidor", "@Puerto", "@Usuario", "@Password", "@Ssl", "@Html" };
            return AccesoDatos.Actualizar("sp_mant_config_correo", parametros, config.Correo, config.Servidor, config.Puerto, config.Usuario, config.Password, config.Ssl, config.Html);
        }

        public static DataTable Consultar()
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("SELECT * FROM t_config WHERE clave = '01'");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static DataTable ConsultarSMTP()
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("SELECT * FROM t_config_correo WHERE clave = '01'");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static bool verificaEnvio()
        {
            DataTable datos = new DataTable();
            datos = AccesoDatos.Consultar("select * from t_config where clave = '01' and ind_correo = '1'");
            if (datos.Rows.Count != 0)
                return true;
            else
                return false;
        }
        public static bool verificaEnvioServ()
        {
            DataTable datos = new DataTable();
            datos = AccesoDatos.Consultar("select * from t_config where clave = '01' and ind_correo = '1' and env_oserv = '1'");
            if (datos.Rows.Count != 0)
                return true;
            else
                return false;
        }
    }
}
