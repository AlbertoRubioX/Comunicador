using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Datos;

namespace Logica
{
    public class ConfigDestLogica
    {
        public int Consec { get; set; }
        public string Proceso { get; set; }
        public string Correo { get; set; }
        public string Tipo { get; set; }
        public string Usuario { get; set; }

        public static int Guardar(ConfigDestLogica dest)
        {
            string[] parametros = { "@Consec", "@Proceso", "@Correo", "@Tipo", "@Usuario" };
            return AccesoDatos.Actualizar("sp_mant_config_dest", parametros, dest.Consec, dest.Proceso, dest.Correo, dest.Tipo, dest.Usuario);
        }

        public static DataTable Consultar(string _asProceso)
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("SELECT * FROM t_config_dest where proceso = '" + _asProceso + "' order by correo");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static DataTable Listar(string _asProceso)
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("SELECT consec,proceso,correo as [Correo Electrónico],case tipo when 'T' then 'To' when 'C' then 'Cc' end as Tipo FROM t_config_dest where proceso = '" + _asProceso + "' order by correo");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }
        public static bool verificaCorreo(ConfigDestLogica dest)
        {
            DataTable datos = new DataTable();
            datos = AccesoDatos.Consultar("select * from t_config_det where proceso = '" + dest.Proceso + "' and correo = '" + dest.Correo + "' ");
            if (datos.Rows.Count != 0)
                return true;
            else
                return false;
        }

        public static bool Eliminar(ConfigDestLogica dest)
        {
            try
            {
                string sQuery = "DELETE FROM t_config_dest WHERE consec = " + dest.Consec + "";
                if (AccesoDatos.Borrar(sQuery) != 0)
                    return true;
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }
    }
}
