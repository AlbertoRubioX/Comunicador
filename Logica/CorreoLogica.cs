using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Datos;

namespace Logica
{
    public class CorreoLogica
    {
        public int Folio { get; set; }
        public string Proceso { get; set; }
        public string Referencia { get; set; }
        public string Estado { get; set; }
        public string Destino { get; set; }
        public string Asunto { get; set; }
        public string Mensaje { get; set; }
        public string Usuario { get; set; }

        public static int Guardar(CorreoLogica mail)
        {
            string[] parametros = { "@Folio", "@Proceso", "@Referencia", "@Estado", "@Destino", "@Asunto", "@Mensaje", "@Usuario" };
            return AccesoDatos.Actualizar("sp_mant_correo", parametros, mail.Folio, mail.Proceso, mail.Referencia, mail.Estado, mail.Destino, mail.Asunto, mail.Mensaje, mail.Usuario);
        }

        public static DataTable Consultar()
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("SELECT * FROM t_correo order by folio");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static DataTable ConsultarProcPend(CorreoLogica mail)
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("SELECT * FROM t_correo where estado = 'P' and proceso = '" + mail.Proceso + "' and referencia = '" + mail.Referencia + "' order by folio desc");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static DataTable ConsultarPend()
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("SELECT * FROM t_correo where estado = 'P' order by folio");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static DataTable ConsultarEnv()
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("SELECT * FROM t_correo where estado = 'E' order by folio");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static DataTable BodyMail(long _alFolio)
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("SELECT * FROM t_correo_bo where folio = " + _alFolio + " order by consec");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }
    }
}
