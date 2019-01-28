using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Datos;
using System.Data;

namespace Logica
{
    public class ComunicaCPROLogica
    {
        public long Folio { get; set; }
        public string Proceso { get; set; }
        public string Referencia { get; set; }
        public string Estado { get; set; }
        public string Destino { get; set; }
        public string Asunto { get; set; }
        public string Mensaje { get; set; }
        public string Usuario { get; set; }

        public static int Guardar(ComunicaLogica com)
        {
            string[] parametros = { "@Folio", "@Proceso", "@Referencia", "@Estado", "@Destino", "@Asunto", "Mensaje", "Usuario" };
            return AccesoDatos.Actualizar("sp_mant_correo", parametros, com.Folio, com.Proceso, com.Referencia, com.Estado, com.Destino, com.Asunto, com.Mensaje, com.Usuario);
        }

        public static int AlertaPRESETUP()
        {
            return AccesoDatos.EjecutaSPCPRO("sp_alert_setupprevio");
        }

        public static int AlertaKANBAN()
        {
            return AccesoDatos.EjecutaSPCPRO("sp_alert_kanban");
        }

        public static int AlertaKANBAlmacen()
        {
            return AccesoDatos.EjecutaSPCPRO("sp_alert_kanban_almacen");
        }

        public static int AlertaHORASETUP()
        {
            return AccesoDatos.EjecutaSPCPRO("sp_alert_setuphora");
        }
        public static int AlertaDURACIONSETUP()
        {
            return AccesoDatos.EjecutaSPCPRO("sp_alert_setupdura");
        }

        public static int AlertaGLOBALENVIOS()
        {
            return AccesoDatos.EjecutaSPCPRO("sp_alert_globalEnvios");
        }

        public static bool AlertaDiaria(ComunicaLogica com)
        {
            DataTable datos = new DataTable();
            try
            {
                string sQuery;
                sQuery = "select count(*) from t_alerta_dia where proceso = '"+com.Proceso+"' and alerta = '"+com.Referencia+ "' and CAST(fecha as date) = CAST('" + DateTime.Today + "' as date) HAVING COUNT(*) > 0 UNION ";
                sQuery += "select COUNT(*) from t_correo where proceso = '"+com.Proceso+"' and referencia = '"+com.Referencia+"' and CAST(fecha as date) = CAST('" + DateTime.Today + "' as date) HAVING COUNT(*) > 0";
                datos = AccesoDatos.Consultar(sQuery);
                if (datos.Rows.Count > 0)
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static bool AlertaMensual(ComunicaLogica com)
        {
            DataTable datos = new DataTable();
            try
            {
                string sQuery;
                sQuery = "select count(*) from t_correo where proceso = '" + com.Proceso + "' and referencia = '" + com.Referencia + "' and YEAR(fecha) = YEAR('" + DateTime.Today + "') and MONTH(fecha) = MONTH('" + DateTime.Today + "') HAVING COUNT(*) > 0";
                datos = AccesoDatos.Consultar(sQuery);
                if (datos.Rows.Count > 0)
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static DataTable EnviosPendientes()
        {
            DataTable datos = new DataTable();
            try
            {
                string sSql = "select co.folio as Folio,co.proceso,mo.descrip as Proceso,co.folio_ref as Referencia,'' as Destinatario, co.asunto as Asunto," +
                "case co.estado when 'P' then 'Pendiente' when 'R' then 'Error' else 'Enviado' end as Estado,'','' " +
                "from t_correo co inner join t_mod02 mo on co.proceso = mo.proceso where co.estado = 'P' or co.estado = 'R' order by co.folio";
                datos = AccesoDatos.ConsultarPRO(sSql);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static DataTable Enviados()
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("select co.f_id as [Fecha de Envio],co.folio as Folio,mo.descrip as Proceso,co.destino as Destinatario, co.asunto as Asunto from t_correo co inner join t_mod02 mo on co.proceso = mo.proceso where co.estado = 'E' order by co.folio desc");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static bool Eliminar(ComunicaLogica com)
        {
            try
            {
                string sQuery = "";
                sQuery = "DELETE FROM t_correo_bo WHERE folio = " + com.Folio + " ";
                AccesoDatos.Borrar(sQuery);

                //sQuery = "DELETE FROM t_correo_to WHERE folio = " + com.Folio + " ";
                //AccesoDatos.Borrar(sQuery);

                sQuery = "DELETE FROM t_correo WHERE folio = " + com.Folio + " ";
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

        public static bool EliminarEnviados()
        {
            try
            {
                string sQuery = "DELETE FROM t_correo WHERE estado = 'E' ";
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
