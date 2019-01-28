using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Datos;
using System.Data;

namespace Logica
{
    public class ComunicaLogica
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

        public static int AlertaPRODUCTEC()
        {
            return AccesoDatos.EjecutaSP("sp_aler_productividadDia");
        }
        public static int AlertaMaqCosto()
        {
            return AccesoDatos.EjecutaSP("sp_aler_preciocosto");
        }
        public static int AlertaTop10Servicios()
        {
            return AccesoDatos.EjecutaSP("sp_aler_top10servicios");
        }
        public static int AlertaTop10CtoServ()
        {
            return AccesoDatos.EjecutaSP("sp_aler_top10ctoserv");
        }
        public static int AlertaOSFECOM()
        {
            return AccesoDatos.EjecutaSP("sp_aler_ordfcomp");
        }

        public static int AlertaOSPROCOM()
        {
            return AccesoDatos.EjecutaSP("sp_aler_reqprocom");
        }

        public static int AlertaEXIS()
        {
            return AccesoDatos.EjecutaSP("sp_aler_parteminmax");
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
                datos = AccesoDatos.Consultar("select co.folio as Folio,co.proceso,mo.descrip as Proceso,co.referencia as Referencia,co.destino as Destinatario, co.asunto as Asunto,case co.estado when 'P' then 'Pendiente' when 'R' then 'Error' end as Estado,co.mensaje,co.u_id from t_correo co inner join t_mod02 mo on co.proceso = mo.proceso where co.estado = 'P' or co.estado = 'R' order by co.folio");
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
