using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Datos;

namespace Logica
{
    public class MonitorSetupLogica
    {
        public DateTime FechaIni { get; set; }
        public DateTime FechaFin { get; set; }
        public string IndPlanta { get; set; }
        public string Planta { get; set; }
        public string IndLinea { get; set; }
        public string Linea { get; set; }
        public string IndTurno { get; set; }
        public string Turno { get; set; }
        public string IndEstatus { get; set; }
        public string Estatus { get; set; }

        public static DataTable ListarSP(MonitorSetupLogica set)
        {
            DataTable datos = new DataTable();
            try
            {
                string[] parametros = { "@FechaIni", "@FechaFin", "@IndPlanta", "@Planta", "@IndLinea", "@Linea", "@IndTurno", "@Turno", "@IndEst", "@Estatus" };
                datos = AccesoDatos.ConsultaSP("sp_mon_lineset", parametros, set.FechaIni, set.FechaFin, set.IndPlanta, set.Planta, set.IndLinea, set.Linea, set.IndTurno, set.Turno, set.IndEstatus, set.Estatus);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static DataTable VisorSP(MonitorSetupLogica set)
        {
            DataTable datos = new DataTable();
            try
            {
                string[] parametros = { "@FechaIni", "@FechaFin", "@IndPlanta", "@Planta", "@IndLinea", "@Linea", "@Turno", "@Estatus" };
                datos = AccesoDatos.ConsultaSP("sp_mon_lineset_visor", parametros, set.FechaIni, set.FechaFin, set.IndPlanta, set.Planta, set.IndLinea, set.Linea, set.Turno, set.Estatus);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static DataTable DuracionSP(MonitorSetupLogica set)
        {
            DataTable datos = new DataTable();
            try
            {
                string[] parametros = { "@FechaIni", "@FechaFin", "@IndPlanta", "@Planta", "@IndLinea", "@Linea" };
                datos = AccesoDatos.ConsultaSPPRO("sp_rep_setup_dura", parametros, set.FechaIni, set.FechaFin, set.IndPlanta, set.Planta, set.IndLinea, set.Linea);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }
        public static DataTable ListarSetup(MonitorSetupLogica det)
        {
            DataTable datos = new DataTable();
            try
            {
                string sSql = "SELECT st.fecha, sd.folio, sd.consec, sd.turno, sd.planta, sd.linea, "+
                "sd.rpo, sd.rpo_sig, case sd.inicio_prog when '' then '' else convert(varchar(10), st.fecha, 101) + ' ' + sd.inicio_prog end, sd.duracion "+  
                "FROM t_lineset st INNER JOIN t_linesedet sd ON st.folio = sd.folio "+
                "WHERE sd.estatus <> 'C' AND (CAST(st.fecha AS DATE) between CAST('" + det.FechaIni+"' AS DATE) AND CAST('"+det.FechaIni+"' AS DATE)) "+
                "ORDER BY sd.planta,sd.linea,st.fecha";
                datos = AccesoDatos.ConsultarPRO(sSql);
                
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static bool Listar(MonitorSetupLogica det)
        {
            DataTable datos = new DataTable();
            try
            {
                string sSql = "SELECT * FROM t_linesedet where u_id = 'SYSCOM' AND f_id >= '"+det.FechaIni+"' ";
                datos = AccesoDatos.ConsultarPRO(sSql);
                if (datos.Rows.Count > 0)
                    return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return false;
        }
    }
}
