using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Datos;

namespace Logica
{
    public class KardexCPROLogica
    {
        public int Folio { get; set; }
        public DateTime Fecha { get; set; }
        public string Proceso { get; set; }
        public string Referencia { get; set; }
        public string Result { get; set; }
        public string Hora { get; set; }
        public string Usuario { get; set; }

        public static int Guardar(KardexCPROLogica kar)
        {
            string[] parametros = { "@Proceso", "@Referencia", "@Result", "@Hora", "@Usuario" };
            return AccesoDatos.ActualizarPRO("sp_kardex_rt", parametros, kar.Proceso, kar.Referencia, kar.Result, kar.Hora, kar.Usuario);
        }
        

        public static DataTable Consultar(KardexCPROLogica kar)
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.ConsultarPRO("SELECT * FROM t_kardex_rt where proceso = '"+kar.Proceso+"' and cast(fecha as date) = cast('" + kar.Fecha + "' as date) and archivo = '" + kar.Referencia + "' ");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static bool Verificar(KardexCPROLogica kar)
        {
            bool bReturn = false; 
            try
            {
                DataTable datos = new DataTable();
                datos = AccesoDatos.ConsultarPRO("SELECT * FROM t_kardex_rt WHERE proceso = '" + kar.Proceso + "' AND cast(fecha as date) = cast('" + kar.Fecha + "' as date) AND hora = '"+kar.Hora+"' ");
                if (datos.Rows.Count > 0)
                    bReturn = true;
                else
                    bReturn = false;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return bReturn;
        }

    }
}
