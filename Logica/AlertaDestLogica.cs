using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Datos;

namespace Logica
{
    public class AlertaDestLogica
    {
        public string Proceso { get; set; }
        public string Alerta { get; set; }
        
        public static DataTable Consultar(AlertaDestLogica alerta)
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("SELECT consec,usuario,correo FROM t_alerta_dest where proceso = '" + alerta.Proceso + "' and alerta = '" + alerta.Alerta + "'");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }
    }
}
