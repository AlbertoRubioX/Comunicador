using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Datos;

namespace Logica
{
    public class EmpleadoLogica
    {
        public string Empleado { get; set; }
        public string Nombre { get; set; }
        public string Puesto { get; set; }
        public string Area { get; set; }
        public string Turno { get; set; }
        public string Calle { get; set; }
        public string Colonia { get; set; }
        public string CP { get; set; }
        public string Latitud { get; set; }
        public string Longitud { get; set; }


        public static int Guardar(EmpleadoLogica emp)
        {
            string[] parametros = { "@Empleado", "@Nombre", "@Puesto", "@Area", "@Turno", "@Calle", "@Colonia", "@CP", "@Latitud", "@Longitud" };
            return AccesoDatos.Actualizar("sp_mant_empleado", parametros, emp.Empleado, emp.Nombre, emp.Puesto, emp.Area, emp.Turno, emp.Calle, emp.Colonia, emp.CP, emp.Latitud, emp.Longitud);
        }

        public static DataTable Consultar()
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("SELECT * FROM t_empleado");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        
    }
}
