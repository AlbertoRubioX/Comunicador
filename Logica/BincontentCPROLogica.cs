using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using Datos;

namespace Logica
{
    public class BincontentCPROLogica
    {
        public long Folio { get; set; }
        public string Hora { get; set; }
        public DateTime Fecha { get; set; }
        public string Planta { get; set; }
        public string BinCode { get; set; }
        public string Item { get; set; }
        public string Descrip { get; set; }
        public string UM { get; set; }
        public double Cantidad { get; set; }
        public double Contador { get; set; }
        public double Diferencia { get; set; }

        public static int Guardar(BincontentCPROLogica bin)
        {
            string[] parametros = { "@Hora", "@Planta", "@Bincode", "@Item", "@Descrip", "@UM", "@Cantidad" };
            return AccesoDatos.ActualizarPRO("sp_mant_bincont", parametros, bin.Hora, bin.Planta, bin.BinCode, bin.Item, bin.Descrip, bin.UM, bin.Cantidad );
        }

        public static bool Verificar(BincontentCPROLogica bin)
        {
            try
            {
                string sQuery;
                sQuery = "SELECT * FROM t_bincont where planta = '"+bin.Planta+"' and hora = '" + bin.Hora + "' and cast(fecha as date) = cast('" + bin.Fecha + "' as date) ";
                DataTable datos = AccesoDatos.ConsultarPRO(sQuery);
                if (datos.Rows.Count != 0)
                    return true;
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }

        public static DataTable ConsultarLinea(BincontentCPROLogica bin)
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.ConsultarPRO("select distinct bincode from t_bincont where planta = '"+bin.Planta+"' and hora = '"+bin.Hora+"' and cast(fecha as date) = cast('"+bin.Fecha+"' as date) order by 1");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static DataTable BinContentLinea(BincontentCPROLogica bin)
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.ConsultarPRO("select folio,item,cantidad,contador,diferencia from t_bincont where planta = '" + bin.Planta + "' and bincode = '"+bin.BinCode+"' and hora = '" + bin.Hora + "' and cast(fecha as date) = cast('" + bin.Fecha + "' as date) order by 1");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static bool GuardaContador(BincontentCPROLogica bin)
        {
            try
            {
                DateTime dt = DateTime.Now;
                string sQuery = string.Empty;
                sQuery = "UPDATE t_bincont SET contador = " + bin.Contador + ", diferencia = " + bin.Diferencia + " WHERE folio = " + bin.Folio + " ";

                if (AccesoDatos.UpdatePRO(sQuery) != 0)
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
