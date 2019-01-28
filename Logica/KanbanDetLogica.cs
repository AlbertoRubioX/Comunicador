using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Datos;

namespace Logica
{
    public class KanbanDetLogica
    {
        public long Folio { get; set; }
        public int Consec { get; set; }
        public string Line { get; set; }
        public string RPO { get; set; }
        public string Item { get; set; }
        public DateTime Creation { get; set; }
        public DateTime Print { get; set; }
        public DateTime Register { get; set; }
        public DateTime Kanban { get; set; }
        public DateTime Start { get; set; }
        public double Quantity { get; set; }
        public double QtyFinish { get; set; }
        public double QtyShipped { get; set; }

        public static int Guardar(KanbanDetLogica kan)
        {
            string[] parametros = { "@Folio", "@Line", "@RPO", "@Creation", "@Item", "@Qty", "@Print", "@Register", "@Kanban", "@Start", "@QtyFinish", "@QtyShipped" };
            return AccesoDatos.ActualizarPRO("sp_mant_kanban_det", parametros, kan.Folio, kan.Line, kan.RPO, kan.Creation, kan.Item, kan.Quantity, kan.Print, kan.Register, kan.Kanban, kan.Start, kan.QtyFinish, kan.QtyShipped);
        }

        public static DataTable Consultar(KanbanDetLogica kan)
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.ConsultarPRO("SELECT * FROM t_kanban_det where folio = " + kan.Folio + "");

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }


        public static DataTable Listar(KanbanDetLogica kan)
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.ConsultarPRO("select * from t_kanban_det where folio = " + kan.Folio + " and consec = " + kan.Consec + "");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static bool Verificar(KanbanDetLogica kan)
        {
            try
            {
                string sQuery;
                sQuery = "SELECT * FROM t_kanban_det where line = '" + kan.Line + "' and rpo = '" + kan.RPO + "' ";
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

    }
}
