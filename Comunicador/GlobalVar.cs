using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Comunicador
{
    public class GlobalVar
    {
        public static string gsOrigen;

        public static string TurnoGlobal()
        {

            string sTurno = "2";
            DateTime dtFecha = DateTime.Now;
            if (dtFecha.Hour >= 6 && dtFecha.Hour <= 16)
            {
                sTurno = "1";
            }

            return sTurno;
        }
    }
}
