using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace Datos
{
    public class AccesoDatos
    {
        public static int EjecutaSP(string as_procedimiento)
        {
            DataTable dt = new DataTable();
            SqlCommand _comando = MetodosDatos.CrearComandoSP(as_procedimiento);
            return MetodosDatos.EjecutaComando(_comando);
        }
        
        //reciben store procedure
        public static int Actualizar(string as_procedimiento, string[] nomparametros, params Object[] valparametros)
        {
            if (nomparametros.Length == valparametros.Length)
            {
                SqlCommand _comando = MetodosDatos.CrearComandoSP(as_procedimiento);
                int i = 0;
                foreach (string nomparam in nomparametros)
                    _comando.Parameters.AddWithValue(nomparam, ToDBNull(valparametros[i++]));

                return MetodosDatos.EjecutaComando(_comando);
            }
            return 0;
        }
        private static object ToDBNull(object value)
        {
            if (null != value)
            {
                Type t = value.GetType();
                if (t.Equals(typeof(DateTime)))
                {
                    if (value.Equals(DateTime.MinValue))
                        return DBNull.Value;
                    else
                        return value;
                }
                else
                    return value;
            }
            return DBNull.Value;
        }
        public static DataTable ConsultaSP(string as_procedimiento, string[] nomparametros, params object[] valparametros)
        {
            DataTable dt = new DataTable();
            if (nomparametros.Length == valparametros.Length)
            {
                SqlCommand _comando = MetodosDatos.CrearComandoSP(as_procedimiento);
                int i = 0;
                foreach (string nomparam in nomparametros)
                    _comando.Parameters.AddWithValue(nomparam, ToDBNull(valparametros[i++]));

                dt = MetodosDatos.EjecutaComandoSelect(_comando);
            }
            return dt;
        }

        
        //reciben querys
        public static DataTable Consultar(string as_query)
        {
            SqlCommand _comando = MetodosDatos.CrearComando();
            _comando.CommandText = as_query;
            return MetodosDatos.EjecutaComandoSelect(_comando);
        }

        public static int Borrar(string as_query)
        {
            SqlCommand _comando = MetodosDatos.CrearComando();
            _comando.CommandText = as_query;
            return MetodosDatos.EjecutaComando(_comando);
        }
        

        //agrega consecutivo
        public static int Consec(string as_proceso)
        {
            int _iFolio;
            SqlCommand _comando = MetodosDatos.CrearComandoSP("sp_consec");
            _comando.Parameters.AddWithValue("@Proceso", as_proceso);
            SqlParameter _Consec = new SqlParameter("@Folio", SqlDbType.Int);
            _Consec.Direction = ParameterDirection.Output;
            _comando.Parameters.Add(_Consec);
            MetodosDatos.EjecutaComando(_comando);
            _iFolio = Int32.Parse(_Consec.Value.ToString());
            return _iFolio;
        }

        #region regCloverPRo
        //actualizar cpro
        public static int ActualizarPRO(string as_procedimiento, string[] nomparametros, params Object[] valparametros)
        {
            if (nomparametros.Length == valparametros.Length)
            {
                SqlCommand _comando = MetodosDatos.CrearComandoSPPRO(as_procedimiento);
                int i = 0;
                foreach (string nomparam in nomparametros)
                    _comando.Parameters.AddWithValue(nomparam, ToDBNull(valparametros[i++]));

                return MetodosDatos.EjecutaComando(_comando);
            }
            return 0;
        }
        public static DataTable ConsultaSPPRO(string as_procedimiento, string[] nomparametros, params object[] valparametros)
        {
            DataTable dt = new DataTable();
            if (nomparametros.Length == valparametros.Length)
            {
                SqlCommand _comando = MetodosDatos.CrearComandoSPPRO(as_procedimiento);
                int i = 0;
                foreach (string nomparam in nomparametros)
                    _comando.Parameters.AddWithValue(nomparam, ToDBNull(valparametros[i++]));

                dt = MetodosDatos.EjecutaComandoSelect(_comando);
            }
            return dt;
        }
        //reciben querys cloverpro
        public static DataTable ConsultarPRO(string as_query)
        {
            SqlCommand _comando = MetodosDatos.CrearComandoPRO();
            _comando.CommandText = as_query;
            return MetodosDatos.EjecutaComandoSelectPRO(_comando);
        }

        public static int EjecutaSPCPRO(string as_procedimiento)
        {
            DataTable dt = new DataTable();
            SqlCommand _comando = MetodosDatos.CrearComandoSPPRO(as_procedimiento);
            return MetodosDatos.EjecutaComando(_comando);
        }

        public static int UpdatePRO(string as_query)
        {
            SqlCommand _comando = MetodosDatos.CrearComandoPRO();
            _comando.CommandText = as_query;
            return MetodosDatos.EjecutaComando(_comando);
        }
        #endregion
        #region regCloverRH
        //actualizar cpro
        public static int ActualizarCRH(string as_procedimiento, string[] nomparametros, params Object[] valparametros)
        {
            if (nomparametros.Length == valparametros.Length)
            {
                SqlCommand _comando = MetodosDatos.CrearComandoSPCRH(as_procedimiento);
                int i = 0;
                foreach (string nomparam in nomparametros)
                    _comando.Parameters.AddWithValue(nomparam, ToDBNull(valparametros[i++]));

                return MetodosDatos.EjecutaComando(_comando);
            }
            return 0;
        }
        public static DataTable ConsultaSPCRH(string as_procedimiento, string[] nomparametros, params object[] valparametros)
        {
            DataTable dt = new DataTable();
            if (nomparametros.Length == valparametros.Length)
            {
                SqlCommand _comando = MetodosDatos.CrearComandoSPCRH(as_procedimiento);
                int i = 0;
                foreach (string nomparam in nomparametros)
                    _comando.Parameters.AddWithValue(nomparam, ToDBNull(valparametros[i++]));

                dt = MetodosDatos.EjecutaComandoSelectCRH(_comando);
            }
            return dt;
        }
        //reciben querys cloverpro
        public static DataTable ConsultarCRH(string as_query)
        {
            SqlCommand _comando = MetodosDatos.CrearComandoCRH();
            _comando.CommandText = as_query;
            return MetodosDatos.EjecutaComandoSelectCRH(_comando);
        }

        public static int EjecutaSPCRH(string as_procedimiento)
        {
            DataTable dt = new DataTable();
            SqlCommand _comando = MetodosDatos.CrearComandoSPCRH(as_procedimiento);
            return MetodosDatos.EjecutaComando(_comando);
        }

        public static int UpdateCRH(string as_query)
        {
            SqlCommand _comando = MetodosDatos.CrearComandoCRH();
            _comando.CommandText = as_query;
            return MetodosDatos.EjecutaComando(_comando);
        }
        #endregion
    }
}