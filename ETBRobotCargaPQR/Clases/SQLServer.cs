using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace ETBRobotAsignarCasosPQR
{
    internal class SqlServer : IDisposable
    {
        private SqlConnection conn;
        private SqlCommand command;
        private readonly List<SqlParameter> args = new List<SqlParameter>();
        private string connectionString;

        public void Crear()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["ATLASConnectionDevelopment"].ToString();
                conn = new SqlConnection(connectionString);
                args.Clear();
            }
            catch (Exception ex)
            {
                throw new Exception("SqlServer.Crear(" + getLineErr(ex) + "): " + ex.Message);
            }
        }

        public void CrearAtlas(string SQLPointerString)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings[SQLPointerString].ToString();
                conn = new SqlConnection(connectionString);
                args.Clear(); 
            }
            catch (Exception ex)
            {
                throw new Exception("SqlServer.CrearAtlas(" + getLineErr(ex) + "): " + ex.Message);
            }
        }

        public void Destruir()
        {
            try
            {
                if (conn != null)
                {
                    if (conn.State != ConnectionState.Closed)
                        conn.Close();
                    conn = null;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("SqlServer.Destruir(" + getLineErr(ex) + "): " + ex.Message);
            }
        }

        public void AgregarParametro(string pName, Object pValue)
        {
            args.Add(new SqlParameter(pName, pValue));
        }

        public void AgregarParametroTabla(string pName, string pTable, Object pValue)
        {
            SqlParameter param = new SqlParameter
            {
                ParameterName = pName,
                Value = pValue,
                SqlDbType = SqlDbType.Structured,
                TypeName = pTable
            };

            args.Add(param);
        }

        public DataTable ObtenerTabla(string InstSQL)
        {
            SqlDataReader rs;
            DataTable dt = new DataTable();

            try
            {
                conn.Open();
                command = new SqlCommand(InstSQL, conn);
                rs = command.ExecuteReader();
                dt.Load(rs);
            }
            catch (Exception ex)
            {
                throw new Exception("SqlServer.ObtenerTabla(" + getLineErr(ex) + "): " + ex.Message);
            }
            finally
            {
                if (conn.State != ConnectionState.Closed)
                    conn.Close();
            }
            return dt;
        }

        public DataTable ObtenerTablaParam(string InstSQL)
        {
            SqlDataReader rs;
            DataTable dt = new DataTable();

            try
            {
                conn.Open();
                command = new SqlCommand(InstSQL, conn);

                for (int i = 0; i < args.Count; i++)
                {
                    command.Parameters.AddWithValue(args[i].ParameterName, args[i].Value);
                }

                rs = command.ExecuteReader();
                dt.Load(rs);
            }
            catch (Exception ex)
            {
                throw new Exception("SqlServer.ObtenerTabla(" + getLineErr(ex) + "): " + ex.Message);
            }
            finally
            {
                if (conn.State != ConnectionState.Closed)
                    conn.Close();
            }
            return dt;
        }

        public DataTable ObtenerTablaSP(string SPName)
        {
            SqlDataReader rs;
            DataTable dt = new DataTable();

            try
            {
                conn.Open();
                command = new SqlCommand(SPName, conn);
                command.CommandType = CommandType.StoredProcedure;
                for (int i = 0; i <= args.Count - 1; i++)
                    command.Parameters.AddWithValue(args[i].ParameterName, args[i].Value);
                rs = command.ExecuteReader();
                dt.Load(rs);
            }
            catch (Exception ex)
            {
                throw new Exception("SqlServer.ObtenerTablaSP(" + getLineErr(ex) + "): " + ex.Message);
            }
            finally
            {
                if (conn.State != ConnectionState.Closed)
                    conn.Close();
            }
            return dt;
        }

        public DataTable ObtenerTablaSPTablaParam(string SPName)
        {
            SqlDataReader rs;
            DataTable dt = new DataTable();

            try
            {
                conn.Open();
                command = new SqlCommand(SPName, conn);
                command.CommandType = CommandType.StoredProcedure;
                command.CommandTimeout = 3600;

                command.Parameters.AddRange(args.ToArray());

                rs = command.ExecuteReader();
                dt.Load(rs);
            }
            catch (Exception ex)
            {
                throw new Exception("SqlServer.ObtenerTablaSPTablaParam(" + getLineErr(ex) + "): " + ex.Message);
            }
            finally
            {
                if (conn.State != ConnectionState.Closed)
                    conn.Close();
            }
            return dt;
        }

        public void Ejecutar(string InstSQL)
        {
            SqlTransaction tran = null;

            try
            {
                conn.Open();
                tran = conn.BeginTransaction();
                command = new SqlCommand(InstSQL, conn, tran);
                command.CommandTimeout = 0;
                command.ExecuteNonQuery();
                tran.Commit();
            }
            catch (Exception ex)
            {
                if (tran != null)
                    tran.Rollback();
                throw new Exception("SqlServer.Ejecutar(" + getLineErr(ex) + "): " + ex.Message);
            }
            finally
            {
                if (conn.State != ConnectionState.Closed)
                    conn.Close();
            }
        }

        public void EjecutarParam(string InstSQL)
        {
            SqlTransaction tran = null;

            try
            {
                conn.Open();
                tran = conn.BeginTransaction();
                command = new SqlCommand(InstSQL, conn, tran);
                command.CommandTimeout = 0;

                for (int i = 0; i < args.Count; i++)
                {
                    command.Parameters.AddWithValue(args[i].ParameterName, args[i].Value);
                }

                command.ExecuteNonQuery();
                tran.Commit();
            }
            catch (Exception ex)
            {
                if (tran != null)
                    tran.Rollback();
                throw new Exception("SqlServer.Ejecutar(" + getLineErr(ex) + "): " + ex.Message);
            }
            finally
            {
                if (conn.State != ConnectionState.Closed)
                    conn.Close();
            }
        }

        public void EjecutarSP(string SPName)
        {
            SqlTransaction tran = null;

            try
            {
                conn.Open();
                tran = conn.BeginTransaction();
                command = new SqlCommand(SPName, conn, tran)
                {
                    CommandType = CommandType.StoredProcedure
                };
                for (int i = 0; i < args.Count; i++)
                {
                    command.Parameters.AddWithValue(args[i].ParameterName, args[i].Value);
                }
                command.ExecuteNonQuery();
                tran.Commit();
            }
            catch (Exception ex)
            {
                if (tran != null)
                    tran.Rollback();
                throw new Exception("SqlServer.EjecutarSP(" + getLineErr(ex) + "): " + ex.Message);
            }
            finally
            {
                if (conn.State != ConnectionState.Closed)
                    conn.Close();
            }
        }

        public static int getLineErr(Exception ex)
        {
            var lineNumber = 0;
            const string lineSearch = ":línea ";
            var index = ex.StackTrace.LastIndexOf(lineSearch);
            if (index != -1)
            {
                var lineNumberText = ex.StackTrace.Substring(index + lineSearch.Length);
                if (int.TryParse(lineNumberText, out lineNumber))
                {
                }
            }
            return lineNumber;
        }

        public void EliminarParametros() {
            args.Clear();

        }

        #region IDisposable Support

        private bool disposedValue = false; // Para detectar llamadas redundantes

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: elimine el estado administrado (objetos administrados).
                    command.Dispose();
                    conn.Dispose();
                }

                // TODO: libere los recursos no administrados (objetos no administrados) y reemplace el siguiente finalizador.
                // TODO: configure los campos grandes en nulos.

                disposedValue = true;
            }
        }

        // TODO: reemplace un finalizador solo si el anterior Dispose(bool disposing) tiene código para liberar los recursos no administrados.
        // ~SqlServer()
        // {
        //   // No cambie este código. Coloque el código de limpieza en el anterior Dispose(colocación de bool).
        //   Dispose(false);
        // }

        // Este código se agrega para implementar correctamente el patrón descartable.
        public void Dispose()
        {
            // No cambie este código. Coloque el código de limpieza en el anterior Dispose(colocación de bool).
            Dispose(true);
            // TODO: quite la marca de comentario de la siguiente línea si el finalizador se ha reemplazado antes.
            // GC.SuppressFinalize(this);
        }

        #endregion IDisposable Support
    }
}