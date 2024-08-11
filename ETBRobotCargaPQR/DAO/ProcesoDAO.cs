using Microsoft.VisualBasic;
using ETBRobotAsignarCasosPQR.Clases;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using ETBRobotCargaPQR.Clases;

namespace ETBRobotAsignarCasosPQR.DAO
{
    internal class ProcesoDAO
    {
        private readonly SqlServer conn;

        public ProcesoDAO()
        {
            conn = new SqlServer();
        }

        public Config GetConfig(string empresa)
        {
            try
            {
                Config config = new Config();

                conn.Crear();
                conn.AgregarParametro("@p_empresa", empresa);
                //conn.AgregarParametro("@p_tipoconsulta", tipoconsulta);

                DataTable dt = conn.ObtenerTablaParam("SELECT [URL], [UserBOE], [PassBOE], [FileName] " +
                    "FROM [JOSE].[dbo].[ConfiguracionIngresos] " +
                    "WHERE [Empresa] = @p_empresa AND [Estado] = 1;");

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow data in dt.Rows)
                    {
                        config = new Config
                        {
                            url = data["URL"].ToString(),
                            userboe = data["UserBOE"].ToString(),
                            passboe = data["PassBOE"].ToString(),
                            filename = data["FileName"].ToString()
                        };
                    }
                }
                else
                {
                    throw new Exception("Configuration not found.");
                }

                conn.Destruir();
                return config;
            }
            catch
            {
                throw;
            }
        }

        public ETB_RobotAsignacionMasiva_Insert_reasignaciones Get_Insert_reasignaciones()
        {
            try
            {
                ETB_RobotAsignacionMasiva_Insert_reasignaciones config = new ETB_RobotAsignacionMasiva_Insert_reasignaciones();

                conn.Crear();
                //conn.AgregarParametro("@numero_pqr", numero_pqr);
                //conn.AgregarParametro("@p_tipoconsulta", tipoconsulta);

                DataTable dt = conn.ObtenerTablaParam("SELECT [usuario_asignado], [usuario_final], [numero_pqr], [Altura] " +
                    "FROM [Atlas].[dbo].[ETB_RobotAsignacionMasiva_Insert_reasignaciones] " );

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow data in dt.Rows)
                    {
                        config = new ETB_RobotAsignacionMasiva_Insert_reasignaciones
                        {
                            usuario_asignado = data["usuario_asignado"].ToString(),
                            usuario_final = data["usuario_final"].ToString(),
                            numero_pqr = data["numero_pqr"].ToString(),
                            Altura = data["Altura"].ToString()
                        };
                    }
                }
                else
                {
                    throw new Exception("Configuration not found.");
                }

                conn.Destruir();
                return config;
            }
            catch
            {
                throw;
            }
        }

        public RegistroEjecucion GetRegistroEjecucion(string empresa,  string Winuser, string enviroment)
        {
            try
            {
                RegistroEjecucion registroEjecucion = new RegistroEjecucion();

                conn.Crear();
                conn.AgregarParametro("@p_empresa", empresa);
                //conn.AgregarParametro("@p_tipoconsulta", tipoconsulta);
                //conn.AgregarParametro("@p_fecha", fecha);
                conn.AgregarParametro("@p_winuser", Winuser);
                conn.AgregarParametro("@enviroment", enviroment);

                DataTable dt = conn.ObtenerTablaSP("[Camp_Data].[dbo].[sp_ETB_RobotAsignacionMasiva_ConsultarRobot]");

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dtr in dt.Rows)
                    {
                        registroEjecucion = new RegistroEjecucion
                        {
                            id = string.IsNullOrWhiteSpace(dtr["ID"].ToString()) ? 0 : Convert.ToInt32(dtr["ID"].ToString()),
                            intentos = string.IsNullOrWhiteSpace(dtr["Intentos"].ToString()) ? 0 : Convert.ToInt32(dtr["Intentos"].ToString()),
                            empresa = empresa,
                            //tipo = tipoconsulta,
                            //fechaproceso = fecha
                        };
                    }
                }

                conn.Destruir();
                return registroEjecucion;
            }
            catch
            {
                throw;
            }
        }

        public void ReiniciarBot(string empresa)
        {
            try
            {
                conn.Crear();
                //conn.AgregarParametro("@p_maquina", empresa + "-" + tipoconsulta);
                conn.EjecutarSP("[Camp_Data].[dbo].[sp_ETB_RobotCargaPQR_ReiniciarBot]");
                conn.Destruir();
            }
            catch
            {
                throw;
            }
        }

        public void ColaHumanaBot(RegistroEjecucion registroEjecucion)
        {
            try
            {
                conn.Crear();
                conn.AgregarParametro("@p_id", registroEjecucion.id);
                conn.AgregarParametro("@p_msj", registroEjecucion.observaciones);
                conn.AgregarParametro("@p_maquina", registroEjecucion.empresa + "-" + registroEjecucion.tipo);

                conn.EjecutarSP("[Camp_Data].[dbo].[sp_ETB_RobotCargaPQR_ColaHumanaBot]");
                conn.Destruir();
            }
            catch
            {
                throw;
            }
        }

        public void InsertnewPasswordOnRegister(string LastPassword, string NewpasswordString)
        {
            try
            {
                conn.Crear();
                conn.AgregarParametro("@NewpasswordString", NewpasswordString);
                conn.AgregarParametro("@LastPassword", LastPassword);
                conn.EjecutarSP("[Camp_Data].[dbo].[sp_ETB_RobotCargaPQR_ChangePassword]");
                conn.Destruir();
            }
            catch
            {
                throw;
            }
        }

        public void GuardarRegistroEjecucion(RegistroEjecucion registroEjecucion)
        {
            try
            {
                conn.Crear();
                conn.AgregarParametro("@p_id", registroEjecucion.id);
                conn.AgregarParametro("@p_procesado", registroEjecucion.procesado);
                conn.AgregarParametro("@p_observaciones", registroEjecucion.observaciones);

                conn.EjecutarSP("[Camp_Data].[dbo].[sp_ETB_RobotCargaPQR_GuardarProceso]");
                conn.Destruir();
            }
            catch
            {
                throw;
            }
        }

        public void GuardarRegistroLogueo(SqlConnection conn, SqlTransaction trans, RegistroEjecucion registroEjecucion, string login_genesys, string fecha_login,
                                          string hora_login, string fecha_logout, string hora_logout, string extension, string login_aib, string identificacion)
        {
            SqlCommand cmd = new SqlCommand();

            try
            {
                cmd = new SqlCommand
                {
                    Connection = conn,
                    Transaction = trans,
                    CommandType = CommandType.StoredProcedure,
                    CommandText = "[Camp_Data].[dbo].[sp_MovistarRobotETL_CargaLogins]"
                };

                cmd.Parameters.AddWithValue("@fecha_carga", registroEjecucion.fechaproceso.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                cmd.Parameters.AddWithValue("@login_genesys", login_genesys);
                cmd.Parameters.AddWithValue("@fecha_login", fecha_login);
                cmd.Parameters.AddWithValue("@hora_login", hora_login);
                cmd.Parameters.AddWithValue("@fecha_logout", fecha_logout);
                cmd.Parameters.AddWithValue("@hora_logout", hora_logout);
                cmd.Parameters.AddWithValue("@extension", extension);
                cmd.Parameters.AddWithValue("@login_aib", login_aib);
                cmd.Parameters.AddWithValue("@identificacion", identificacion);

                cmd.ExecuteNonQuery();
            }
            catch
            {
                throw;
            }
            finally
            {
                cmd.Dispose();
            }
        }

        public void GetEmpleado(string login_genesys, out string login_aib, out string identificacion)
        {
            login_aib = "";
            identificacion = "";

            try
            {
                conn.Crear();
                conn.AgregarParametro("@login_genesys", login_genesys);
                DataTable dt = conn.ObtenerTablaParam(@"SELECT 		  l.login_aib, e.Identificacion
                                                        FROM          [Camp_Data].[dbo].[Movistar_RelacionLogins] l
                                                        LEFT JOIN     [MIS].[dbo].[NM_Empleados] e ON e.[LoginAIB] COLLATE Latin1_General_CI_AS = l.[login_aib]
                                                        WHERE         l.[login_genesys] = @login_genesys
                                                        AND           l.[estado] = 1; ");
                conn.Destruir();

                if (dt.Rows.Count > 0)
                {
                    login_aib = dt.Rows[0]["login_aib"].ToString();
                    identificacion = dt.Rows[0]["Identificacion"].ToString();
                }
            }
            catch
            {
                throw;
            }
        }

        public void CalcularTiempo(RegistroEjecucion registroEjecucion, string login_genesys, string fecha_login, string hora_login,
                                   string fecha_logout, string hora_logout, string extension, string login_aib, string identificacion,
                                   ref DataTable dt_logueo)
        {
            try
            {
                DateTime FechaHoraInicial = Convert.ToDateTime(fecha_login.Substring(6, 4) + "-" + fecha_login.Substring(3, 2) + "-" + fecha_login.Substring(0, 2) + " " + hora_login);
                DateTime FechaHoraFinal = Convert.ToDateTime(fecha_logout.Substring(6, 4) + "-" + fecha_logout.Substring(3, 2) + "-" + fecha_logout.Substring(0, 2) + " " + hora_logout);

                DateTime FechaRango1 = Convert.ToDateTime(registroEjecucion.fechaproceso.ToString("yyyy-MM-dd 00:00:00"));
                DateTime FechaRango2 = FechaRango1.AddMinutes(30);

                long TiempoRango = 0;

                if (FechaHoraInicial.ToString("yyyy-MM-dd") == registroEjecucion.fechaproceso.AddDays(-1).ToString("yyyy-MM-dd"))
                {
                    FechaHoraInicial = registroEjecucion.fechaproceso;
                }

                while (Convert.ToDateTime(FechaRango1.ToString("yyyy-MM-dd")) <= Convert.ToDateTime(FechaHoraFinal.ToString("yyyy-MM-dd")))
                {
                    DateTime FechaRangoActual = FechaRango1;

                    if ((FechaHoraInicial >= FechaRango1) && (FechaHoraInicial <= FechaRango2))
                    {
                        if ((FechaHoraFinal >= FechaRango1) && (FechaHoraFinal <= FechaRango2))
                        {
                            TiempoRango += Math.Abs(DateAndTime.DateDiff(DateInterval.Second, FechaHoraInicial, FechaHoraFinal));
                        }
                        else
                        {
                            TiempoRango += Math.Abs(DateAndTime.DateDiff(DateInterval.Second, FechaHoraInicial, FechaRango2));
                            FechaHoraInicial = FechaRango2.AddSeconds(0);
                        }
                    }

                    if (TiempoRango != 0)
                    {
                        DataRow datarow = dt_logueo.NewRow();

                        datarow["logid_gen"] = login_genesys;
                        datarow["logid"] = login_aib;
                        datarow["cedula"] = identificacion;
                        datarow["fechahorarango"] = FechaRangoActual;
                        datarow["tiempo"] = TiempoRango;
                        datarow["extension"] = long.Parse(extension.Substring(1));
                        datarow["Aux0"] = 0;
                        datarow["Aux1"] = 0;
                        datarow["Aux2"] = 0;
                        datarow["Aux3"] = 0;
                        datarow["Aux4"] = 0;
                        datarow["Aux5"] = 0;
                        datarow["Aux6"] = 0;
                        datarow["Aux7"] = 0;
                        datarow["Aux8"] = 0;
                        datarow["Aux9"] = 0;

                        dt_logueo.Rows.Add(datarow);
                        TiempoRango = 0;
                    }

                    FechaRango1 = FechaRango2;
                    FechaRango2 = FechaRango2.AddMinutes(30);
                }
            }
            catch
            {
                throw;
            }
        }

        public void GuardarProceso(SqlConnection conn, SqlTransaction trans, RegistroEjecucion registroEjecucion, DataTable dt_logueo, out bool continuarProceso)
        {
            continuarProceso = true;

            try
            {
                IEnumerable<ItemLogueo> listProceso = dt_logueo
                    .AsEnumerable()
                    .GroupBy(r => new { Logid = r.Field<int>("logid"), Cedula = r.Field<string>("cedula"), FechaHoraRango = r.Field<DateTime>("fechahorarango") })
                    .Select(g => new ItemLogueo()
                    {
                        Logid = g.Key.Logid,
                        Cedula = g.Key.Cedula,
                        FechaHoraRango = g.Key.FechaHoraRango,
                        Extension = g.Max(x => x.Field<long>("extension")),
                        Tiempo = g.Sum(x => x.Field<int>("tiempo")),
                        Aux0 = g.Sum(x => x.Field<int>("Aux0")),
                        Aux1 = g.Sum(x => x.Field<int>("Aux1")),
                        Aux2 = g.Sum(x => x.Field<int>("Aux2")),
                        Aux3 = g.Sum(x => x.Field<int>("Aux3")),
                        Aux4 = g.Sum(x => x.Field<int>("Aux4")),
                        Aux5 = g.Sum(x => x.Field<int>("Aux5")),
                        Aux6 = g.Sum(x => x.Field<int>("Aux6")),
                        Aux7 = g.Sum(x => x.Field<int>("Aux7")),
                        Aux8 = g.Sum(x => x.Field<int>("Aux8")),
                        Aux9 = g.Sum(x => x.Field<int>("Aux9"))
                    })
                    .ToList();

                string sql = "";

                SqlCommand cmd = new SqlCommand
                {
                    Connection = conn,
                    Transaction = trans,
                    CommandType = CommandType.Text,
                    CommandTimeout = TimeSpan.FromMinutes(5).Seconds
                };

                if (registroEjecucion.empresa.Contains("Metrotel"))
                {
                    // Log("OK", "Limpiar tabla temporal")
                    sql += "TRUNCATE TABLE [CMS13_TMP].[dbo].[hlogueoAIB_temp_metrotel]; ";

                    // Log("OK", "Limpiar tabla Nom_CMS13")
                    sql += "DELETE FROM [RH].[dbo].[Nom_CMS13] WHERE (CAST(row_date AS DATE) = '" + registroEjecucion.fechaproceso.ToString("yyyy-MM-dd") + "') " +
                           "AND logid IN (SELECT logid FROM [RH].[dbo].[Datos_ListadoTodosEmpleados_ConFecRetiro] datos WHERE datos.[GLBDescripcionCentroCostos] LIKE '%Metrotel%'); ";

                    // Log("OK", "Insertar registros")
                    foreach (ItemLogueo item in listProceso as IEnumerable<ItemLogueo>)
                    {
                        sql += "INSERT INTO [CMS13_TMP].[dbo].[hlogueoAIB_temp_metrotel] (logid, cedula, extension, fechahorarango, tiempo, aux0, aux1, aux2, aux3, aux4, aux5, aux6, aux7, aux8, aux9) " +
                               "VALUES ('" + item.Logid.ToString() + "', '" + item.Cedula + "', '" + item.Extension.ToString() + "', '" +
                               Convert.ToDateTime(item.FechaHoraRango).ToString("yyyy-MM-dd HH:mm:ss") + "', '" + item.Tiempo + "', '" +
                               item.Aux0 + "', '" + item.Aux1 + "', '" + item.Aux2 + "', '" + item.Aux3 + "', '" + item.Aux4 + "', '" +
                               item.Aux5 + "', '" + item.Aux6 + "', '" + item.Aux7 + "', '" + item.Aux8 + "', '" + item.Aux9 + "'); ";

                        sql += "UPDATE temp SET " +
                               "temp.[Aux1] = aux.[Tiempo  Descanso], " +
                               "temp.[Aux2] = aux.[Tiempo  WC], " +
                               "temp.[Aux4] = aux.[Tiempo  Reuniones_de_Grupo], " +
                               "temp.[Aux5] = aux.[Tiempo  Llamadas_Salientes], " +
                               "temp.[Aux6] = aux.[Tiempo  Retroalimentacion], " +
                               "temp.[Aux7] = aux.[Tiempo  Tarea_Especial], " +
                               "temp.[Aux9] = aux.[Tiempo  Dano_Equipo] " +
                               "FROM [CMS13_TMP].[dbo].[hlogueoAIB_temp_metrotel] AS temp " +
                               "INNER JOIN [Camp_Data].[dbo].[Movistar_RelacionLogins] AS login ON login.[login_aib] = temp.[logid] AND login.[estado] = 1 " +
                               "INNER JOIN [CMS13_TMP].[dbo].[Genesys_hagent_aux] AS aux ON aux.[Agente] = login.[login_genesys] AND CAST(temp.[fechahorarango] AS DATE) = CAST(aux.[Dia] AS DATE) " +
                               "AND aux.[Hora] = SUBSTRING(CONVERT(CHAR(5), temp.[fechahorarango], 108), 1, 5) + '-' + SUBSTRING(CONVERT(CHAR(5), DATEADD(MINUTE, 30, temp.[fechahorarango]), 108), 1, 5); ";
                    }

                    // Log("OK", "Limpiar tablas hagent-dagent")
                    sql += "DELETE FROM [CMS13_TMP].[dbo].[hagent] WHERE [split] IN (252) AND [row_date] = '" + registroEjecucion.fechaproceso.ToString("yyyy-MM-dd") + "'; ";
                    sql += "DELETE FROM [CMS13_TMP].[dbo].[dagent] WHERE [split] IN (252) AND [row_date] = '" + registroEjecucion.fechaproceso.ToString("yyyy-MM-dd") + "'; ";

                    // Log("OK", "Insertar en hagent")
                    sql += "INSERT INTO [CMS13_TMP].[dbo].[hagent] (row_date, starttime, intrvl, acd, split, extension, logid, loc_id, rsv_level, " +
                           "i_stafftime, ti_stafftime, ti_auxtime, ti_auxtime0, ti_auxtime1, ti_auxtime2, ti_auxtime3, ti_auxtime4, ti_auxtime5, ti_auxtime6, ti_auxtime7, ti_auxtime8, ti_auxtime9) " +
                           "SELECT CONVERT(DATE, [fechahorarango]), [CMS13_TMP].dbo.hcms2([fechahorarango]), 30, 7, 252, SUBSTRING([extension], 1, 7), [logid], 1, 0, ISNULL([tiempo], 0), ISNULL([tiempo], 0), " +
                           "(ISNULL([aux0], 0) + ISNULL([aux1], 0) + ISNULL([aux2], 0) + ISNULL([aux3], 0) + ISNULL([aux4], 0) + ISNULL([aux5], 0) + ISNULL([aux6], 0) + ISNULL([aux7], 0) + ISNULL([aux8], 0) + ISNULL([aux9], 0)), " +
                           "ISNULL([aux0], 0), ISNULL([aux1], 0), ISNULL([aux2], 0), ISNULL([aux3], 0), ISNULL([aux4], 0), ISNULL([aux5], 0), ISNULL([aux6], 0), ISNULL([aux7], 0), ISNULL([aux8], 0), ISNULL([aux9], 0) " +
                           "FROM [CMS13_TMP].[dbo].[hlogueoAIB_temp_metrotel]; ";

                    // Log("OK", "Insertar en dagent")
                    sql += "INSERT INTO [CMS13_TMP].[dbo].[dagent] ([row_date], [acd], [split], [extension], [logid], [loc_id], [rsv_level], " +
                           "[i_stafftime], [ti_stafftime], [ti_auxtime], [i_auxtime], [ti_auxtime0], [ti_auxtime1], [ti_auxtime2], [ti_auxtime3], " +
                           "[ti_auxtime4], [ti_auxtime5], [ti_auxtime6], [ti_auxtime7], [ti_auxtime8], [ti_auxtime9]) " +
                           "SELECT CONVERT(DATE, [fechahorarango]), 7, 252, 99999, [logid], 1, 0, ISNULL(SUM([tiempo]), 0), ISNULL(SUM([tiempo]), 0), " +
                           "SUM(ISNULL([aux0], 0) + ISNULL([aux1], 0) + ISNULL([aux2], 0) + ISNULL([aux3], 0) + ISNULL([aux4], 0) + ISNULL([aux5], 0) + ISNULL([aux6], 0) + ISNULL([aux7], 0) + ISNULL([aux8], 0) + ISNULL([aux9], 0)), " +
                           "SUM(ISNULL([aux0], 0) + ISNULL([aux1], 0) + ISNULL([aux2], 0) + ISNULL([aux3], 0) + ISNULL([aux4], 0) + ISNULL([aux5], 0) + ISNULL([aux6], 0) + ISNULL([aux7], 0) + ISNULL([aux8], 0) + ISNULL([aux9], 0)), " +
                           "ISNULL(SUM([aux0]), 0), ISNULL(SUM([aux1]), 0), ISNULL(SUM([aux2]), 0), ISNULL(SUM([aux3]), 0), ISNULL(SUM([aux4]), 0), ISNULL(SUM([aux5]), 0), ISNULL(SUM([aux6]), 0), ISNULL(SUM([aux7]), 0), " +
                           "ISNULL(SUM([aux8]), 0), ISNULL(SUM([aux9]), 0) FROM [CMS13_TMP].[dbo].[hlogueoAIB_temp_metrotel] GROUP BY [logid], CONVERT(DATE, [fechahorarango]); ";

                    // Log("OK", "Guardar tabla Nom_CMS13")
                    sql += @"INSERT INTO [RH].[dbo].[Nom_CMS13] " +
                            "SELECT SUM(Staffed) AS LOGUEADO, SUM(HORAS_NOCTURNAS) AS HORAS_NOCTURNAS, SUM(HORAS_DIURNAS) AS HORAS_DIURNAS, logid, row_date, MIN(STARTTIME) AS hcms_min, MAX(STARTTIME) AS hcms_max " +
                            "FROM (SELECT SUM(ti_stafftime) / 3600.0 AS Staffed, SUM(ti_auxtime) AS Aux_Time, row_date, logid, [CMS13_TMP].dbo.HINC2(DATEPART(HOUR, row_date + [CMS13_TMP].dbo.HCMS(starttime))) AS starttime, " +
                            "CASE WHEN ([CMS13_TMP].dbo.hinc2(DATEPART([HOUR], row_date + [CMS13_TMP].dbo.HCMS(starttime))) >= 2100 AND [CMS13_TMP].dbo.hinc2(DATEPART([HOUR], row_date + [CMS13_TMP].dbo.HCMS(starttime))) <= 2300) " +
                            "OR ([CMS13_TMP].dbo.hinc2(DATEPART([HOUR], row_date + [CMS13_TMP].dbo.HCMS(starttime))) >= 0 AND [CMS13_TMP].dbo.hinc2(DATEPART([HOUR], row_date + [CMS13_TMP].dbo.HCMS(starttime)) ) <= 500) THEN " +
                            "SUM(ti_stafftime) / 3600.0 ELSE 0 END AS HORAS_NOCTURNAS, " +
                            "CASE WHEN([CMS13_TMP].dbo.hinc2(DATEPART([HOUR], row_date + [CMS13_TMP].[dbo].HCMS(starttime))) >= 600 AND [CMS13_TMP].dbo.hinc2(DATEPART([HOUR], row_date + [CMS13_TMP].dbo.HCMS(starttime)) ) <= 2000 ) THEN " +
                            "SUM(ti_stafftime) / 3600.0 ELSE 0 END AS HORAS_DIURNAS " +
                            "FROM [CMS13_TMP].dbo.hagent WHERE (CAST(row_date AS DATE) = '" + registroEjecucion.fechaproceso.ToString("yyyy-MM-dd") + "') " +
                            "AND split <> 2000 AND logid IN (SELECT logid FROM [RH].[dbo].[Datos_ListadoTodosEmpleados_ConFecRetiro] datos WHERE datos.[GLBDescripcionCentroCostos] LIKE '%Metrotel%') " +
                            "GROUP BY row_date, logid, [CMS13_TMP].dbo.HINC2(DATEPART(HOUR, row_date + [CMS13_TMP].dbo.HCMS(starttime)))) Agent_STAFFINGTIME_HORA GROUP BY logid, row_date; ";
                }

                if (registroEjecucion.empresa.Contains("Telebucaramanga"))
                {
                    // Log("OK", "Limpiar tabla temporal")
                    sql += "TRUNCATE TABLE [CMS13_TMP].[dbo].[hlogueoAIB_temp_telebucaramanga]; ";

                    // Log("OK", "Limpiar tabla Nom_CMS13")
                    sql += "DELETE FROM [RH].[dbo].[Nom_CMS13] WHERE (CAST(row_date AS DATE) = '" + registroEjecucion.fechaproceso.ToString("yyyy-MM-dd") + "') " +
                           "AND logid IN (SELECT logid FROM [RH].[dbo].[Datos_ListadoTodosEmpleados_ConFecRetiro] datos WHERE datos.[GLBDescripcionCentroCostos] LIKE '%Telebucaramanga%'); ";

                    // Log("OK", "Insertar registros")
                    foreach (ItemLogueo item in listProceso as IEnumerable<ItemLogueo>)
                    {
                        sql += "INSERT INTO [CMS13_TMP].[dbo].[hlogueoAIB_temp_telebucaramanga] (logid, cedula, extension, fechahorarango, tiempo, aux0, aux1, aux2, aux3, aux4, aux5, aux6, aux7, aux8, aux9) " +
                               "VALUES ('" + item.Logid.ToString() + "', '" + item.Cedula + "', '" + item.Extension.ToString() + "', '" +
                               Convert.ToDateTime(item.FechaHoraRango).ToString("yyyy-MM-dd HH:mm:ss") + "', '" + item.Tiempo + "', '" +
                               item.Aux0 + "', '" + item.Aux1 + "', '" + item.Aux2 + "', '" + item.Aux3 + "', '" + item.Aux4 + "', '" +
                               item.Aux5 + "', '" + item.Aux6 + "', '" + item.Aux7 + "', '" + item.Aux8 + "', '" + item.Aux9 + "'); ";

                        sql += "UPDATE temp SET " +
                               "temp.[Aux1] = aux.[Tiempo  Descanso], " +
                               "temp.[Aux2] = aux.[Tiempo  WC], " +
                               "temp.[Aux4] = aux.[Tiempo  Reuniones_de_Grupo], " +
                               "temp.[Aux5] = aux.[Tiempo  Llamadas_Salientes], " +
                               "temp.[Aux6] = aux.[Tiempo  Retroalimentacion], " +
                               "temp.[Aux7] = aux.[Tiempo  Tarea_Especial], " +
                               "temp.[Aux9] = aux.[Tiempo  Dano_Equipo] " +
                               "FROM [CMS13_TMP].[dbo].[hlogueoAIB_temp_telebucaramanga] AS temp " +
                               "INNER JOIN [Camp_Data].[dbo].[Movistar_RelacionLogins] AS login ON login.[login_aib] = temp.[logid] AND login.[estado] = 1 " +
                               "INNER JOIN [CMS13_TMP].[dbo].[Genesys_hagent_aux] AS aux ON aux.[Agente] = login.[login_genesys] AND CAST(temp.[fechahorarango] AS DATE) = CAST(aux.[Dia] AS DATE) " +
                               "AND aux.[Hora] = SUBSTRING(CONVERT(CHAR(5), temp.[fechahorarango], 108), 1, 5) + '-' + SUBSTRING(CONVERT(CHAR(5), DATEADD(MINUTE, 30, temp.[fechahorarango]), 108), 1, 5); ";
                    }

                    // Log("OK", "Limpiar tablas hagent-dagent")
                    sql += "DELETE FROM [CMS13_TMP].[dbo].[hagent] WHERE [split] IN (235, 236, 237, 322, 323) AND [row_date] = '" + registroEjecucion.fechaproceso.ToString("yyyy-MM-dd") + "'; ";
                    sql += "DELETE FROM [CMS13_TMP].[dbo].[dagent] WHERE [split] IN (235, 236, 237, 322, 323) AND [row_date] = '" + registroEjecucion.fechaproceso.ToString("yyyy-MM-dd") + "'; ";

                    // Log("OK", "Insertar en hagent")
                    sql += "INSERT INTO [CMS13_TMP].[dbo].[hagent] (row_date, starttime, intrvl, acd, split, extension, logid, loc_id, rsv_level, " +
                           "i_stafftime, ti_stafftime, ti_auxtime, ti_auxtime0, ti_auxtime1, ti_auxtime2, ti_auxtime3, ti_auxtime4, ti_auxtime5, ti_auxtime6, ti_auxtime7, ti_auxtime8, ti_auxtime9) " +
                           "SELECT CONVERT(DATE, [fechahorarango]), [CMS13_TMP].[dbo].hcms2([fechahorarango]), 30, 7, 322, SUBSTRING([extension], 1, 7), [logid], 1, 0, ISNULL([tiempo], 0), ISNULL([tiempo], 0), " +
                           "(ISNULL([aux0], 0) + ISNULL([aux1], 0) + ISNULL([aux2], 0) + ISNULL([aux3], 0) + ISNULL([aux4], 0) + ISNULL([aux5], 0) + ISNULL([aux6], 0) + ISNULL([aux7], 0) + ISNULL([aux8], 0) + ISNULL([aux9], 0)), " +
                           "ISNULL([aux0], 0), ISNULL([aux1], 0), ISNULL([aux2], 0), ISNULL([aux3], 0), ISNULL([aux4], 0), ISNULL([aux5], 0), ISNULL([aux6], 0), ISNULL([aux7], 0), ISNULL([aux8], 0), ISNULL([aux9], 0) " +
                           "FROM [CMS13_TMP].[dbo].[hlogueoAIB_temp_telebucaramanga]; ";

                    // Log("OK", "Insertar en dagent")
                    sql += "INSERT INTO [CMS13_TMP].[dbo].[dagent] ([row_date], [acd], [split], [extension], [logid], [loc_id], [rsv_level], " +
                           "[i_stafftime], [ti_stafftime], [ti_auxtime], [i_auxtime], [ti_auxtime0], [ti_auxtime1], [ti_auxtime2], [ti_auxtime3], " +
                           "[ti_auxtime4], [ti_auxtime5], [ti_auxtime6], [ti_auxtime7], [ti_auxtime8], [ti_auxtime9]) " +
                           "SELECT CONVERT(DATE, [fechahorarango]), 7, 322, 99999, [logid], 1, 0, ISNULL(SUM([tiempo]), 0), ISNULL(SUM([tiempo]), 0), " +
                           "SUM(ISNULL([aux0], 0) + ISNULL([aux1], 0) + ISNULL([aux2], 0) + ISNULL([aux3], 0) + ISNULL([aux4], 0) + ISNULL([aux5], 0) + ISNULL([aux6], 0) + ISNULL([aux7], 0) + ISNULL([aux8], 0) + ISNULL([aux9], 0)), " +
                           "SUM(ISNULL([aux0], 0) + ISNULL([aux1], 0) + ISNULL([aux2], 0) + ISNULL([aux3], 0) + ISNULL([aux4], 0) + ISNULL([aux5], 0) + ISNULL([aux6], 0) + ISNULL([aux7], 0) + ISNULL([aux8], 0) + ISNULL([aux9], 0)), " +
                           "ISNULL(SUM([aux0]), 0), ISNULL(SUM([aux1]), 0), ISNULL(SUM([aux2]), 0), ISNULL(SUM([aux3]), 0), ISNULL(SUM([aux4]), 0), ISNULL(SUM([aux5]), 0), ISNULL(SUM([aux6]), 0), ISNULL(SUM([aux7]), 0), " +
                           "ISNULL(SUM([aux8]), 0), ISNULL(SUM([aux9]), 0) FROM [CMS13_TMP].[dbo].[hlogueoAIB_temp_telebucaramanga] GROUP BY [logid], CONVERT(DATE, [fechahorarango]); ";

                    // Log("OK", "Guardar tabla Nom_CMS13")
                    sql += "INSERT INTO [RH].[dbo].[Nom_CMS13] " +
                            "SELECT SUM(Staffed) AS LOGUEADO, SUM(HORAS_NOCTURNAS) AS HORAS_NOCTURNAS, SUM(HORAS_DIURNAS) AS HORAS_DIURNAS, logid, row_date, MIN(STARTTIME) AS hcms_min, MAX(STARTTIME) AS hcms_max " +
                            "FROM (SELECT SUM(ti_stafftime) / 3600.0 AS Staffed, SUM(ti_auxtime) AS Aux_Time, row_date, logid, [CMS13_TMP].dbo.HINC2(DATEPART(HOUR, row_date + [CMS13_TMP].dbo.HCMS(starttime))) AS starttime, " +
                            "CASE WHEN ([CMS13_TMP].dbo.hinc2(DATEPART([HOUR], row_date + [CMS13_TMP].dbo.HCMS(starttime))) >= 2100 AND [CMS13_TMP].[dbo].hinc2(DATEPART([HOUR], row_date + [CMS13_TMP].dbo.HCMS(starttime))) <= 2300) " +
                            "OR ([CMS13_TMP].dbo.hinc2(DATEPART([HOUR], row_date + [CMS13_TMP].dbo.HCMS(starttime))) >= 0 AND [CMS13_TMP].dbo.hinc2(DATEPART([HOUR], row_date + [CMS13_TMP].dbo.HCMS(starttime)) ) <= 500) THEN " +
                            "SUM(ti_stafftime) / 3600.0 ELSE 0 END AS HORAS_NOCTURNAS, " +
                            "CASE WHEN([CMS13_TMP].dbo.hinc2(DATEPART([HOUR], row_date + [CMS13_TMP].dbo.HCMS(starttime))) >= 600 AND [CMS13_TMP].dbo.hinc2(DATEPART([HOUR], row_date + [CMS13_TMP].dbo.HCMS(starttime)) ) <= 2000 ) THEN " +
                            "SUM(ti_stafftime) / 3600.0 ELSE 0 END AS HORAS_DIURNAS " +
                            "FROM [CMS13_TMP].dbo.hagent WHERE (CAST(row_date AS DATE) = '" + registroEjecucion.fechaproceso.ToString("yyyy-MM-dd") + "') " +
                            "AND split <> 2000 AND logid IN (SELECT logid FROM [RH].[dbo].[Datos_ListadoTodosEmpleados_ConFecRetiro] datos WHERE datos.[GLBDescripcionCentroCostos] LIKE '%telebucaramanga%') " +
                            "GROUP BY row_date, logid, [CMS13_TMP].dbo.HINC2(DATEPART(HOUR, row_date + [CMS13_TMP].dbo.HCMS(starttime)))) Agent_STAFFINGTIME_HORA GROUP BY logid, row_date; ";
                }

                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();
                cmd.Dispose();
            }
            catch (Exception)
            {
                continuarProceso = false;
                throw;
            }
        }

        public void GetDataPersonalPanel(long dias, ref DataTable dataArchivosFinal)
        {
            try
            {
                conn.Crear();
                DataTable loginsPanel = conn.ObtenerTabla("SELECT login_aib FROM [Camp_Data].[dbo].[Movistar_RelacionLogins] WHERE [tipo_conexion] = 2 AND [estado] = 1;");

                if (loginsPanel.Rows.Count > 0)
                {
                    string sql = "SELECT * FROM OPENQUERY([POSTGRE], 'SELECT id, logid, cedula, COALESCE(fechainicial, (now() + interval ''-7'' day)) fechainicial, COALESCE(fechafinal, (now() + interval ''-7'' day)) fechafinal, extension, tiporegistro " +
                             "FROM publico.hlogueoaib WHERE ((fechainicial::date = (now() + interval ''" + dias + "'' day)::date) " +
                             "OR (fechafinal::date = (now() + interval ''" + dias + "'' day)::date)) AND tiporegistro = ''auto-in'' AND logid in (";

                    int i = 0;
                    foreach (DataRow dataRow in loginsPanel.Rows)
                    {
                        if (i > 0) sql += ",";
                        sql += "''" + dataRow["login_aib"].ToString() + "''";
                        i++;
                    }

                    sql += ") ORDER BY logid, id');";

                    DataTable data = conn.ObtenerTabla(sql);
                    conn.Destruir();

                    foreach (DataRow dtr in data.Rows)
                    {
                        DataRow dataNuevo = dataArchivosFinal.NewRow();

                        dataNuevo["Agente"] = dtr["logid"].ToString();
                        dataNuevo["Fecha Login"] = Convert.ToDateTime(dtr["fechainicial"].ToString()).ToString("dd/MM/yyyy");
                        dataNuevo["Hora Login"] = Convert.ToDateTime(dtr["fechainicial"].ToString()).ToString("HH:mm:ss");
                        dataNuevo["Fecha Logout"] = Convert.ToDateTime(dtr["fechafinal"].ToString()).ToString("dd/MM/yyyy");
                        dataNuevo["Hora Logout"] = Convert.ToDateTime(dtr["fechafinal"].ToString()).ToString("HH:mm:ss");
                        dataNuevo["Tiempo Login"] = 0;
                        dataNuevo["Login en Segundos"] = 0;
                        dataNuevo["Extension"] = 99999;
                        dataNuevo["Switch"] = "Agentes Fija";

                        dataArchivosFinal.Rows.Add(dataNuevo.ItemArray);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }

}