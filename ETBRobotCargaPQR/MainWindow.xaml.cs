using Microsoft.VisualBasic;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using ETBRobotAsignarCasosPQR.Clases;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using WindowScrape.Types;
using Newtonsoft.Json;
using ETBRobotAsignarCasosPQR.DAO;
using ETBRobotCargaPQR.Clases;
using System.Windows.Controls;

namespace ETBRobotAsignarCasosPQR
{
    public partial class MainWindow : Window
    {
        private string _msj = "";
        private IWebDriver driverChrome;
        private BackgroundWorker _hiloProceso;
        private DispatcherTimer _dispatcherTimer;
        private DateTime _fechaActual, _fechaUltimaAccion, _fechaBusqueda;
        private string llaveApp, empresa, numero_pqr;
        private RegistroEjecucion registroEjecucion;
        private SqlConnection conn = null;
        private SqlTransaction trans = null;
        private int NumberOfColumnsAcceptedPQRS = 67;//ACEPTED ROWS FOR PQRS
        private string ApplicationErrorFileName;
        private string ApplicationCaseBadDataFileName;
        private string ApplicationCaseNewPasswordFileName;
        private string SQLPointerString;

        private string ApplicationCaseChangeBadDataFileName;
        private string ApplicationDifferentCasesFileName;
        private DataTable PQRCompareData;

        
        private string[] NotificationFailEmails;
        private string[] NotificationSuccessEmails;
        private string[] NotificationPSWEmails;

        //DataADI
        string UrlWebOpen = ConfigurationManager.AppSettings["UrlSecop"];
        string UserNameO = ConfigurationManager.AppSettings["UserNameO"];
        string PasswordO = ConfigurationManager.AppSettings["PasswordO"];
        public MainWindow()
        {
            InitializeComponent();
        }
        private static Random random = new Random();
        private object driver;

        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }
        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
            { this.Opacity = 0.8; this.DragMove(); }

            this.Opacity = 1;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                llaveApp = ConfigurationManager.AppSettings["llaveApp"].ToString();
                empresa = ConfigurationManager.AppSettings["Empresa"].ToString();
                
                lbl_subtitulo.Content = empresa;

                _dispatcherTimer = new DispatcherTimer();
                _dispatcherTimer.Tick += new EventHandler(DispatcherTimer_Tick);
                _dispatcherTimer.Interval = new TimeSpan(0, 0, 1);

                NuevoHilo();
            }
            catch
            {
            }
        }

        private void DispatcherTimer_Tick(object sender, EventArgs e)
        {
            TimeSpan tiempoTranscurrido = (DateTime.Now - _fechaActual);

            Dispatcher.BeginInvoke(new Action(() =>
            {
                lbl_contador.Content = string.Format("{0:hh\\:mm\\:ss}", tiempoTranscurrido);
            }));
        }

        private void HiloProceso_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                //SE ESTABLECE EL AMBIENTE DEL ROBOT
                switch (ConfigurationManager.AppSettings["enviroment"].ToString())
                {
                    case "0"://Apunta a desarrollo
                        SQLPointerString = "ATLASConnectionDevelopment";
                        break;
                    case "1"://Apunta a qa
                        SQLPointerString = "ATLASConnectionQA";
                        break;
                    case "2"://Apunta a produccion
                        SQLPointerString = "ATLASConnection";
                        break;
                    default:
                        SQLPointerString = "ATLASConnectionDevelopment";
                        break;
                }
                NotificationFailEmails = new string[] { ConfigurationManager.AppSettings["notificationEmailsFail"].ToString() };
                NotificationSuccessEmails = new string[] { ConfigurationManager.AppSettings["notificationEmailsSuccess"].ToString() };
//                NotificationPSWEmails = new string[] { ConfigurationManager.AppSettings["notificationEmailsPSW"].ToString() };
                var worker = (BackgroundWorker)sender;
                if (worker.CancellationPending) return;
                // Inicio
                _msj = "(Robot) Inicio Proceso";
                _hiloProceso.ReportProgress(1);

                Thread.Sleep(500);

                var processes = Process.GetProcesses().Where(pr => pr.ProcessName.Contains("chromedriver"));
                foreach (var process in processes)
                {
                    process.Kill();
                }

                // Validar Configuraciones
                _msj = "(Robot) Validar configuraciones";
                _hiloProceso.ReportProgress(2);
                 Config config = new ProcesoDAO().GetConfig(empresa);

               // ETB_RobotAsignacionMasiva_Insert_reasignaciones Reasignar = new ProcesoDAO().Get_Insert_reasignaciones();
                //
                DateTime CurrentDate = DateTime.Now;
               

                // Validar fecha de busqueda
                string[] clArgs = Environment.GetCommandLineArgs();
                if (clArgs.Count() >= 3)
                {
                    for (int i = 0; i < 3; i++)
                    {
                        if (clArgs[i] == "-f")
                        {
                           // _fechaBusqueda = Convert.ToDateTime(clArgs[i + 1]);
                        }
                    }
                }
                else
                {
                   // _fechaBusqueda = _fechaActual;
                }
                           
                _hiloProceso.ReportProgress(3);
                Thread.Sleep(500);
               
                string rutaDescarga = Environment.ExpandEnvironmentVariables("%USERPROFILE%\\Downloads");
               
                bool continuarProceso = true;
               
                //// Obtener información del numero
                //registroEjecucion = new ProcesoDAO().GetRegistroEjecucion(empresa, Environment.UserName, ConfigurationManager.AppSettings["enviroment"].ToString());

                //if (registroEjecucion.id == 0)
                //{
                //    //THERE ARE DIFERENCES
                //    string ContenidoMsj = "<strong>ERROR EN CARGA DEL ROBOT</strong></br></br>";
                //    ContenidoMsj += "(Robot) Nada por procesar para el día de hoy";
                //    SendMessages SM = new SendMessages();
                //    SM.SendAMessage("Falla en carga del robot", ContenidoMsj, NotificationFailEmails, SQLPointerString);

                //    _msj = "(Robot) Nada por procesar para ese día";
                //    _hiloProceso.ReportProgress(50);

                //    Thread.Sleep(500);
                //    goto FinalProceso;
                //}

                //_msj = "(Data) Intentos: " + registroEjecucion.intentos.ToString();
                _hiloProceso.ReportProgress(5);
                Thread.Sleep(500);

                //// Fechas para descargar archivos
                //List<DateTime> fechasArchivos = new List<DateTime>
                //{
                //    registroEjecucion.fechaproceso
                //};

                //List<string> archivosDescargados = new List<string>();

                int progreso = 5;
                continuarProceso = true;
                string nombreArchivoDescargado = "";

                // Descargar archivos por listado de fecha              

                    //continuarProceso = true;
                    //nombreArchivoDescargado = rutaDescarga + "\\" + config.filename.ToString() + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
                 continuarProceso = true;
                    nombreArchivoDescargado = rutaDescarga + "\\ContratosADI-" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
                



                _msj = "(Robot) Procesar y cargar archivo";
                _hiloProceso.ReportProgress(56);
                Thread.Sleep(500);

                _msj = "";
                //convierte excel a datatable
                string nombreHoja = "";
               
                DataTable dataArchivosFinal = new DataTable();

                if (nombreArchivoDescargado != null)
                {
                     string connStrArchivo = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + nombreArchivoDescargado + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1'";
                        OleDbConnection connArchivo = new OleDbConnection(connStrArchivo);

                        connArchivo.Open();
                        DataTable hojas = connArchivo.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        connArchivo.Close();

                        foreach (DataRow rows in hojas.Rows)
                        {
                            nombreHoja = rows["TABLE_NAME"].ToString();
                        }

                        // Pasar Excel a DataTable
                        DataTable dataHoja = new DataTable();
                        string strDataHoja = "SELECT * FROM [" + nombreHoja + "];";

                        OleDbDataAdapter dataAdapterHoja = new OleDbDataAdapter(strDataHoja, connArchivo);
                        dataAdapterHoja.Fill(dataHoja);
                        dataHoja.TableName = "tbl_data";

                        dataArchivosFinal = dataHoja;
                                                         
                    
                }
                else
                {
                    //THERE ARE DIFERENCES
                    string ContenidoMsj = "<strong>ERROR EN CARGA DEL ROBOT</strong></br></br>";
                    ContenidoMsj += "(Robot) No se cuenta con la data necesaria (1)";
                    SendMessages SM = new SendMessages();
                    SM.SendAMessage("Falla en carga del robot", ContenidoMsj, NotificationFailEmails, SQLPointerString);

                    _msj = "(Robot) No se cuenta con todos los archivos necesarios";
                    _hiloProceso.ReportProgress(90);

                    registroEjecucion.procesado = false;
                    registroEjecucion.observaciones = _msj;

                    Thread.Sleep(500);
                    goto GuardarProceso;
                }

               
                _msj = "(Robot) Cargando información a base de datos";
                _hiloProceso.ReportProgress(80);
                Thread.Sleep(500);
                //descarga de contratos              

                IngresoSecop(config, dataArchivosFinal, progreso, out continuarProceso);

              
                    

                    registroEjecucion.procesado = true;
                    registroEjecucion.observaciones = "OK";

                    _msj = "(Robot) Guardado exitosamente!";
                    _hiloProceso.ReportProgress(90);
                    Thread.Sleep(500);                  
                

            // Final
            GuardarProceso:
                //if (!registroEjecucion.procesado && registroEjecucion.intentos == 50)
                //{
                //    _msj = "(Robot) Mandando a la cola humana";
                //    _hiloProceso.ReportProgress(93);
                //    Thread.Sleep(500);

                //    new ProcesoDAO().ColaHumanaBot(registroEjecucion);
                //    e.Result = 0;
                //}
                //else if (!registroEjecucion.procesado && registroEjecucion.intentos != 5)
                //{
                //    e.Result = 0;
                //}
                //else
                //{
                //    e.Result = 1;
                //}

                //new ProcesoDAO().GuardarRegistroEjecucion(registroEjecucion);

                //_msj = "(Robot) Proceso guardado";
                //_hiloProceso.ReportProgress(95);
                //Thread.Sleep(500);

                //LogWeb.Send("Guardar", llaveApp, registroEjecucion.id.ToString(), _fechaActual);

            FinalProceso:
                _msj = "(Robot) Final" + Environment.NewLine;
                _hiloProceso.ReportProgress(100);
                Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                //File.AppendAllText(ApplicationErrorFileName, "Error al insertar la consulta," + ex.Message.Replace(",", " ") + "," + _fechaBusqueda.ToString("yyyyMMddHHmm") + Environment.NewLine);
                _msj = ex.Message;
                _hiloProceso.ReportProgress(95);
                Thread.Sleep(1000);

                try
                {
                    //registroEjecucion.observaciones = _msj;
                    //new ProcesoDAO().GuardarRegistroEjecucion(registroEjecucion);
                }
                catch (Exception)
                {
                }
                //THERE ARE DIFERENCES
                string ContenidoMsj = "<strong>ERROR EN CARGA DEL ROBOT</strong></br></br>";
                ContenidoMsj += ex.Message;
                SendMessages SM = new SendMessages();
                SM.SendAMessage("Falla en carga del robot", ContenidoMsj, NotificationFailEmails, SQLPointerString);
                e.Cancel = true;
            }
            finally
            {
                if (driverChrome != null)
                {
                    driverChrome.Quit();
                    driverChrome.Dispose();
                }
            }
        }

        private void HiloProceso_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            EscribirProceso();
            prg_estado.Value = e.ProgressPercentage;
        }

        private void HiloProceso_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                _msj = "(Robot) Operacion Cancelada! Reiniciando ..." + Environment.NewLine;
                Thread.Sleep(10000);

                EscribirProceso();
                CerrarAplicacion();
            }
            else if (e.Error != null)
            {
                _msj = "(Robot) Error: " + e.Error + ". Reiniciando ..." + Environment.NewLine;
                Thread.Sleep(10000);

                EscribirProceso();
                CerrarAplicacion();
            }
            else if (Convert.ToInt32(e.Result) == 0)
            {
                _msj = "(Robot) Reiniciando ..." + Environment.NewLine;
                EscribirProceso();
                CerrarAplicacion();
            }
            else
            {
                CerrarAplicacion(1);
            }
        }

        private void Btn_min_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.WindowState = WindowState.Minimized;
        }

        private void Btn_cerrar_Click(object sender, RoutedEventArgs e)
        {
            CerrarAplicacion(1);
        }

        private void NuevoHilo()
        {
        Inicio:
            _hiloProceso = new BackgroundWorker
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };

            _hiloProceso.DoWork += HiloProceso_DoWork;
            _hiloProceso.ProgressChanged += HiloProceso_ProgressChanged;
            _hiloProceso.RunWorkerCompleted += HiloProceso_RunWorkerCompleted;

            _fechaActual = DateTime.Now;

            if (!_hiloProceso.IsBusy)
            {
                _hiloProceso.RunWorkerAsync();
                _dispatcherTimer.Start();
            }
            else
            {
                _hiloProceso.CancelAsync();
                goto Inicio;
            }
        }

        private void IngresoSecop(Config config, DataTable dtasignarCasos, int progreso, out bool continuarProceso)
        {
            continuarProceso = true;
            
            try
            {
                string url = ConfigurationManager.AppSettings["UrlSecop"].ToString();                
                
                _msj = "";
                _hiloProceso.ReportProgress(progreso + 1);
                
                //selenium

                ChromeOptions options = new ChromeOptions();
                //options.AddArguments("--no-sandbox --incognito --disable-gpu");                
                options.AddUserProfilePreference("disable-popup-blocking", true);

                ChromeDriverService service = ChromeDriverService.CreateDefaultService();
                // service.SuppressInitialDiagnosticInformation = true;

                driverChrome = new ChromeDriver(service, options, TimeSpan.FromMinutes(10));
                _hiloProceso.ReportProgress(progreso + 2);

                List<HwndObject> cmds = HwndObject.GetWindows();
                HwndObject cmd = cmds.FirstOrDefault<HwndObject>(x => x.Title.Contains("chromedriver"));
                cmd.Location = new System.Drawing.Point(355, 0);

                _hiloProceso.ReportProgress(progreso + 3);

                driverChrome.Navigate().GoToUrl(url);
                driverChrome.Manage().Window.Maximize();
                driverChrome.SwitchTo().ActiveElement();
                IWebElement webElementOpen = null;
                string pathO = string.Empty;
                // iniciar conexion a bd
                SqlServer conn = new SqlServer();
                conn.CrearAtlas(SQLPointerString);

                try {
                    if (dtasignarCasos.Rows.Count > 0) {

                        foreach (DataRow elemento in dtasignarCasos.Rows) {
                            conn.EliminarParametros();
                            conn.AgregarParametro("@NumeroCuenta", elemento["NUMEROCUENTA"].ToString());
                            conn.AgregarParametro("@Estado", "NOTRABAJADO");
                            conn.AgregarParametro("@Observacion", "NOTRABAJADO");
                            conn.EjecutarSP("SP_SAVE_CUENTAS");
                            
                        }
                                               
                    }


                } catch (Exception ex) {
                }

                

                _msj = "(BO) Iniciando sesión";
                _hiloProceso.ReportProgress(progreso + 4);
                //inicio login
                //Pasamos Usuario                    
                pathO = "/html/body/div[2]/div[2]/div[1]/form/div/div[2]/div/input";
                while (!ValidarElemento(driverChrome, ref webElementOpen, pathO)) { /* Wating...*/ }
                while (!webElementOpen.Enabled || !webElementOpen.Displayed) { /* Wating...*/ }
                webElementOpen.Click();
                webElementOpen.SendKeys(UserNameO);
                _msj = "-> Pasamos Usuario SECOP...Ok \n";
                Thread.Sleep(600);

                //Pasamos Password                  
                pathO = "/html/body/div[2]/div[2]/div[1]/form/div/div[4]/div/input";
                while (!ValidarElemento(driverChrome, ref webElementOpen, pathO)) { /* Wating...*/ }
                while (!webElementOpen.Enabled || !webElementOpen.Displayed) { /* Wating...*/ }
                webElementOpen.Click();
                webElementOpen.SendKeys(PasswordO);
                _msj= "-> Pasamos Password SECOP...Ok \n";
                Thread.Sleep(600);

                /////Click en ingresar                
                pathO = "/html/body/div[2]/div[2]/div[1]/form/div/div[11]/input";
                while (!ValidarElemento(driverChrome, ref webElementOpen, pathO)) { /* Wating...*/ }
                while (!webElementOpen.Enabled || !webElementOpen.Displayed) { /* Wating...*/ }
                webElementOpen.Click();
                _msj = "-> Click Acceder...Ok \n";
                Thread.Sleep(2000);

                /////Click en Entrar                
                pathO = "/html/body/div[2]/div[2]/div[1]/div[1]/div[7]/input";
                while (!ValidarElemento(driverChrome, ref webElementOpen, pathO)) { /* Wating...*/ }
                while (!webElementOpen.Enabled || !webElementOpen.Displayed) { /* Wating...*/ }
                webElementOpen.Click();
                _msj = "-> Click Entrar...Ok \n";
                Thread.Sleep(3000);
                /////Click en menu                
                pathO = "/html/body/div[2]/div[1]/div/div/div[5]/div[1]/div[2]/div/div";
                while (!ValidarElemento(driverChrome, ref webElementOpen, pathO)) { /* Wating...*/ }
                while (!webElementOpen.Enabled || !webElementOpen.Displayed) { /* Wating...*/ }
                webElementOpen.Click();
                  _msj= "-> Click Menu...Ok \n";
                Thread.Sleep(1000);
                /////Click en Submenu mis contratos                
                pathO = "/html/body/div[2]/div[1]/div/div/div[5]/div[1]/div[2]/div/div/div[1]/div/table/tbody/tr[2]/td/a";
                while (!ValidarElemento(driverChrome, ref webElementOpen, pathO)) { /* Wating...*/ }
                while (!webElementOpen.Enabled || !webElementOpen.Displayed) { /* Wating...*/ }
                webElementOpen.Click();
                  _msj= "-> Click Sub Menú Mis contratos...Ok \n";
                Thread.Sleep(3000);

                //validar que el caso no haya sido trabajado
                try {
                    SqlServer conn2 = new SqlServer();
                    conn2.Crear();
                    DataTable dtElementos = conn2.ObtenerTabla(string.Format("SELECT * FROM [ADI].[dbo].[NumerosDeContratosADI_PDF] where Estado='NOTRABAJADO'"));
                    conn2.Destruir();
                    if (dtElementos.Rows.Count > 0) {
                        // Reasignar pqr
                        int cont = 0;
                        string NumContrato = "";
                        string observacion = "";
                        string estado = "";
                        //string usuario_Actual = "";
                        //string usuario_Asignar = "";
                        //string Altura = "";

                        foreach (DataRow elemento in dtasignarCasos.Rows)
                        {
                            if (cont > 0)
                            {

                                conn.EliminarParametros();
                                //conn.AgregarParametro("@usuario_asignado",  usuario_Actual);
                                //conn.AgregarParametro("@usuario_final", usuario_Asignar);
                                conn.AgregarParametro("@NumContrato", NumContrato);
                                //conn.AgregarParametro("@Altura", Altura);
                                conn.AgregarParametro("@estado", estado);
                                conn.AgregarParametro("@observacion", observacion);
                                conn.EjecutarSP("Sp_Insertar_ContratosADI_pdf");
                            }
                            NumContrato = "";
                            //usuario_Actual = "";
                            //usuario_Asignar = "";
                            NumContrato = elemento["NumeroContrato"].ToString();
                            //usuario_Actual = elemento["usuario_asignado"].ToString();
                            //usuario_Asignar = elemento["usuario_final"].ToString();
                            //Altura = elemento["Altura"].ToString();
                            observacion = "";
                            estado = "";
                            cont++;

                            try
                            {
                                //Buscar por Número de contrato Usuario     

                                pathO = "/html/body/div[2]/div[2]/table/tbody/tr/td[3]/form/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/fieldset/div/div/table[1]/tbody/tr/td[1]/input";
                                while (!ValidarElemento(driverChrome, ref webElementOpen, pathO)) { /* Wating...*/ }
                                while (!webElementOpen.Enabled || !webElementOpen.Displayed) { /* Wating...*/ }
                                webElementOpen.Click();
                                webElementOpen.SendKeys(NumContrato);
                                _msj = "-> Buscar por Número de contrato Usuario...Ok \n";
                                Thread.Sleep(600);

                                //CLIC boton buscar contrato
                                pathO = "/html/body/div[2]/div[2]/table/tbody/tr/td[3]/form/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/fieldset/div/div/table[1]/tbody/tr/td[2]/input";
                                while (!ValidarElemento(driverChrome, ref webElementOpen, pathO)) { /* Wating...*/ }
                                while (!webElementOpen.Enabled || !webElementOpen.Displayed) { /* Wating...*/ }
                                webElementOpen.Click();
                                _msj = "-> Click Boton buscar contrato...Ok \n";
                                Thread.Sleep(3000);

                                //CLIC boton detalle de la tabla
                                pathO = "/html/body/div[2]/div[2]/table/tbody/tr/td[3]/form/table/tbody/tr[2]/td/table/tbody/tr[5]/td/div[1]/table/tbody/tr[2]/td[11]/a[1]";
                                while (!ValidarElemento(driverChrome, ref webElementOpen, pathO)) { /* Wating...*/ }
                                while (!webElementOpen.Enabled || !webElementOpen.Displayed) { /* Wating...*/ }
                                webElementOpen.Click();
                                _msj = "-> Click boton detalle de la tabla...Ok \n";
                                Thread.Sleep(3000);

                                //CLIC boton Documentos del provedor
                                pathO = "/html/body/div[2]/div[2]/table/tbody/tr/td[1]/div/div/div[4]/span";
                                while (!ValidarElemento(driverChrome, ref webElementOpen, pathO)) { /* Wating...*/ }
                                while (!webElementOpen.Enabled || !webElementOpen.Displayed) { /* Wating...*/ }
                                webElementOpen.Click();
                                _msj = "-> Click boton Documentos del provedor...Ok \n";
                                Thread.Sleep(3000);

                                //CLIC boton Documentos Descargar PDF
                                pathO = "/html/body/div[2]/div[2]/table/tbody/tr/td[3]/form/table/tbody/tr[3]/td/table/tbody/tr[2]/td/div/div[5]/div[1]/table/tbody/tr/td/table/tbody/tr[4]/td[8]/a[1]";
                                while (!ValidarElemento(driverChrome, ref webElementOpen, pathO)) { /* Wating...*/ }
                                while (!webElementOpen.Enabled || !webElementOpen.Displayed) { /* Wating...*/ }
                                webElementOpen.Click();
                                _msj = "-> Click boton Documentos Descargar PDF...Ok \n";
                                Thread.Sleep(3000);

                                //webElementOpen.SendKeys(Keys.Enter);
                                //Thread.Sleep(3000);
                                //CLIC Admin contratos
                                pathO = "/html/body/div[2]/table/tbody/tr[1]/td[1]/div/a[3]";
                                while (!ValidarElemento(driverChrome, ref webElementOpen, pathO)) { /* Wating...*/ }
                                while (!webElementOpen.Enabled || !webElementOpen.Displayed) { /* Wating...*/ }
                                webElementOpen.Click();
                                _msj = "-> Click boton Documentos Descargar PDF...Ok \n";
                                Thread.Sleep(3000);

                                ////CLIC boton Ejecucion del contrato
                                //pathO = "/html/body/div[2]/div[2]/table/tbody/tr/td[1]/div/div/div[7]/span";
                                //while (!ValidarElemento(driverChrome, ref webElementOpen, pathO)) { /* Wating...*/ }
                                //while (!webElementOpen.Enabled || !webElementOpen.Displayed) { /* Wating...*/ }
                                //webElementOpen.Click();
                                //_msj = "-> Click boton ejecucion del contrato ...Ok \n";
                                //Thread.Sleep(2000);

                                ////CLIC boton Cargar nuevo
                                //pathO = "/html/body/div[2]/div[2]/table/tbody/tr/td[3]/form/table/tbody/tr[3]/td/table/tbody/tr[2]/td/div/div[8]/fieldset[4]/div/table/tbody/tr[3]/td/div/input[2]";
                                //while (!ValidarElemento(driverChrome, ref webElementOpen, pathO)) { /* Wating...*/ }
                                //while (!webElementOpen.Enabled || !webElementOpen.Displayed) { /* Wating...*/ }
                                //webElementOpen.Click();
                                //_msj = "-> Click boton cargar nuevo ...Ok \n";
                                //Thread.Sleep(3000);

                                ////documentos anexos
                                //pathO = "/html/body/div[2]/div/div[2]/form/div[1]/div[2]/div/div/input";
                                //ValidarElemento(driverChrome, ref webElementOpen, pathO);
                                //if (webElementOpen != null)
                                //{
                                //    FileInfo[] adjuntosPDF = ArchivosAdjuntosPQR();
                                //    List<string> rutasAdjuntosPqr = new List<string>();
                                //    foreach (var f2 in adjuntosPDF)
                                //    {
                                //        if (!f2.FullName.Contains("$"))
                                //        {
                                //            rutasAdjuntosPqr.Add(f2.FullName);
                                //        }

                                //    }

                                //    foreach (var rutaadjunto in rutasAdjuntosPqr)
                                //    {
                                //        try
                                //        {
                                //            ValidarElemento(driverChrome, ref webElementOpen, pathO);
                                //            webElementOpen.SendKeys(rutaadjunto);
                                //            Thread.Sleep(600);
                                //            System.IO.File.Delete(@"" + rutaadjunto);
                                //        }
                                //        catch (Exception e)
                                //        {

                                //            throw;
                                //        }

                                //    }
                                //}
                                //Thread.Sleep(10000);


                                observacion = "Se descargo contrato " + NumContrato;
                                estado = "Exitoso";



                            }
                            catch (Exception ex)
                            {
                                observacion = ex.Message;
                                estado = "No exitoso";
                            }


                        }

                        conn.EliminarParametros();
                        //conn.AgregarParametro("@usuario_asignado", usuario_Asignar);
                        //conn.AgregarParametro("@usuario_final", usuario_Actual);
                        conn.AgregarParametro("@NumContrato", NumContrato);
                        //conn.AgregarParametro("@Altura", Altura);
                        conn.AgregarParametro("@estado", estado);
                        conn.AgregarParametro("@observacion", observacion);
                        conn.EjecutarSP("Sp_Insertar_ContratosADI_pdf");
                        conn.Destruir();
                        // driverChrome.Navigate().GoToUrl(config.url);
                        //driverChrome.SwitchTo().ActiveElement();

                        // Thread.Sleep(2000);
                        // _msj = "(BO) Descargando reporte";
                        // _hiloProceso.ReportProgress(progreso + 15);

                        // int contArchivo = 0;
                        // string FullNameFile = "LISTADO_PQRS.xlsx";


                        // if (contArchivo > 300)
                        // {
                        //     //_msj = "(BO) Tiempo superado para descargar el reporte del dia " + fecha.ToString("yyyy-MM-dd");
                        //     _hiloProceso.ReportProgress(progreso + 25);

                        //     Thread.Sleep(500);
                        //     continuarProceso = false;
                        //     return;
                        // }

                        //// _msj = "(BO) Reporte del dia " + fecha.ToString("yyyy-MM-dd") + " descargado";
                        // _hiloProceso.ReportProgress(progreso + 25);
                        // Thread.Sleep(500);

                        // nombreArchivoDescargado = FullNameFile;

                        //THERE ARE DIFERENCES
                        string ContenidoMsj = "<strong>EXITOSO</strong></br></br>";
                        SendMessages SM = new SendMessages();
                        SM.SendAMessage("Exito en descargar contratos", ContenidoMsj, NotificationSuccessEmails, SQLPointerString);
                    }

                } catch (Exception ex){
                    _msj = ex.Message;
                }


             
            }
            catch (Exception ex)
            {
                _msj = ex.Message;
                continuarProceso = false;
            }
            finally
            {
                if (driverChrome != null)
                {
                    driverChrome.Quit();
                    driverChrome.Dispose();
                }
            }
        }

        
        private void EscribirProceso()
        {
            _fechaUltimaAccion = DateTime.Now;
            if (!string.IsNullOrWhiteSpace(_msj)) txt_log.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + ": " + _msj + Environment.NewLine);
            if (chk_autoscroll.IsChecked == true) txt_log.ScrollToEnd();
        }

        private void CerrarAplicacion(int reiniciar = 0)
        {
            try
            {
                new CerrandoWindow().Show();

                var processes = Process.GetProcesses().Where(pr => pr.ProcessName.Contains("chromedriver") || pr.ProcessName.Contains("ETB"));
                foreach (var process in processes)
                {
                    process.Kill();
                }

                this.Close();
                if (reiniciar == 0) new ProcesoDAO().ReiniciarBot(empresa);
                Thread.Sleep(2000);
            }
            catch
            {
            }

            Application.Current.Shutdown();
        }

        public static FileInfo[] ArchivosAdjuntosPQR()
        {
            string RutaCartaLocal = ConfigurationManager.AppSettings["RutaCartaLocal"];
            string rutaadjuntosPqr = RutaCartaLocal;
            FileInfo[] files = new FileInfo[] { };
            try
            {
                DirectoryInfo di = new DirectoryInfo(rutaadjuntosPqr);
                files = di.GetFiles("*", SearchOption.AllDirectories);
            }
            catch (Exception e)
            {
                //esta pqr no tiene adjuntos
                files = new FileInfo[] { };
            }
            return files;
        }
        public static bool ValidarElemento(IWebDriver pDriver, ref IWebElement pElemento, string pLlave)
        {
            bool resp = default(bool);
            pElemento = null;


            // Busco por Id
            try
            {
                pElemento = pDriver.FindElement(By.Id(pLlave));

                return true;
            }
            catch (Exception)
            {
                resp = false;
            }

            // Busco por Name
            try
            {
                pElemento = pDriver.FindElement(By.Name(pLlave));
                return true;
            }
            catch (Exception)
            {
                resp = false;
            }

            // Busco por XPath
            try
            {
                pElemento = pDriver.FindElement(By.XPath(pLlave));
                return true;
            }
            catch (Exception)
            {
                resp = false;
            }

            return resp;
        }
        private IWebElement BuscarElemento(IWebDriver driver, string strElemento)
        {
            try
            {
                SqlServer conn = new SqlServer();
                conn.Crear();
                DataTable dtElementos = conn.ObtenerTabla(string.Format("SELECT * FROM [Camp_Data].[dbo].[ETB_RobotAsignacionMasiva_Elementos] WHERE [Elemento] = '{0}' AND [Estado] = '{1}';", strElemento, '1'));
                conn.Destruir();

                IWebElement result = null;
                foreach (DataRow elemento in dtElementos.Rows)
                {
                    try
                    {
                        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(int.Parse(elemento["TiempoBusquedaSeg"].ToString())));

                        if (elemento["Tipo"].ToString() == "Id")
                        {
                            result = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id(elemento["Valor"].ToString())));
                        }
                        else if (elemento["Tipo"].ToString() == "Name")
                        {
                            result = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Name(elemento["Valor"].ToString())));
                        }
                        else if (elemento["Tipo"].ToString() == "XPath")
                        {
                            result = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(elemento["Valor"].ToString())));
                        }
                        else if (elemento["Tipo"].ToString() == "CssSelector")
                        {
                            result = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.CssSelector(elemento["Valor"].ToString())));
                        }
                        else if (elemento["Tipo"].ToString() == "ClassName")
                        {
                            result = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.ClassName(elemento["Valor"].ToString())));
                        }
                        else if (elemento["Tipo"].ToString() == "LinkText")
                        {
                            result = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.LinkText(elemento["Valor"].ToString())));
                        }

                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error buscando elemento " + strElemento + ": " + ex.Message);
                    }
                }

                return result;
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}