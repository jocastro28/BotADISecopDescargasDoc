using System;
using System.Diagnostics;
using System.Threading;
using System.Windows;
using System.Windows.Threading;

namespace ETBRobotAsignarCasosPQR
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private DispatcherTimer _dispatcherTimer;
        private DateTime _fechaActual;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            _fechaActual = DateTime.Now;
            _dispatcherTimer = new DispatcherTimer();
            _dispatcherTimer.Tick += new EventHandler(DispatcherTimer_Tick);
            _dispatcherTimer.Interval = new TimeSpan(0, 0, 1);
            _dispatcherTimer.Start();
        }

        private void DispatcherTimer_Tick(object sender, EventArgs e)
        {
            TimeSpan tiempoTranscurrido = (DateTime.Now - _fechaActual);

            if (tiempoTranscurrido.Seconds == 5)
            {
                Application.Current.Shutdown();
            }
        }

        private void Window_Initialized(object sender, EventArgs e)
        {
            
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            
        }
    }
}
