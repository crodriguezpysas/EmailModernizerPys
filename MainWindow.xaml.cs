using System;
using System.Threading.Tasks;
using System.Windows;

namespace EmailModernizerPys
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void BtnProcesar_Click(object sender, RoutedEventArgs e)
        {
            btnProcesar.IsEnabled = false;
            txtLog.Clear();
            AgregarLog("Iniciando procesamiento...");

            try
            {
                var processor = new Services.EmailProcessor(AgregarLog);
                await processor.ProcesarEmailsAsync();
                AgregarLog("Procesamiento completado.");
            }
            catch (Exception ex)
            {
                AgregarLog($"Error: {ex.Message}");
            }
            finally
            {
                btnProcesar.IsEnabled = true;
            }
        }

        private void AgregarLog(string mensaje)
        {
            Dispatcher.Invoke(() =>
            {
                txtLog.AppendText(mensaje + Environment.NewLine);
                txtLog.ScrollToEnd();
            });
        }
    }
}