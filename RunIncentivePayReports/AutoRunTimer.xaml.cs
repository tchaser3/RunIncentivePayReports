using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using System.Timers;

namespace RunIncentivePayReports
{
    /// <summary>
    /// Interaction logic for AutoRunTimer.xaml
    /// </summary>
    public partial class AutoRunTimer : Window
    {

        private static System.Timers.Timer aTimer;

        public AutoRunTimer()
        {
            InitializeComponent();
        }
        private void SetTimer()
        {
            // Create a timer with a two second interval.
            aTimer = new System.Timers.Timer(10000);
            // Hook up the Elapsed event for the timer. 
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            aTimer.AutoReset = true;
            aTimer.Enabled = true;
        }
        private void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            MainWindow.gblnAutoRun = true;
            aTimer.Stop();
            aTimer.Close();

            this.Dispatcher.Invoke(Close);

            //Close();
        }

        private void btnYes_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.gblnAutoRun = true;
            aTimer.Stop();
            aTimer.Close();
            this.Close();
        }

        private void btnNo_Click(object sender, RoutedEventArgs e)
        {
            aTimer.Stop();
            aTimer.Close();
            MainWindow.gblnAutoRun= false;
            
            this.Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            SetTimer();
        }
    }
}
