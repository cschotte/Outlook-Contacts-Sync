using NavaTron.Outlook.Contacts.Sync.Controllers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
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

namespace NavaTron.Outlook.Contacts.Sync.Views
{
    public partial class MainView : Window
    {
        private readonly BackgroundWorker worker = new BackgroundWorker();

        public MainView()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Title += Assembly.GetExecutingAssembly().GetName().Version.ToString();

            worker.DoWork += Worker_DoWork;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.WorkerReportsProgress = true;
        }

        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            StartButton.IsEnabled = false;
            SettingsButton.IsEnabled = false;

            worker.RunWorkerAsync();
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            StartButton.IsEnabled = true;
            SettingsButton.IsEnabled = true;
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            StatusTextBlock.Text = e.UserState.ToString();
            SyncProgressBar.Value = e.ProgressPercentage;
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                SyncController sync = new SyncController();

                worker.ReportProgress(0, Properties.Resources.GetDomainUsers);
                sync.GetDomainUsers();

                worker.ReportProgress(20, Properties.Resources.GetOutlookUsers);
                sync.GetOutlookUsers();

                worker.ReportProgress(40, Properties.Resources.UpdateOutlookUsers);
                sync.UpdateOutlookUsers();

                worker.ReportProgress(60, Properties.Resources.RemoveOutlookUsers);
                sync.RemoveOutlookUsers();

                worker.ReportProgress(80, Properties.Resources.AddOutLookUsers);
                sync.AddOutLookUsers();

                worker.ReportProgress(100, Properties.Resources.Ready);
            }
            catch (Exception ex)
            {
                worker.ReportProgress(0, ex.Message);
            }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            StartButton.IsEnabled = false;
            SettingsButton.IsEnabled = false;

            SettingsView window = new SettingsView();
            window.ShowDialog();

            StartButton.IsEnabled = true;
            SettingsButton.IsEnabled = true;
        }

    }
}