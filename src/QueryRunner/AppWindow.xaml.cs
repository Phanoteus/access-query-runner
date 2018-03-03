using System.Windows;
using System.Windows.Input;

namespace QueryRunner
{
    /// <summary>
    /// Interaction logic for AppWindow.xaml
    /// </summary>
    public partial class AppWindow : Window
    {
        private AppViewModel _viewModel = null;

        public AppWindow()
        {
            InitializeComponent();
            this.Loaded += AppWindow_Loaded;
        }

        private void AppWindow_Loaded(object sender, RoutedEventArgs e)
        {
            SetDataContext();
        }

        private void SetDataContext()
        {
            _viewModel = new AppViewModel();
            DataContext = _viewModel;
        }

        private void Close_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void Close_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            this.Close();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if ((_viewModel != null) && (_viewModel.Idle == false))
            {

                MessageBox.Show("Queries are being processed.\r\n\r\n" +
                    "The application window can be closed after processing is complete.\r\n\r\n" +
                    "You can minimize the window to continue working while queries are being processed.",
                    "Queries in Progress",
                    MessageBoxButton.OK, MessageBoxImage.Warning, MessageBoxResult.OK);
                e.Cancel = true;
            }
            else
            {
                e.Cancel = false;
            }
        }
    }
}
