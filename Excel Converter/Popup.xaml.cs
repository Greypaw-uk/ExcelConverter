using System.Windows;

namespace Excel_Converter
{
    public partial class Popup : Window
    {
        MainWindowDataContext context = new MainWindowDataContext();

        public Popup()
        {
            InitializeComponent();
            DataContext = context;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            Hide();
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            var main = new MainWindow();

            main.YearGroup = YGPicker.Text;

            main.ConvertData();

            Hide();
        }
    }
}
