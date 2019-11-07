using System.Windows;

namespace Excel_Converter
{
    public partial class Popup : Window
    {
        public Popup()
        {
            InitializeComponent();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            var main = new MainWindow();

            main.ConvertData();

            this.Close();
        }
    }
}
