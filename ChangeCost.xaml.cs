using SecurityPlus.Models;
using System.Windows;

namespace SecurityPlus
{
    /// <summary>
    /// Interaction logic for ChangeCost.xaml
    /// </summary>
    public partial class ChangeCost : Window
    {
        public ChangeCost()
        {
            InitializeComponent();
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            
            FileManager.SetSettings(new Settings
            {
                CostInCar = decimal.Parse(TbInCar.Text),
                CostOutCar = decimal.Parse(TbNoCar.Text)
            });

            MessageBox.Show("Для применения параметров перезапустите приложение");
        }
    }
}
