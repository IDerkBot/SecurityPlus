using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MaterialDesignThemes.Wpf;
using SecurityPlus.Models;
using Word = Microsoft.Office.Interop.Word;

namespace SecurityPlus
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            _costInCar = FileManager.GetSettings().CostInCar;
            _costOutCar = FileManager.GetSettings().CostOutCar;
        }

        private decimal _costInCar;
        private decimal _costOutCar;

        private void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
            DgData.ItemsSource = new List<Duty>();
        }

        private void DgData_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            // Корректный метод
            GetTimeAndSum();
        }

        private void GetTimeAndSum()
        {
            // Получаем ячейки
            var items = DgData.ItemsSource.Cast<Duty>().ToList();

            int minutes = 0;

            foreach (var item in items)
            {
                item.VisibleId = item.Id;
                item.VisibleFullname = item.Fullname;
                item.VisibleDateStart = item.DateStart;
                item.VisibleDateEnd = item.DateEnd;
                item.VisibleTimeStart = item.TimeStart;
                item.VisibleTimeEnd = item.TimeEnd;

                item.DateStart = new DateTime(item.DateStart.Year, item.DateStart.Month, item.DateStart.Day, item.TimeStart.Hour, item.TimeStart.Minute, 0);
                item.DateEnd = new DateTime(item.DateEnd.Year, item.DateEnd.Month, item.DateEnd.Day, item.TimeEnd.Hour, item.TimeEnd.Minute, 0);

                var time = item.DateEnd - item.DateStart;
                minutes = (time.Days * 24 + time.Hours) * 60 + time.Minutes;
                item.Time = minutes;
                item.TimeString = $"{time.Days * 24 + time.Hours}:{time.Minutes}";
                
                item.Sum = item.IsCar ? minutes * (_costInCar * 100) / 6000 : minutes * (_costOutCar * 100) / 6000;
            }

            DgData.ItemsSource = items;

            TbTotalTime.Text = $"{items.Sum(x => x.Time) / 60}ч:{items.Sum(x => x.Time) % 60}м";
            TbTotalSum.Text = items.Sum(x => x.Sum).ToString(CultureInfo.InvariantCulture);
        }

        private void DatePickerDateEnd_OnSelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            var duty = (sender as DatePicker).DataContext as Duty;
            if (duty != null)
            {
                duty.DateEnd = (DateTime)(sender as DatePicker).SelectedDate;
            }
        }

        private void DatePickerDateStart_OnSelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            var duty = (sender as DatePicker).DataContext as Duty;
            if (duty != null)
            {
                duty.DateStart = (DateTime)(sender as DatePicker).SelectedDate;
            }
        }

        private void TimePickerStart_OnSelectedTimeChanged(object sender, RoutedPropertyChangedEventArgs<DateTime?> e)
        {
            var duty = (sender as TimePicker).DataContext as Duty;
            if (duty != null)
            {
                duty.TimeStart = (DateTime)((TimePicker)sender).SelectedTime;
            }
        }

        private void TimePickerEnd_OnSelectedTimeChanged(object sender, RoutedPropertyChangedEventArgs<DateTime?> e)
        {
            var duty = (sender as TimePicker).DataContext as Duty;
            if (duty != null)
            {
                duty.TimeEnd = (DateTime)((TimePicker)sender).SelectedTime;
            }
        }

        private void TextBoxFullname_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var duty = (sender as TextBox).DataContext as Duty;
            if (duty != null)
            {
                duty.Fullname = ((TextBox)sender).Text;
            }
        }

        private void TextBoxId_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var duty = (sender as TextBox).DataContext as Duty;
            if (duty != null)
            {
                duty.Id = int.Parse(((TextBox)sender).Text);
            }
        }

        private void ToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            var duty = (sender as CheckBox).DataContext as Duty;
            if (duty != null)
            {
                duty.IsCar = (bool)((CheckBox)sender).IsChecked;
            }
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                new Thread(() =>
                {
                    try
                    {
                        // Создали объект ворда
                        var app = new Word.Application();

                        //Открываем файл
                        //app.Documents.Open($@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\Template.docx");
                        app.Documents.Open($@"{AppDomain.CurrentDomain.BaseDirectory}\Views\Resources\Template.docx");

                        //Заменяем слова
                        FindAndReplace(app, GetDictionary());

                        if (!Directory.Exists(
                                $@"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\Дежурства\"))
                            Directory.CreateDirectory(
                                $@"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\Дежурства\");
                        
                        //Сохраняем файл
                        //app.Visible = true;
                        var dateNow = DateTime.Now.ToString("D");
                        app.ActiveDocument.SaveAs(
                            $@"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\Дежурства\{dateNow}.docx");
                        app.ActiveDocument.Close();
                        app.Quit();
                        //Process.Start(
                        //    $@"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\Дежурства\{DateTime.Now}.docx");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка {ex.Message}\n\n\nВнутреннее исключение{ex.InnerException}");
                    }

                }).Start();
            }
            catch (Exception exception)
            {
                MessageBox.Show($"Ошибка {exception.Message}\n\n\nВнутреннее исключение{exception.InnerException}");
            }
        }

		private Dictionary<string, string> GetDictionary()
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            Dispatcher.Invoke(() =>
            {
                var culture = new CultureInfo("ru-RU");

                var items = DgData.ItemsSource.Cast<Duty>().ToList();
                var count = 1;
                dict = new Dictionary<string, string>();

                dict.Add("<TITLE>", TbName.Text);
                string info = " ";
                foreach (var item in items)
                {
                    dict.Add($"<ID{count}>", item.Id.ToString());
                    dict.Add($"<FULLNAME{count}>", item.Fullname);
                    dict.Add($"<ISCAR{count}>", item.VisibleIsCar);
                    dict.Add($"<DATESTART{count}>", item.VisibleDateStart.ToString("d"));
                    dict.Add($"<TIMESTART{count}>", item.VisibleTimeStart.ToString("t"));
                    dict.Add($"<TIMEEND{count}>", item.VisibleTimeEnd.ToString("t"));
                    dict.Add($"<MINUTES{count}>", item.Time.ToString());
                    dict.Add($"<SUM{count}>", item.Sum.ToString());

                    var itemsInfo = items.Where(x => x.Fullname == item.Fullname).ToList();
                    if (!itemsInfo.Any(x => x.IsPrint))
                    {
                        info += $"{item.Fullname} отработано {itemsInfo.Sum(x => x.Time) / 60} часов {itemsInfo.Sum(x => x.Time) % 60} минут - сумма к оплате {itemsInfo.Sum(x => x.Sum)} рублей{Environment.NewLine}";
                        item.IsPrint = true;
                    }
                    count++;
                }
                dict.Add("<INFO>", info);
                dict.Add("<TOTALTIME>", TbTotalTime.Text);
                dict.Add("<TOTALSUM>", TbTotalSum.Text);

                for (int i = 0; i < 22; i++)
                {
                    dict.Add($"<ID{count}>", "");
                    dict.Add($"<FULLNAME{count}>", "");
                    dict.Add($"<ISCAR{count}>", "");
                    dict.Add($"<DATESTART{count}>", "");
                    dict.Add($"<TIMESTART{count}>", "");
                    dict.Add($"<TIMEEND{count}>", "");
                    dict.Add($"<MINUTES{count}>", "");
                    dict.Add($"<SUM{count}>", "");
                    count++;
                }
            });
            return dict;
        }

        private void FindAndReplace(Word._Application app, Dictionary<string, string> words)
		{
			try
			{
				var missing = Type.Missing;
				foreach (var item in words)
				{
					var find = app.Selection.Find;
					find.Text = item.Key;
					find.Replacement.Text = item.Value.Length >= 255 ? $"{item.Value.Substring(0, 250)}..." : item.Value;
					object wrap = Word.WdFindWrap.wdFindContinue;
					object replace = Word.WdReplace.wdReplaceAll;

					find.Execute(Type.Missing, false, false, false, missing, false, true, wrap, false, missing, replace);
				}
			}
			catch (Exception exception)
			{
				MessageBox.Show($"Ошибка {exception.Message}\n\n\nВнутреннее исключение{exception.InnerException}");
			}
		}

        private void MenuItem_OnClick(object sender, RoutedEventArgs e)
        {
            new ChangeCost().Show();
        }
    }
}
