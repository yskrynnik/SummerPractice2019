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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.IO;
using Microsoft.Win32;
using System.Xml;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace SummerPractice
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        // Для работы с исходным файлом
        FileStream fs;
        // Для выбора исходного файла
        OpenFileDialog opfile = new OpenFileDialog();
        // Для сохранения отчета Microsoft Excel
        SaveFileDialog svfile = new SaveFileDialog();
        public MainWindow()
        {
            InitializeComponent();
            // Указание фильтров для открытия и сохранения файлов
            opfile.Filter = "json files (*.json)|*.json";
            opfile.FilterIndex = 1;
            svfile.Filter = "xlsx file (*.xlsx)|*.xlsx";
            svfile.FilterIndex = 1;
        }
        private void Calculate_button_Click(object sender, RoutedEventArgs e)
        {
            if (opfile.ShowDialog() != false)
            {
                fs = new FileStream(opfile.FileName, FileMode.Open, FileAccess.Read);
                using (fs)
                {
                    double counter = 0; // Считает кол-во отличников для нахождения процентного соотношения
                    double stud_counter = 0; // Считает кол-во студентов для нахождения процентного соотношения
                    int tmp; // Хранит в себе текущую обрабатываемую оценку
                    bool noerrors = true; // Флаг если true обозначает что у студента нет оценки ниже 90 тоесть он отличник
                    string text = File.ReadAllText(opfile.FileName, Encoding.GetEncoding(1251)); // Считываем текст из исходного файла в кодировке 1251
                    text = text.Replace(",\"items\":\n", String.Empty);
                    text = text.Replace("}{", "},{");
                    text = text.Replace("]}", "]");
                    // Далее десериализуем и получаем список всех оценок, при помощи LINQ запроса разбиваем оценки на группы по id студентов.
                    // В итоге получаем список групп оценок для каждого студента, по сути список списков.
                    var group = JsonConvert.DeserializeObject<List<StudMark>>(text).GroupBy(f => f.id).Select(grp => grp.ToList()).ToList(); 
                    // Рассматриваем каждую группу и если у студента есть оценка ниже 90 то флаг noerrors устанивавливается в false и вся группа пропускается
                    // Иначе идет проверка всех оценок студента и если все нормально то счетчик добавляет единицу
                    foreach (List<StudMark> a in group)
                    {
                        stud_counter++;
                        foreach (StudMark element in a)
                        {
                            if (noerrors)
                            {
                                if (int.TryParse(element.name, out tmp) == true)
                                {
                                    if (tmp < 90) noerrors = false;
                                }
                                else
                                {
                                    noerrors = false;
                                }
                            }
                        }
                        if (noerrors) counter++;
                        noerrors = true;
                    }
                    // Работа с Microsoft Excel
                    // Открываем приложение
                    Excel.Application excelapplication = new Excel.Application();
                    // Создание книги
                    Excel.Workbook workBook = excelapplication.Workbooks.Add();
                    // Создание листа
                    Excel.Worksheet workSheet= (Excel.Worksheet)workBook.Worksheets.get_Item(1);
                    // Настройка имени листа
                    workSheet.Name = "Результат работы программы";
                    // Заполнение ячеек
                    workSheet.Cells[1, 1] = "Процент студентов которые учаться только на оценки «5»";
                    workSheet.Cells[1, 2] = $"{Math.Round(counter * 100 / stud_counter, 2)}%";
                    // Автовыравнивание колонок
                    workSheet.Columns.AutoFit();
                    if (svfile.ShowDialog() != false)
                    {
                        // Сохранение
                        excelapplication.Application.ActiveWorkbook.SaveAs(svfile.FileName, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        MessageBox.Show($"Отчет сгенерирован!","Информация",MessageBoxButton.OK,MessageBoxImage.Information);
                    }
                    // Закрытие приложения
                    excelapplication.Quit();
                }
            }   
        }
        private void AboutAuthor_Button_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show($"Программа считает процент студентов, из файла в формате .json, которые учатся только на пятерки, и сохраняет результат в виде отчета Microsoft Excel в формате .xslx\nАвтор:cтудент 525а группы Скрынник Егор", "О программе", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
