using System.Diagnostics;
using System.Windows;

namespace SyllabusChecker
{
    /// <summary>
    /// Логика взаимодействия для ResultWindow.xaml
    /// </summary>
    public partial class ResultWindow : Window
    {
        private readonly string DocPath;

        /// <summary>
        /// Окно с результатами проверки
        /// </summary>
        /// <param name="path">Путь к файлу с результатами</param>
        /// <param name="errorCount">Количество найденных ошибок</param>
        public ResultWindow(string path, int errorCount)
        {
            InitializeComponent();
            TbPath.Text = path;
            DocPath = path;
            TbErrorsCount.Text = errorCount.ToString();
        }

        /// <summary>
        /// Обработка нажатия на кнопку "Показать результат"
        /// </summary>
        private void BtnShowResult_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(DocPath);
            Close();
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e) => Close();
    }
}
