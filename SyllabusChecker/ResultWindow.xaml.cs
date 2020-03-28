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

        public ResultWindow(string path, int errorCount)
        {
            InitializeComponent();
            TbPath.Text = path;
            DocPath = path;
            TbErrorsCount.Text = errorCount.ToString();
        }

        private void BtnShowResult_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(DocPath);
            Close();
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e) => Close();
    }
}
