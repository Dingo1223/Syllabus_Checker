using System;
using System.Windows;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;
using MessageBox = System.Windows.MessageBox;

namespace SyllabusChecker
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    /// 
    public partial class MainWindow : Window
    {
        private readonly InputData InputData;

        public MainWindow()
        {
            InitializeComponent();
            InputData = new InputData();
        }

        //Обработка нажатия на кнопку выбора расположения файла с макетом
        private void BtnSelectModelPath_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Doc or Docx files|*.doc;*.docx"
            };
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                InputData.ModelPath = ofd.FileName;
                TbModelPath.Text = ofd.FileName;
            }
        }

        //Обработка нажатия на кнопку выбора расположения готового файла рабочей программы
        private void BtnSelectSyllablePath_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Doc or Docx files|*.doc;*.docx"
            };
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                InputData.SyllablePath = ofd.FileName;
                TbSyllablePath.Text = ofd.FileName;
            }
        }

        //Обработка нажатия на кнопку выбора расположения проверенного файла рабочей программы
        private void BtnSelectResultFolder_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog
            {
                RootFolder = Environment.SpecialFolder.Desktop
            };
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                InputData.ResultFolderPath = fbd.SelectedPath;
                TbResultFolderPath.Text = fbd.SelectedPath;
            }
        }

        //Обработка нажатия на кнопку "Выполнить проверку"
        private void BtnCheckSyllableStart_Click(object sender, RoutedEventArgs e)
        {
            if (InputData.ModelPath == null)
            {
                MessageBox.Show("Не выбран файл модели рабочей программы!");
                return;
            }
            else if (InputData.SyllablePath == null)
            {
                MessageBox.Show("Не выбран проверяемый файл рабочей программы!");
                return;
            }
            else if (InputData.ResultFolderPath == null)
            {
                MessageBox.Show("Не выбран путь для сохранения результата!");
                return;
            }

            //Запуск проверки
            Checker checker = new Checker(InputData);
        }
    }
}
