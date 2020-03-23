using System;
using System.IO;
using System.Windows;
using System.Windows.Forms;
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
            TbModelPath.Text = InputData.ModelPath;
            TbSyllablePath.Text = InputData.SyllablePath;
            TbResultFolderPath.Text = InputData.ResultFolderPath;
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
                SelectedPath = InputData.ResultFolderPath
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
            if (!(new FileInfo(InputData.ModelPath).Exists) || (InputData.ModelPath == ""))
            {
                MessageBox.Show("Файл модели рабочей программы не выбран или не существует");
                return;
            }
            else if (!(new FileInfo(InputData.SyllablePath).Exists) || (InputData.SyllablePath == ""))
            {
                MessageBox.Show("Проверяемый файл рабочей программы не выбран или не существует");
                return;
            }

            Checker checker;
            //Запуск проверки
            try
            {
                checker = new Checker(InputData);
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message + "\nЗакройте все приложения, использующие данный файл, чтобы продолжить.");
                return;
            }
            //checker.checkParagraphEquality();
        }

        //При закрытии программы сохраняет выбранные пути
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            StreamWriter sw = new StreamWriter(Path.Combine(appDataPath, "paths.txt"), false);
            sw.WriteLine(InputData.ModelPath);
            sw.WriteLine(InputData.SyllablePath);
            sw.WriteLine(InputData.ResultFolderPath);
            sw.Close();
        }
    }
}
