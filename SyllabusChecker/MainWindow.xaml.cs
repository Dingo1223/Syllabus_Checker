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
    public partial class MainWindow : Window
    {
        private readonly InputData InputData;

        public MainWindow()
        {
            InitializeComponent();
            InputData = new InputData();
            TbModelPath.Text = InputData.ModelPath;
            TbSyllablePath.Text = InputData.DocumentPath;
            TbResultFolderPath.Text = InputData.ResultFolderPath;
        }

        /// <summary>
        /// Обработка нажатия на кнопку выбора расположения файла с макетом
        /// </summary>
        private void BtnSelectModelPath_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Docx files|*.docx"
            };
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                InputData.ModelPath = ofd.FileName;
                TbModelPath.Text = ofd.FileName;
            }
        }

        /// <summary>
        /// Обработка нажатия на кнопку выбора расположения проверяемого документа
        /// </summary>
        private void BtnSelectSyllablePath_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Docx files|*.docx"
            };
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                InputData.DocumentPath = ofd.FileName;
                TbSyllablePath.Text = ofd.FileName;
            }
        }

        /// <summary>
        /// Обработка нажатия на кнопку выбора места сохранения проверенного документа
        /// </summary>
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

        /// <summary>
        /// Обработка нажатия на кнопку "Выполнить проверку"
        /// </summary>
        private void BtnCheckSyllableStart_Click(object sender, RoutedEventArgs e)
        {
            if (!(new FileInfo(InputData.ModelPath).Exists) || (InputData.ModelPath == ""))
            {
                MessageBox.Show("Файл модели документа не выбран или не существует");
                return;
            }
            else if (!(new FileInfo(InputData.DocumentPath).Exists) || (InputData.DocumentPath == ""))
            {
                MessageBox.Show("Проверяемый файл не выбран или не существует");
                return;
            }

            Checker checker;

            //Запуск проверки
            if (rbSyllable.IsChecked == true) //Если проверяется рабочая программа
            {
                try
                {
                    checker = new Checker(InputData);
                    ResultWindow rw = new ResultWindow(InputData.ResultFolderPath + "\\" +
                        Path.GetFileNameWithoutExtension(InputData.DocumentPath) + @"_checked.docx", checker.ErrorsCount);
                    rw.ShowDialog();
                }
                catch (IOException ex)
                {
                    MessageBox.Show(ex.Message + "\nЗакройте все приложения, использующие данный файл, чтобы продолжить.");
                    return;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
            else if (rbOther.IsChecked == true) //Если проверяется произвольный документ через подсветку
            {
                try
                {
                    HighlightHandler highlightHandler = new HighlightHandler();
                    int errors_count = highlightHandler.CheckDocumentsEquality(InputData);
                    ResultWindow rw = new ResultWindow(InputData.ResultFolderPath + "\\" +
                         Path.GetFileNameWithoutExtension(InputData.DocumentPath) + @"_checked.docx", errors_count);
                    rw.ShowDialog();
                }
                catch (IOException ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        /// <summary>
        /// При закрытии программы сохраняет выбранные пути
        /// </summary>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            StreamWriter sw = new StreamWriter(Path.Combine(appDataPath, "paths.txt"), false);
            sw.WriteLine(InputData.ModelPath);
            sw.WriteLine(InputData.DocumentPath);
            sw.WriteLine(InputData.ResultFolderPath);
            sw.Close();
        }
    }
}
