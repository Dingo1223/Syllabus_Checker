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

            var result = "";

            var document = DocX.Load(InputData.ModelPath);
            var sections = document.GetSections();
            result += "Секций: " + sections.Count.ToString() + " \n";
            for (int i = 0; i < sections.Count; i++)
            {
                result += "В секции " + i.ToString() + " параграфов: " + sections[i].SectionParagraphs.Count.ToString() + " \n";
                if (sections[i].SectionParagraphs.Count > 0)
                {
                    result += "И начинается она с: " + sections[i].SectionParagraphs[0].Text + " \n";
                }
                result += "Таблиц в ней: " + sections[i].Tables.Count + " \n";
                if (sections[i].Tables.Count > 0)
                {
                    result += "Пример доставания ячейки: " + sections[i].Tables[0].Rows[0].Cells[0] + " \n";
                }
            }

            MessageBox.Show(result);

            for (int i = 0; i < sections.Count; i++)
            {
                var content = "";
                for (int j = 0; j < sections[i].SectionParagraphs.Count; j++)
                {
                    if (sections[i].SectionParagraphs[j].Text.Length > 0)
                    {
                        content += sections[i].SectionParagraphs[j].Text + " " + j.ToString() + "\n";
                    }
                }
                MessageBox.Show(content);
            }

            //Здесь -- запуск проверки
        }
    }
}
