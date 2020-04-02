using System;
using System.IO;

namespace SyllabusChecker
{
    /// <summary>
    /// Для хранения и передачи путей к файлам, выбранных в главном окне
    /// </summary>
    public class InputData
    {
        /// <summary>
        /// Путь к файлу макета
        /// </summary>
        public string ModelPath { get; set; }

        /// <summary>
        /// Путь к проверяемому файлу
        /// </summary>
        public string SyllablePath { get; set; }

        /// <summary>
        /// Путь к папке для сохранения результата
        /// </summary>
        public string ResultFolderPath { get; set; }

        public InputData()
        {
            //Чтение ранее использованных в програме местоположений файлов
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            StreamReader sr;
            try
            {
                sr = new StreamReader(Path.Combine(appDataPath, "paths.txt"));
                ModelPath = sr.ReadLine();
                SyllablePath = sr.ReadLine();
                ResultFolderPath = sr.ReadLine();
                sr.Close();
            }
            catch (FileNotFoundException)
            {
                StreamWriter sw = new StreamWriter(Path.Combine(appDataPath, "paths.txt"));
                sw.WriteLine("");
                sw.WriteLine("");
                sw.WriteLine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                sw.Close();
                ModelPath = "";
                SyllablePath = "";
                ResultFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
        }
    }
}
