using System;
using System.Configuration;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;
using System.Text.RegularExpressions;

namespace SyllabusChecker
{
    public class Checker
    {
        /// <summary>
        /// Макет доумента
        /// </summary>
        private DocX Model { get; set; }

        /// <summary>
        /// Проверяемый документ
        /// </summary>
        private DocX Syllabus { get; set; }

        /// <summary>
        /// Счётчик найденных ошибок
        /// </summary>
        public int ErrorsCount;

        /// <summary>
        /// Проверка документа в соответствии с макетом
        /// </summary>
        /// <param name="inputData">Выбранные пользователем расположения файлов</param>
        public Checker(InputData inputData)
        {
            Model = DocX.Load(inputData.ModelPath);
            Syllabus = DocX.Load(inputData.DocumentPath);

            //Проверка титульника
            Dictionary<int, string> IndsTitle = CheckTitlePage();

            //Получаем разбитые на секции модель и РП
            if (Model.Sections.Count < 2)
            {
                throw new Exception("Нарушен формат документа (возможно, вы пытаетесь проверить " +
                    "не рабочую программу).\nИспользуйте другой режим проверки.");
            }
            List<DocSection> ModelSections = GetDocSections(Model.Sections[1]);
            List<DocSection> SyllabusSections = GetDocSections(Syllabus.Sections[1]);
            CheckSections(ModelSections, ref SyllabusSections);

            //Проверка секций РП (кроме титульника)
            Dictionary<int, string> IndsBody = CheckSyllabusSections(ModelSections, SyllabusSections);

            //Создание результирующего документа
            CreateResultDoc(IndsTitle, IndsBody, inputData);
            ErrorsCount = IndsTitle.Count + IndsBody.Count;
        }

        /// <summary>
        /// Создание результирующего документа
        /// </summary>
        /// <param name="errorsTitle">Индексы параграфов с ошибками в титульнике</param>
        /// <param name="errorsBody">Индексы параграфов с ошибками в остальном документе</param>
        /// <param name="inputData">Выбранные пользователем расположения файлов</param>
        public void CreateResultDoc(Dictionary<int, string> errorsTitle, 
            Dictionary<int, string> errorsBody, InputData inputData)
        {
            string path = Path.Combine(inputData.ResultFolderPath,
                Path.GetFileNameWithoutExtension(inputData.DocumentPath) + "_checked.docx");

            //Размечаем титульник -- отмечаем ошибки жёлтой подсветкой
            foreach (int ind in errorsTitle.Keys)
            {
                Syllabus.Sections[0].SectionParagraphs[ind].Highlight(Highlight.yellow);
            }

            //Размечаем остальное -- отмечаем ошибки жёлтой подсветкой
            foreach (int ind in errorsBody.Keys)
            {
                Syllabus.Sections[1].SectionParagraphs[ind].Highlight(Highlight.yellow);
            }

            //Сохраняем новый файл РП, с подсветкой
            Syllabus.SaveAs(path);
            
            //Добавляем комментарии об ошибках
            DocComments.AddComments(errorsTitle, errorsBody, Syllabus.Sections[0].SectionParagraphs.Count, path);
        }

        /// <summary>
        /// Проверка титульника
        /// </summary>
        /// <returns>Dictionary с парами значений "индекс параграфа с ошибкой; описание ошибки"</returns>
        public Dictionary<int, string> CheckTitlePage()
        {
            Dictionary<int, string> errorsTitle = new Dictionary<int, string>();
            Section title_model = Model.Sections[0];
            Section title_syllabus = Syllabus.Sections[0];
            int ind = 0;

            for (int i = 0; i < title_model.SectionParagraphs.Count; i++)
            {
                if (title_model.SectionParagraphs[i].Text == "") continue;

                for (int j = ind; j < title_syllabus.SectionParagraphs.Count; j++)
                {
                    ind++;
                    if (title_syllabus.SectionParagraphs[j].Text == "") continue;

                    if (title_model.SectionParagraphs[i].Text != title_syllabus.SectionParagraphs[j].Text)
                    {
                        //Записываем номер параграфа, в котором ошибка
                        errorsTitle.Add(j, "Несовпадение с макетом, должно быть: " + 
                            title_model.SectionParagraphs[i].Text);
                    }
                    break;
                }
            }

            //Проверка колонтитула (номер должен совпадать)
            string model_footer = Model.Sections[0].Footers.Odd.Paragraphs[Model.Sections[0].Footers.Odd.Paragraphs.Count - 1].Text,
                syllabus_footer = Syllabus.Sections[0].Footers.Odd.Paragraphs[Syllabus.Sections[0].Footers.Odd.Paragraphs.Count - 1].Text;
            if (syllabus_footer != model_footer)
            {
                errorsTitle.Add(title_syllabus.SectionParagraphs.Count - 1, 
                    "Колонтитул не совпадает с макетом, должно быть: " + model_footer);
            }

            return errorsTitle;
        }

        /// <summary>
        /// Проверка рабочей программы, по разделам
        /// </summary>
        /// <param name="modelSections">Разбитый на секции макет документа</param>
        /// <param name="syllabusSections">Разбитый на секции проверяемый документ</param>
        /// <returns>Dictionary с парами значений "индекс параграфа с ошибкой; описание ошибки"</returns>
        public Dictionary<int, string> CheckSyllabusSections(List<DocSection> modelSections, List<DocSection> syllabusSections)
        {
            // Для описания ошибки в документе:
            // ФОРМАТ: (индекс_параграфа_в_Syllabus; описание_ошибки)
            Dictionary<int, string> errorsBody = new Dictionary<int, string>();

            //Section 0 = Рабочая программа рассмотрена и утверждена на заседании кафедры
            if (syllabusSections[0].Paragraphs[0].Text.Trim() != modelSections[0].Paragraphs[0].Text.Trim())
            {
                errorsBody.Add(syllabusSections[0].StartedAt, "Отсутствует раздел \"Рабочая программа рассмотрена и утверждена ...\"");
            }
            else
            {
                //Поскольку игнорируем пустые строки, избавляемся от них
                int tempForModel = 0, tempForSyllabus = 0;

                for (int i = 0; i < modelSections[0].Paragraphs.Count; i++)
                {
                    if (modelSections[0].Paragraphs[i].Text != "")
                        tempForModel++;
                }

                for (int i = 0; i < syllabusSections[0].Paragraphs.Count; i++)
                {
                    if (syllabusSections[0].Paragraphs[i].Text != "")
                        tempForSyllabus++;
                }

                string[,] model = new string[tempForModel, 3];
                string[,] syllabus = new string[tempForSyllabus, 2];

                for (int i = 0, j = 0; i < modelSections[0].Paragraphs.Count; i++)
                {
                    if (modelSections[0].Paragraphs[i].Text != "")
                    {
                        model[j, 0] = modelSections[0].Paragraphs[i].Text;
                        model[j, 1] = (modelSections[0].StartedAt + i).ToString();
                        j++;
                    }
                }

                for (int i = 0, j = 0; i < syllabusSections[0].Paragraphs.Count; i++)
                {
                    if (syllabusSections[0].Paragraphs[i].Text != "")
                    {
                        syllabus[j, 0] = syllabusSections[0].Paragraphs[i].Text;
                        syllabus[j, 1] = (syllabusSections[0].StartedAt + i).ToString();
                        j++;
                    }
                }

                //Проверяем, что все обязательные строки присутствуют
                int temp9 = 0;
                for (int i = 0; i < tempForModel; i++)
                {
                    if ((i >= 0 && i <= 7) || i == 9 || (i >= 11 && i <= 19) || i == 21 || i == 22 || i == 24)
                    {
                        int k = 0;
                        while (k < tempForSyllabus && model[i, 0] != syllabus[k, 0])
                            k++;
                        if (k < tempForSyllabus)
                        {
                            model[i, 2] = "1";

                            if (i == 9) temp9 = k;
                        }
                        else model[i, 2] = "0";
                    }
                }
                //Поскольку у нас есть одинаковые строки, то проверяем начилие вторых одинаковых строк отдельно
                {
                    int k = tempForSyllabus - 1;
                    while (k != -1 && model[11, 0] != syllabus[k, 0])
                        k--;
                    if (k != -1 && k != temp9) model[11, 2] = "1";
                    else model[11, 2] = "0";
                }
                {
                    int k = tempForSyllabus - 1;
                    while (k != -1 && model[21, 0] != syllabus[k, 0])
                        k--;
                    if (k != -1 && k != temp9) model[21, 2] = "1";
                    else  model[21, 2] = "0";
                }

                {
                    int i = 0, j = 0;
                    while (j < tempForSyllabus - 1)
                    {

                        if (model[i, 2] == "1" && model[i, 0] == syllabus[j, 0])
                        {
                            i++;
                            j++;
                        }
                        else
                        {
                            if (model[i, 2] == "0")
                            {
                                //Такой строки нет, говорим об этом. Будет отмечена первая строка в документе
                                if (!errorsBody.ContainsKey(int.Parse(syllabus[j, 1])))
                                {
                                    errorsBody.Add(int.Parse(syllabus[j, 1]), "Несовпадение с макетом, должно быть: " + model[i, 0]);
                                }
                                i++;
                            }
                            else
                            {
                                if (model[i, 2] == "1" && model[i, 0] != syllabus[j, 0])
                                {
                                    //Ошибка, что строка неправильная
                                    if (!errorsBody.ContainsKey(int.Parse(syllabus[j, 1])))
                                    {
                                        errorsBody.Add(int.Parse(syllabus[j, 1]), "Несовпадение с макетом, должно быть: " + model[i, 0]);
                                    }
                                    j++;
                                }
                            }
                        }
                        if (i == 23)
                        {
                            if (model[i, 0] == syllabus[j, 0])
                            {
                                //Ошибка, ФИО должно быть указано
                                if (!errorsBody.ContainsKey(int.Parse(syllabus[j, 1])))
                                {
                                    errorsBody.Add(int.Parse(syllabus[j, 1]), "Не указано ФИО");
                                }
                            }
                        }

                        if (model[i, 2] == null)
                        {
                            if (i == 10)
                            {
                                i++;
                                j++;
                            }
                            if (i == 23)
                            {
                                i++;
                                j++;
                            }
                            if (i == 8)
                            {
                                if (model[i, 0].Replace(" ", "") == syllabus[j, 0].Replace(" ", ""))
                                {
                                    //Ошибка, строки должны быть различны, т.е. поле заполнено
                                    if (!errorsBody.ContainsKey(int.Parse(syllabus[j, 1])))
                                    {
                                        errorsBody.Add(int.Parse(syllabus[j, 1]), "Не указан исполнитель");
                                    }
                                    i++;
                                    j++;
                                }
                                else
                                {
                                    i++;
                                    j++;
                                }
                            }
                            if (i == 20)
                            {
                                if (model[i, 0].Replace(" ", "") == syllabus[j, 0].Replace(" ", ""))
                                {
                                    //Ошибка, они должны быть различны, т.е. поле заполнено
                                    if (!errorsBody.ContainsKey(int.Parse(syllabus[j, 1])))
                                    {
                                        errorsBody.Add(int.Parse(syllabus[j, 1]), "Не указан уполномоченный");
                                    }
                                    i++;
                                    j++;
                                }
                                else
                                {
                                    i++;
                                    j++;
                                }
                            }
                        }
                    }
                }
            }

            //Section 1 = 1 Цели и задачи освоения дисциплины
            {
                int ind = 1;
                int indOfTargets = -1, indOfGoals = -1;

                //Ищем, где начинаются цели
                while (!syllabusSections[1].Paragraphs[ind].Text.Contains(" освоения дисциплины:") &&
                    (ind < syllabusSections[1].Paragraphs.Count))
                    ind++;

                if (ind >= syllabusSections[1].Paragraphs.Count)
                {
                    errorsBody.Add(syllabusSections[1].StartedAt, "Нет пункта \"Цели освоения дисциплины\"");
                }
                else
                {
                    indOfTargets = syllabusSections[1].StartedAt + ind;
                    ind++;
                }

                bool hasTargets = false, hasGoals = false;

                //Ищем прописанные цели; ищем, где начинаются задачи
                while ((ind < syllabusSections[1].Paragraphs.Count) &&
                    !syllabusSections[1].Paragraphs[ind].Text.Contains("Задачи:"))
                {
                    if (!hasTargets && (syllabusSections[1].Paragraphs[ind].Text != ""))
                        hasTargets = true;
                    ind++;
                }

                if (ind >= syllabusSections[1].Paragraphs.Count)
                {
                    errorsBody.Add(syllabusSections[1].StartedAt, "Нет пункта \"Задачи\"");
                }
                else
                {
                    indOfGoals = syllabusSections[1].StartedAt + ind;
                    ind++;
                }

                //Ищем прописанные задачи
                while (ind < syllabusSections[1].Paragraphs.Count)
                {
                    if (!hasGoals && (syllabusSections[1].Paragraphs[ind].Text != ""))
                    {
                        hasGoals = true;
                        break;
                    }
                    ind++;
                }

                if (!hasTargets && (indOfTargets != -1))
                {
                    errorsBody.Add(indOfTargets, "Не заполнен пункт \"Цели освоения дисциплины\"");
                }

                if (!hasGoals && (indOfGoals != -1))
                {
                    errorsBody.Add(indOfGoals, "Не заполнен пункт \"Задачи\"");
                }
            }

            //Section 2 = 2 Место дисциплины в структуре образовательной программы
            {
                Dictionary<int, string> mod = new Dictionary<int, string>(),
                    syl = new Dictionary<int, string>();

                for (int i = 1; i < modelSections[2].Paragraphs.Count; i++)
                {
                    if (modelSections[2].Paragraphs[i].Text != "")
                    {
                        mod.Add(i + modelSections[2].StartedAt, modelSections[2].Paragraphs[i].Text);
                    }
                }

                for (int i = 1; i < syllabusSections[2].Paragraphs.Count; i++)
                {
                    if (syllabusSections[2].Paragraphs[i].Text != "")
                    {
                        syl.Add(i + syllabusSections[2].StartedAt, syllabusSections[2].Paragraphs[i].Text);
                    }
                }

                if (mod.Count == syl.Count)
                {
                    for (int i = 0; i < syl.Count; i++)
                    {
                        if (syl.Values.ElementAt(i) != mod.Values.ElementAt(i))
                        {
                            errorsBody.Add(syl.Keys.ElementAt(i), "Несовпадение с макетом, должно быть: " + mod.Values.ElementAt(i));
                        }
                    }
                }
                else
                {
                    errorsBody.Add(syllabusSections[2].StartedAt, "Раздел не заполнен полностью");
                }
            }

            //Section 3 = 3 Требования к результатам обучения по дисциплине
            {
                int j = 0;
                for (int i = 0; i < syllabusSections[3].Paragraphs.Count - 1; i++)
                {
                    if (i > 0 && modelSections[3].Paragraphs[i - 1].FollowingTables != null)
                    {
                        int pId = i - 1;
                        if (modelSections[3].Paragraphs[pId].FollowingTables[0].ColumnCount == 2)//если 2 колонки
                        {
                            for (int y = 0; y < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows.Count; y++)//по строкам
                            {
                                for (int l = 0; l < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells.Count; l++)//по ячейкам в строке
                                {
                                    if (y == 0)
                                    {
                                        for (int g = 0; g < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                        {
                                            if (modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text != 
                                                syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text)
                                            {
                                                errorsBody.Add(syllabusSections[3].StartedAt + i, "Несовпадение с макетом, должно быть " + 
                                                    modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text);
                                            }
                                            i++;
                                        }
                                    }
                                    else if (syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Width > 300)
                                    {
                                        int n = 0;
                                        for (int g = 0; g < syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                        {
                                            if (syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text.Contains("Знать:")
                                                || syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text.Contains("Уметь:")
                                                || syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text.Contains("Владеть:"))
                                                n++;
                                            i++;
                                        }
                                        if (n != 3)
                                        {
                                            errorsBody.Add(syllabusSections[3].StartedAt + i - 1, 
                                                "Несовпадение с макетом, отсутствуют некоторые из элементов: 'Знать:' 'Уметь:' 'Владеть:'");
                                            //for (int g = 0; g < syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                            //{
                                            //    {
                                            //        errorsBody.Add(syllableSections[3].StartedAt + i - 
                                            //            syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count + g, 
                                            //            "Несовпадение с макетом, отсутствуют некоторые из элементов: 'Знать:' 'Уметь:' 'Владеть:'");
                                            //    }
                                            //}
                                        }
                                    }
                                    else
                                    {
                                        int maxSylMod = Math.Max(modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count, 
                                            syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count);
                                        for (int m = 0; m < maxSylMod; m++)
                                        {
                                            for (int g = 0; g < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                            {
                                                if (modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text != 
                                                    syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text)
                                                {
                                                    errorsBody.Add(syllabusSections[3].StartedAt + i, "Несовпадение с макетом, должно быть " + 
                                                        modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text);
                                                    m++;
                                                    i++;
                                                }
                                                while (m < maxSylMod - 1)
                                                {
                                                    errorsBody.Add(syllabusSections[3].StartedAt + i - 1, "Несовпадение с макетом");
                                                    i++;
                                                    m++;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else if (modelSections[3].Paragraphs[pId].FollowingTables[0].ColumnCount == 3)//если 3 колонки
                        {
                            for (int y = 0; y < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows.Count; y++)//бегаем по строкам
                            {
                                for (int l = 0; l < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells.Count; l++)//бегаем по ячейкам в строке
                                {
                                    if (y == 0)
                                    {
                                        for (int g = 0; g < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                        {
                                            if (modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text != 
                                                syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text)
                                            {
                                                errorsBody.Add(syllabusSections[3].StartedAt + i, "Несовпадение с макетом, должно быть " + 
                                                    modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text);
                                            }
                                            i++;
                                        }
                                    }
                                    else if (syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Width < 150)
                                    {
                                        int n = 0;
                                        for (int g = 0; g < syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                        {
                                            if (syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text.Contains("Знать:")
                                                || syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text.Contains("Уметь:")
                                                || syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text.Contains("Владеть:"))
                                                n++;
                                            i++;
                                        }
                                        if (n != 3)
                                        {
                                            errorsBody.Add(syllabusSections[3].StartedAt + i - 1, 
                                                "Несовпадение с макетом, отсутствуют некоторые из элементов: 'Знать:' 'Уметь:' 'Владеть:'");
                                            //for (int g = 0; g < syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                            //{
                                            //    errorsBody.Add(syllableSections[3].StartedAt + i - 
                                            //        syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count + g, 
                                            //        "Несовпадение с макетом, отсутствуют некоторые из элементов: 'Знать:' 'Уметь:' 'Владеть:'");
                                            //}
                                        }
                                    }
                                    else
                                    {
                                        if (modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count > 
                                            syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count)
                                        {
                                            for (int g = 0; g < syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                            {
                                                errorsBody.Add(syllabusSections[3].StartedAt + i, "Несовпадение с макетом");
                                                i++;
                                            }
                                        }
                                        else if (modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count < 
                                            syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count)
                                        {
                                            for (int g = 0; g < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                            {
                                                if (modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text != 
                                                    syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text)
                                                {
                                                    errorsBody.Add(syllabusSections[3].StartedAt + i, "Несовпадение с макетом, должно быть " + 
                                                        modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text);
                                                    i++;
                                                }
                                            }
                                            int m = 0;
                                            while (m < syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count - modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count)
                                            {
                                                errorsBody.Add(syllabusSections[3].StartedAt + i, "Несовпадение с макетом");
                                                m++;
                                                i++;
                                            }
                                        }
                                        else
                                        {
                                            for (int g = 0; g < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                            {
                                                if (modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text != 
                                                    syllabusSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text)
                                                {
                                                    errorsBody.Add(syllabusSections[3].StartedAt + i, 
                                                        "Несовпадение с макетом, должно быть " + 
                                                        modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text);
                                                    i++;
                                                }
                                                else i++;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        bool b = false;
                        while (b == false)
                        {
                            if (syllabusSections[3].Paragraphs[i + 1].FollowingTables != null) b = true;
                            else
                            {
                                if (syllabusSections[3].Paragraphs[i + 1].Text != modelSections[3].Paragraphs[i + 1].Text)
                                {
                                    errorsBody.Add(syllabusSections[3].StartedAt + i + 1, "Несовпадение с макетом, должно быть " + 
                                        modelSections[3].Paragraphs[i + 1].Text);
                                }
                            }
                            i++;
                        }
                    }
                    j++;
                }
            }

            //Section 4 = 4 Структура и содержание дисциплины
            // Должна быть пуста
            if (IsSectionHasContent(syllabusSections[4]))
            {
                errorsBody.Add(syllabusSections[4].StartedAt, "В секции не должно быть содержимого");
            }

            //Section 5 = 4.1 Структура дисциплины
            {
                List<Table> modelTables = Model.Tables;
                List<Table> syllabusTables = Syllabus.Tables;
                int numParagraphInModel = 0, numOfTable = 3, ind = 0;
                for (int i = 0; i < syllabusSections[5].Paragraphs.Count - 1; i++)//по параграфам в РП
                {
                    if (i == 0)
                    {
                        if (syllabusSections[5].Paragraphs[i].Text != modelSections[5].Paragraphs[numParagraphInModel].Text)
                        {
                            errorsBody.Add(syllabusSections[5].StartedAt + i, "Несовпадение с макетом, должно быть: " + 
                                modelSections[5].Paragraphs[numParagraphInModel].Text);
                        }
                    }
                    else if (syllabusSections[5].Paragraphs[i - 1].FollowingTables != null) //если начинается таблица
                    {
                        if (numOfTable == 3) //если первая таблица
                        {
                            for (int j = 0; j < syllabusTables[numOfTable].Rows.Count; j++)
                            {
                                for (int f = 0; f < syllabusTables[numOfTable].Rows[j].Cells.Count; f++)
                                {
                                    for (int h = 0; h < syllabusTables[numOfTable].Rows[j].Cells[f].Paragraphs.Count; h++)
                                    {
                                        if (j == modelTables[numOfTable].Rows.Count - 2)
                                        {
                                            i++;
                                            while (ind < modelTables[numOfTable].Rows[j].Paragraphs.Count)
                                            {
                                                numParagraphInModel++;
                                                ind++;
                                            }
                                        }
                                        else if (syllabusTables[numOfTable].Rows[j].Cells[f].Paragraphs[h].Text != 
                                            modelTables[numOfTable].Rows[j].Cells[f].Paragraphs[h].Text)
                                        {
                                            i++;
                                            numParagraphInModel++;
                                            errorsBody.Add(syllabusSections[5].StartedAt + i - 1, "Несовпадение с макетом, должно быть: " + 
                                                modelTables[numOfTable].Rows[j].Cells[f].Paragraphs[h].Text);
                                        }
                                        else
                                        {
                                            i++;
                                            numParagraphInModel++;
                                        }
                                    }
                                }
                            }
                            numOfTable++;
                        }
                        else if (numOfTable == 4 && modelTables.Count == 7)//если вторая таблица
                        {
                            bool skip = false;
                            int jM = 0;
                            for (int j = 0; j < syllabusTables[numOfTable].Rows.Count; j++)
                            {
                                for (int g = 0; g < syllabusTables[numOfTable].Rows[j].Cells.Count; g++)
                                {
                                    for (int f = 0; f < syllabusTables[numOfTable].Rows[j].Cells[g].Paragraphs.Count; f++)
                                    {
                                        if (j > 2 && j < syllabusTables[numOfTable].Rows.Count - 1)
                                        {
                                            i++;
                                            for (int d = 3; (d < modelTables[numOfTable].Rows.Count - 1) && skip == false; d++)
                                            {
                                                jM++;
                                                for (int par = 0; par < modelTables[numOfTable].Rows[d].Paragraphs.Count; par++)
                                                    numParagraphInModel++;
                                            }
                                            skip = true;
                                        }
                                        else if (syllabusTables[numOfTable].Rows[j].Cells[g].Paragraphs[f].Text != 
                                            modelTables[numOfTable].Rows[jM].Cells[g].Paragraphs[f].Text)
                                        {
                                            i++;
                                            numParagraphInModel++;
                                            errorsBody.Add(syllabusSections[5].StartedAt + i - 1, "Несовпадение с макетом, должно быть: " + 
                                                modelTables[numOfTable].Rows[jM].Cells[g].Paragraphs[f].Text);
                                        }
                                        else
                                        {
                                            i++;
                                            numParagraphInModel++;
                                        }
                                    }
                                }
                                if (!(j > 2 && j < syllabusTables[numOfTable].Rows.Count - 1)) jM++;
                            }
                            numOfTable++;
                        }

                        else if (numOfTable == modelTables.Count - 2)//если третья таблица
                        {
                            bool skip = false;
                            int jM = 0;
                            for (int j = 0; j < syllabusTables[numOfTable].Rows.Count; j++)
                            {
                                for (int g = 0; g < syllabusTables[numOfTable].Rows[j].Cells.Count; g++)
                                {
                                    for (int f = 0; f < syllabusTables[numOfTable].Rows[j].Cells[g].Paragraphs.Count; f++)
                                    {
                                        if (j > 2 && j < syllabusTables[numOfTable].Rows.Count - 2)
                                        {
                                            i++;
                                            for (int d = 3; (d < modelTables[numOfTable].Rows.Count - 2) && skip == false; d++)
                                            {
                                                jM++;
                                                for (int par = 0; par < modelTables[numOfTable].Rows[d].Paragraphs.Count; par++)
                                                    numParagraphInModel++;
                                            }
                                            skip = true;
                                        }
                                        else if (syllabusTables[numOfTable].Rows[j].Cells[g].Paragraphs[f].Text != 
                                            modelTables[numOfTable].Rows[jM].Cells[g].Paragraphs[f].Text)
                                        {
                                            i++;
                                            numParagraphInModel++;
                                            errorsBody.Add(syllabusSections[5].StartedAt + i - 1, "Несовпадение с макетом, должно быть: "
                                                + modelTables[numOfTable].Rows[jM].Cells[g].Paragraphs[f].Text);
                                        }
                                        else
                                        {
                                            i++;
                                            numParagraphInModel++;
                                        }
                                    }
                                }
                                if (!(j > 2 && j < syllabusTables[numOfTable].Rows.Count - 2)) jM++;
                            }
                            numOfTable++;
                        }
                    }
                    else
                    {
                        if (syllabusSections[5].Paragraphs[i].Text != modelSections[5].Paragraphs[numParagraphInModel].Text)
                        {
                            errorsBody.Add(syllabusSections[5].StartedAt + i - 1, "Несовпадение с макетом, должно быть: " + 
                                modelSections[5].Paragraphs[numParagraphInModel].Text);
                        }
                    }
                    numParagraphInModel++;
                }
            }

            //Section 6 = 4.2 Содержание разделов дисциплины
            //Должна быть заполнена
            {
                if (IsSectionHasContent(syllabusSections[6]))
                {
                    int counter = 0;
                    for (int i = 1; i < syllabusSections[6].Paragraphs.Count; i++)
                    {
                        if (syllabusSections[6].Paragraphs[i].Text != "")
                        {
                            int x;
                            //Проверяем нумерацию разделов
                            if (int.TryParse(syllabusSections[6].Paragraphs[i].Text.Split(' ')[0], out x))
                            {
                                if (x != counter + 1)
                                {
                                    errorsBody.Add(syllabusSections[6].StartedAt + i, "Нарушена нумерация разделов");
                                }
                                counter = x;
                            }
                        }
                    }
                }
                else
                {
                    errorsBody.Add(syllabusSections[6].StartedAt, "Секция не заполнена");
                }
            }

            //Section 7 = 4.3 Практические занятия(семинары)
            {
                int spaces_in_begin = 1, spaces_in_end = 0, spaces_in_end_model = 0;
                bool hasTable = true;

                //Считаем, есть ли пустые строки перед таблицей, чтобы если что их пропустить
                while (!syllabusSections[7].Paragraphs[spaces_in_begin].Text.Contains("№ занятия"))
                {
                    spaces_in_begin++;
                    if (spaces_in_begin >= syllabusSections[7].Paragraphs.Count)
                    {
                        hasTable = false;
                        break;
                    }
                }

                if (hasTable)
                {
                    //Считаем, есть ли пустые строки после таблицы, чтобы если что их пропустить
                    int ind = syllabusSections[7].Paragraphs.Count - 1;
                    while (syllabusSections[7].Paragraphs[ind].Text == "")
                    {
                        spaces_in_end++;
                        ind--;
                    }

                    //Считаем, есть ли пустые строки после таблицы в макете, чтобы если что их пропустить
                    int ind_model = modelSections[7].Paragraphs.Count - 1;
                    while (modelSections[7].Paragraphs[ind_model].Text == "")
                    {
                        spaces_in_end_model++;
                        ind_model--;
                    }

                    int cols = int.Parse(ConfigurationManager.AppSettings["cols_in_4_3_table"]);
                    //Заголовки в таблице должны совпадать с моделью
                    for (int i = spaces_in_begin; i < spaces_in_begin + cols; i++)
                    {
                        if (syllabusSections[7].Paragraphs[i].Text != modelSections[7].Paragraphs[i].Text)
                        {
                            errorsBody.Add(syllabusSections[7].StartedAt + i,
                                "Несовпадение с макетом, должно быть: " + modelSections[7].Paragraphs[i].Text);
                        }
                    }

                    //Проверяем строчки в таблице
                    int ind_lesson = 0, sum = 0;
                    for (int i = spaces_in_begin + 4; i < syllabusSections[7].Paragraphs.Count - cols - spaces_in_end; i += cols)
                    {
                        int s = 0, x, _;
                        bool isCorrect = int.TryParse(syllabusSections[7].Paragraphs[i].Text, out x) &&
                            int.TryParse(syllabusSections[7].Paragraphs[i + 1].Text, out _) &&
                            (syllabusSections[7].Paragraphs[i + 2].Text != "") &&
                            int.TryParse(syllabusSections[7].Paragraphs[i + 3].Text, out s) &&
                            (x == ind_lesson + 1);

                        ind_lesson = x;
                        sum += s;

                        //Если где-то в строке ошибка -- подсвечиваем всю строку
                        if (!isCorrect)
                        {
                            errorsBody.Add(syllabusSections[7].StartedAt + i, "Ошибка в данной строке таблицы");
                        }
                    }

                    int sum_syl;
                    if (int.TryParse(syllabusSections[7].Paragraphs[syllabusSections[7].Paragraphs.Count -
                        spaces_in_end - 1].Text, out sum_syl))
                    {
                        int sum_model = int.Parse(modelSections[7].Paragraphs[modelSections[7].Paragraphs.Count -
                            spaces_in_end_model - 1].Text);

                        //Если сумма в РП не совпадает с суммой в макете
                        // или если итого по таблице не сходится = ошибка
                        if ((sum_syl != sum_model) || (sum_syl != sum))
                        {
                            errorsBody.Add(syllabusSections[7].StartedAt + syllabusSections[7].Paragraphs.Count - spaces_in_end - 1,
                                "Сумма часов не совпадает с макетом либо не сходится");
                        }
                    }
                    else
                    {
                        errorsBody.Add(syllabusSections[7].StartedAt + syllabusSections[7].Paragraphs.Count - spaces_in_end - 1,
                            "Не указана сумма часов");
                    }
                }
                else
                {
                    errorsBody.Add(syllabusSections[7].StartedAt, "Раздел заполнен неверно");
                }
            }

            //Section 8 = 5 Учебно - методическое обеспечение дисциплины
            // Должна быть пуста
            if (IsSectionHasContent(syllabusSections[8]))
            {
                errorsBody.Add(syllabusSections[8].StartedAt, "В секции не должно быть содержимого");
            }

            //Section 9 = 5.1 Основная литература
            //Section 10 = 5.2 Дополнительная литература
            //Section 11 = 5.3 Периодические издания
            //Section 12 = 5.4 Интернет - ресурсы
            //Section 13 = 5.5 Программное обеспечение, профессиональные базы данных и информационные справочные системы
            // Должны быть заполнены (не можем проверить содержимое дословно)
            for (int i = 9; i <= 13; i++)
            {
                if (!IsSectionHasContent(syllabusSections[i]))
                {
                    errorsBody.Add(syllabusSections[i].StartedAt, "Секция не заполнена");
                }
            }

            //Section 14 = 6 Материально - техническое обеспечение дисциплины
            {
                //Сравниваем абзацы, которые должны совпадать с макетом
                try
                {
                    int pars = int.Parse(ConfigurationManager.AppSettings["important_pars_in_6"]);
                    for (int i = 1; i <= pars; i++)
                    {
                        if (syllabusSections[14].Paragraphs[i].Text != modelSections[14].Paragraphs[i].Text)
                        {
                            errorsBody.Add(syllabusSections[14].StartedAt + i,
                                "Несовпадение с макетом, должно быть: " + modelSections[14].Paragraphs[i].Text);
                        }
                    }
                }
                catch //Если не хватает обязательных абзацев
                {
                    errorsBody.Add(syllabusSections[14].StartedAt, "Раздел не заполнен полностью");
                }
            }

            return errorsBody;
        }

        /// <summary>
        /// Для лёгких случаев, когда нужно проверить только заполнена ли секция
        /// </summary>
        /// <param name="section">Секция документа, которую нужно проверить</param>
        /// <returns>True, если в секции есть какой-либо текст; False - если нет</returns>
        private bool IsSectionHasContent(DocSection section)
        {
            bool hasText = false;
            for (int i = 1; i < section.Paragraphs.Count; i++)
            {
                if (section.Paragraphs[i].Text != "")
                {
                    hasText = true;
                    break;
                }
            }

            return hasText;
        }

        /// <summary>
        /// Разбиение документа на секции по разделам в документе
        /// </summary>
        /// <param name="doc">Документ для разбиения</param>
        /// <returns>Список секций; каждая секция - заголовок секции + всё, что после него и до следующего заголовка</returns>
        public List<DocSection> GetDocSections(Section doc)
        {
            //Получаем заголовки в doc
            List<Paragraph> headers = doc.SectionParagraphs.Where(x => x.IsKeepWithNext == true).ToList();
            List<string> names = new List<string>();
            foreach (Paragraph p in headers)
            {
                //^\d+(\.\d+)?[^a-zA-Z0-9_а-яА-Я]+
                string header = p.Text.Trim();
                if (Regex.IsMatch(header, @"^\d+(\.\d+)?[^a-zA-Z0-9_а-яА-Я]+"))
                    names.Add(header);
            }
            if (names.Count == 0)
            {
                throw new Exception("Критическая ошибка: документ не разбит на подразделы." +
                    "\nВозможно, вы хотите использовать другой режим проверки?");
            }

            //Разбиваем doc на разделы
            List<DocSection> docSections = new List<DocSection>();
            int ind = 0;

            //Получаем первый раздел -- "Рабочая программа рассмотрена и утверждена..."
            int f = 0;
            while (doc.SectionParagraphs[f].Text.Trim() == "") f++;
            DocSection firstSection = new DocSection
            {
                Paragraphs = new List<Paragraph>(),
                StartedAt = f
            };
            while (doc.SectionParagraphs[f].Text.Trim() != names[ind])
            {
                firstSection.Paragraphs.Add(doc.SectionParagraphs[f]);
                f++;
            }
            firstSection.EndedAt = f - 1;
            docSections.Add(firstSection);
            ind++;

            //Получаем осталььные разделы
            for (int i = f; i < doc.SectionParagraphs.Count; i++)
            {
                //Проверяем, достигли ли последнего раздела
                bool isDocEnding = ind >= names.Count;

                int startedAt = i;
                List<Paragraph> paragraphs = new List<Paragraph>();
                if (isDocEnding)
                {
                    //Если достигли последнего раздела -- записываем все параграфы до самого конца документа
                    while (i < doc.SectionParagraphs.Count)
                    {
                        paragraphs.Add(doc.SectionParagraphs[i]);
                        i++;
                    }
                }
                else
                {
                    //Если раздел не последний -- записываем все параграфы между двумя соседними разделами
                    while (doc.SectionParagraphs[i].Text.Trim() != names[ind])
                    {
                        paragraphs.Add(doc.SectionParagraphs[i]);
                        i++;
                    }
                }
                i--;
                ind++;
                docSections.Add(new DocSection(startedAt, i, paragraphs));
            }
            return docSections;
        }

        public void CheckSections(List<DocSection> model, ref List<DocSection> syllabus)
        {
            for (int i = 1; i < model.Count; i++)
            {
                Regex regex = new Regex(@"^\d+(\.\d+)?");
                double mod_num = double.Parse(regex.Match(model[i].Paragraphs[0].Text.Trim()).Value);
                bool hasSection = false;
                for (int j = i; j < syllabus.Count; j++)
                {
                    double syl_num = double.Parse(regex.Match(syllabus[j].Paragraphs[0].Text.Trim()).Value);
                    if (syl_num != mod_num)
                    {
                        syllabus[j - 1].Paragraphs.AddRange(syllabus[j].Paragraphs);
                        syllabus.RemoveAt(j);
                    }
                    else
                    {
                        hasSection = true;
                        break;
                    }
                    if (!hasSection) j--;
                }
                if (!hasSection)
                {
                    throw new Exception("Критическая ошибка: в проверяемом документе отсутствует раздел " +
                        mod_num.ToString() + ", присутствующий в макете.");
                }
            }
        }
    }
}
