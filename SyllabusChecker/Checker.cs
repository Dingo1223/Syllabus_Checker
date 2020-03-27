using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;
using MessageBox = System.Windows.MessageBox;

namespace SyllabusChecker
{
    public class Checker
    {
        private DocX Model { get; set; }
        private DocX Syllable { get; set; }

        //Счётчик найденных ошибок
        public int ErrorsCount;

        public Checker(InputData inputData)
        {
            Model = DocX.Load(inputData.ModelPath);
            Syllable = DocX.Load(inputData.SyllablePath);
            List<int> IndsTitle = CheckTitlePage();

            //Получаем разбитые на секции модель и РП
            List<DocSection> ModelSections = GetDocSections(Model.Sections[1]);
            List<DocSection> SyllableSections = GetDocSections(Syllable.Sections[1]);

            //Дальше надо их проверять
            List<int> IndsBody = CheckSyllableSections(ModelSections, SyllableSections);

            //Создание результирующего документа
            CreateResultDoc(IndsTitle, IndsBody, inputData);
            ErrorsCount = IndsTitle.Count + IndsBody.Count;
        }

        //Создание результирующего документа (с подсветкой ошибочных мест)
        //Строится на основе Syllable
        //Параметр indsTitle -- индексы параграфов с ошибками в титульнике
        //Параметр indsBody -- индексы параграфов с ошибками в остальном документе
        public void CreateResultDoc(List<int> indsTitle, List<int> indsBody, InputData inputData)
        {
            string path = Path.Combine(inputData.ResultFolderPath,
                Path.GetFileNameWithoutExtension(inputData.SyllablePath) + "_checked.docx");

            //Размечаем титульник
            foreach (int ind in indsTitle)
                Syllable.Sections[0].SectionParagraphs[ind].Highlight(Highlight.red);

            //Размечаем остальное
            foreach (int ind in indsBody)
                Syllable.Sections[1].SectionParagraphs[ind].Highlight(Highlight.red);

            Syllable.SaveAs(path);
        }

        //Проверка титульника
        public List<int> CheckTitlePage()
        {
            List<int> indsTitle = new List<int>();
            Section title_model = Model.Sections[0];
            Section title_syllable = Syllable.Sections[0];
            int ind = 0;

            for (int i = 0; i < title_model.SectionParagraphs.Count; i++)
            {
                Paragraph par_model = title_model.SectionParagraphs[i];
                if (par_model.Text == "") continue;
                for (int j = ind; j < title_model.SectionParagraphs.Count; j++)
                {
                    ind++;
                    Paragraph par_syllable = title_syllable.SectionParagraphs[j];
                    if (par_syllable.Text == "") continue;

                    if (par_model.Text != par_syllable.Text)
                    {
                        //Записываем номер параграфа, в котором ошибка
                        indsTitle.Add(j);
                    }
                    break;
                }
            }

            return indsTitle;
        }

        //Проверка рабочей программы, по разделам
        public List<int> CheckSyllableSections(List<DocSection> modelSections, List<DocSection> syllableSections)
        {
            // !!!!!
            // СЮДА ЗАПИСЫВАЕМ ТОЧНЫЕ ИНДЕКСЫ ПАРАГРАФОВ ИЗ Syllable, В КОТОРЫХ ОШИБКА
            // !!!!!
            List<int> indsBody = new List<int>();
            //Section 0 = Рабочая программа рассмотрена и утверждена на заседании кафедры
            {
              /*  int CurrentIndexParagraphSyllable = 0, CurrentIndexParagraphModel = 0;
                for(int i = 0; i < syllableSections[0].Paragraphs.Count; i++)
                {
                    while(syllableSections[0].Paragraphs[i].Text != modelSections[0].Paragraphs[CurrentIndexParagraphModel].Text)
                    {
                    }
                    if(syllableSections[0].Paragraphs[i].Text != "" || (syllableSections[0].Paragraphs[i].Text == "" && modelSections[0].Paragraphs[i].Text == ""))
                    {
                    }
                    if(syllableSections[0].Paragraphs[i].Text == modelSections[0].Paragraphs[CurrentIndexParagraphModel].Text)
                    {
                        CurrentIndexParagraphModel++;
                    }
                    else
                    {
                        while (CurrentIndexParagraphModel < modelSections[0].Paragraphs.Count && syllableSections[0].Paragraphs[i].Text != modelSections[0].Paragraphs[CurrentIndexParagraphModel].Text);
                        {
                            CurrentIndexParagraphModel++;
                        }
                        if(CurrentIndexParagraphModel != modelSections[0].Paragraphs.Count)
                        {
                            CurrentIndexParagraphModel++;
                        }
                        else
                        {
                            //ошибка
                        }
                   }
                }
                */
                for (int i = 0; i <= 9; i++) //проверяем первые 10 параграфов, они должны быть идентичны
                {
                    if (i < syllableSections[0].Paragraphs.Count && i < modelSections[0].Paragraphs.Count && syllableSections[0].Paragraphs[i].Text != modelSections[0].Paragraphs[i].Text)
                    {
                        indsBody.Add(syllableSections[0].StartedAt + i);
                    }
                }

                //Обязательно должен быт указан исполнитель, т.е. параграф должен отличаться от того, что в макете
                if (10 < syllableSections[0].Paragraphs.Count && 10 < modelSections[0].Paragraphs.Count && syllableSections[0].Paragraphs[10].Text == modelSections[0].Paragraphs[10].Text)
                {
                    indsBody.Add(syllableSections[0].StartedAt + 10);
                    //ошибка
                }

                //обязательно должно быть идентичено
                if (11 < syllableSections[0].Paragraphs.Count && 11 < modelSections[0].Paragraphs.Count && syllableSections[0].Paragraphs[11].Text != modelSections[0].Paragraphs[11].Text)
                {
                    indsBody.Add(syllableSections[0].StartedAt + 11);
                }

                //12 может быть заполнен, а может быть не заполнен, поэтому не проверяем
                //с 13 по  22 должны быть идентичны
                for (int i = 13; i <= 22; i++)
                {
                    if (i < syllableSections[0].Paragraphs.Count && i < modelSections[0].Paragraphs.Count && syllableSections[0].Paragraphs[i].Text != modelSections[0].Paragraphs[i].Text)
                    {
                        indsBody.Add(syllableSections[0].StartedAt + i);
                    }
                }

                //Обязательно должен быт указан исполнитель, т.е. параграф должен отличаться от того, что в макете
                if (23 < syllableSections[0].Paragraphs.Count && 23 < modelSections[0].Paragraphs.Count && syllableSections[0].Paragraphs[23].Text == modelSections[0].Paragraphs[23].Text)
                {
                    indsBody.Add(syllableSections[0].StartedAt + 23);
                }

                //обязательно должно быть идентичено
                if (24 < syllableSections[0].Paragraphs.Count && 24 < modelSections[0].Paragraphs.Count && syllableSections[0].Paragraphs[24].Text != modelSections[0].Paragraphs[24].Text)
                {
                    indsBody.Add(syllableSections[0].StartedAt + 24);
                }
                //25 может быть заполнен, а может быть не заполнен, поэтому не проверяем
                //с 26 параграфа до конца(53) должно быть идентично
                for (int i = 26; i <= 54; i++)
                {
                    if ( i < syllableSections[0].Paragraphs.Count && i < modelSections[0].Paragraphs.Count && syllableSections[0].Paragraphs[i].Text != modelSections[0].Paragraphs[i].Text)
                    {
                        indsBody.Add(syllableSections[0].StartedAt + i);
                    }
                }
            }

            //Section 1 = 1 Цели и задачи освоения дисциплины
            {
                int ind = 0;
                int indOfTargets = -1, indOfGoals = -1;

                //Ищем, где начинаются цели
                while (!syllableSections[1].Paragraphs[ind].Text.Contains(" освоения дисциплины:") &&
                    (ind < syllableSections[1].Paragraphs.Count))
                {
                    ind++;
                }

                if (ind == syllableSections[1].Paragraphs.Count)
                {
                    //нет пункта "Цели..."
                    indsBody.Add(syllableSections[1].StartedAt);
                }
                else
                {
                    indOfTargets = syllableSections[1].StartedAt + ind;
                    ind++;
                }

                bool hasTargets = false, hasGoals = false;

                //Ищем прописанные цели; ищем, где начинаются задачи
                while ((ind < syllableSections[1].Paragraphs.Count) &&
                    !syllableSections[1].Paragraphs[ind].Text.Contains("Задачи:"))
                {
                    if (!hasTargets && (syllableSections[1].Paragraphs[ind].Text != "")) hasTargets = true;
                    ind++;
                }

                if (ind == syllableSections[1].Paragraphs.Count)
                {
                    //нет пункта "Задачи"
                    indsBody.Add(syllableSections[1].StartedAt);
                }
                else
                {
                    indOfGoals = syllableSections[1].StartedAt + ind;
                    ind++;
                }

                //Ищем прописанные задачи
                while (ind < syllableSections[1].Paragraphs.Count)
                {
                    if (!hasGoals && (syllableSections[1].Paragraphs[ind].Text != ""))
                    {
                        hasGoals = true;
                        break;
                    }
                    ind++;
                }

                if (!hasTargets && (indOfTargets != -1)) indsBody.Add(indOfTargets);
                if (!hasGoals && (indOfGoals != -1)) indsBody.Add(indOfGoals);
            }

            //Section 2 = 2 Место дисциплины в структуре образовательной программы
            {
                //Если не хватает абзацев, подсвечиваем заголовок
                if (syllableSections[2].Paragraphs.Count < 6)
                    indsBody.Add(syllableSections[2].StartedAt);

                for (int i = 0; i < 6; i++)
                {
                    //Сравниваем абзацы, которые должны совпадать с макетом
                    if (syllableSections[2].Paragraphs[i].Text != modelSections[2].Paragraphs[i].Text)
                        indsBody.Add(syllableSections[2].StartedAt + i);
                }
            }

            //Section 3 = 3 Требования к результатам обучения по дисциплине

            {
                int j = 0;
                for (int i = 0; i < syllableSections[3].Paragraphs.Count - 1; i++)
                {
                    int pId = 0;
                    if (i > 0 && modelSections[3].Paragraphs[i - 1].FollowingTables != null)
                    {
                        pId = i - 1;
                        if (modelSections[3].Paragraphs[pId].FollowingTables[0].ColumnCount == 2)//If the table has 2 columns
                        {
                            for (int y = 0; y < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows.Count; y++)//бегаем по строкам
                            {
                                for (int l = 0; l < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells.Count; l++)//бегаем по ячейкам в строке
                                {
                                    if (y == 0)
                                    {
                                        for (int g = 0; g < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                        {
                                            if (modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text != syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text)
                                            { indsBody.Add(syllableSections[3].StartedAt + i); }
                                            i++;
                                        }
                                    }
                                    else if (syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Width > 300)
                                    {
                                        int n = 0;
                                        for (int g = 0; g < syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                        {
                                            if (syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text.Contains("Знать:")
                                                || syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text.Contains("Уметь:")
                                                || syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text.Contains("Владеть:"))
                                            { n++; }
                                            i++;
                                        }
                                        if (n != 3)
                                        {
                                            for (int g = 0; g < syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                            {
                                                { indsBody.Add(syllableSections[3].StartedAt + i - syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count + g); }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        int maxSylMod = Math.Max(modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count, syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count);
                                        for (int m = 0; m < maxSylMod; m++)
                                        {
                                            for (int g = 0; g < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                            {
                                                if (modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text != syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text)
                                                { i++; indsBody.Add(syllableSections[3].StartedAt + i); m++; }
                                                while (m < maxSylMod - 1)
                                                {
                                                    i++;
                                                    indsBody.Add(syllableSections[3].StartedAt + i);
                                                    m++;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else if (modelSections[3].Paragraphs[pId].FollowingTables[0].ColumnCount == 3)//If the table has 3 columns
                        {
                            for (int y = 0; y < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows.Count; y++)//бегаем по строкам
                            {
                                for (int l = 0; l < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells.Count; l++)//бегаем по ячейкам в строке
                                {
                                    if (y == 0)
                                    {
                                        for (int g = 0; g < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                        {
                                            if (modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text != syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text)
                                            { indsBody.Add(syllableSections[3].StartedAt + i); }
                                            i++;
                                        }
                                    }
                                    else if (syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Width < 150)
                                    {
                                        int n = 0;
                                        for (int g = 0; g < syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                        {
                                            if (syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text.Contains("Знать:")
                                                || syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text.Contains("Уметь:")
                                                || syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text.Contains("Владеть:"))
                                            { n++; }
                                            i++;
                                        }
                                        if (n != 3)
                                        {
                                            for (int g = 0; g < syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                            {
                                                indsBody.Add(syllableSections[3].StartedAt + i - syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count + g);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count > syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count)
                                        {
                                            for (int g = 0; g < syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                            {
                                                
                                                indsBody.Add(syllableSections[3].StartedAt + i);
                                                i++;
                                            }
                                        }
                                        else if (modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count < syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count)
                                        {
                                            for (int g = 0; g < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                            {
                                                if (modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text != syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text)
                                                {  indsBody.Add(syllableSections[3].StartedAt + i); i++; }
                                            }
                                            int m = 0;
                                            while (m < syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count - modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count)
                                            {  indsBody.Add(syllableSections[3].StartedAt + i); m++; i++; }
                                        }
                                        else
                                        {
                                            for (int g = 0; g < modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs.Count; g++)
                                            {
                                                if (modelSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text != syllableSections[3].Paragraphs[pId].FollowingTables[0].Rows[y].Cells[l].Paragraphs[g].Text)
                                                {  indsBody.Add(syllableSections[3].StartedAt + i); i++; }
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
                            if (syllableSections[3].Paragraphs[i + 1].FollowingTables != null)
                            {
                                b = true;
                            }
                            else
                            {
                                if (syllableSections[3].Paragraphs[i + 1].Text != modelSections[3].Paragraphs[i + 1].Text)
                                {
                                    indsBody.Add(syllableSections[3].StartedAt + i + 1);
                                }
                            }
                            i++;
                        }
                    }
                    j++;
                }
            }
            //Section 4 = 4 Структура и содержание дисциплины
            {
                bool hasText = false;
                for (int i = 1; i < syllableSections[4].Paragraphs.Count; i++)
                {
                    if (syllableSections[4].Paragraphs[i].Text != "")
                    {
                        hasText = true;
                        break;
                    }
                }

                //Т.к. в данном разделе не должно быть ничего написано,
                // подсвечиваем заголовок ошибкой, если что-то есть
                if (hasText) indsBody.Add(syllableSections[4].StartedAt);
            }

            //Section 5 = 4.1 Структура дисциплины


            //Section 6 = 4.2 Содержание разделов дисциплины
            {
                bool hasText = false;
                int counter = 0;
                for (int i = 1; i < syllableSections[6].Paragraphs.Count; i++)
                {
                    if (syllableSections[6].Paragraphs[i].Text != "")
                    {
                        hasText = true;

                        //Проверяем нумерацию разделов
                        //Если пропущен раздел -- подсвечиваем вокруг этого сцуко пустого места
                        if (int.TryParse(syllableSections[6].Paragraphs[i].Text.Split(' ')[0], out int x))
                        {
                            if (x != counter + 1)
                            {
                                indsBody.Add(syllableSections[6].StartedAt + i - 2);
                                indsBody.Add(syllableSections[6].StartedAt + i - 1);
                                indsBody.Add(syllableSections[6].StartedAt + i);
                            }
                            counter = x;
                        }
                    }
                }

                //Т.к. в данном разделе должно быть написано хоть что-то,
                // подсвечиваем заголовок ошибкой, если пусто
                // (не можем проверить конкретное содержимое, too hard)
                if (!hasText) indsBody.Add(syllableSections[6].StartedAt);
            }

            //Section 7 = 4.3 Практические занятия(семинары)
            try
            {
                int spaces_in_begin = 1, spaces_in_end = 0, spaces_in_end_model = 0;
                //Считаем, есть ли пустые строки перед таблицей, чтобы если что их пропустить
                while (!syllableSections[7].Paragraphs[spaces_in_begin].Text.Contains("№ занятия"))
                {
                    spaces_in_begin++;
                    if (spaces_in_begin >= syllableSections[7].Paragraphs.Count)
                        throw new Exception();
                }

                //Считаем, есть ли пустые строки после таблицы, чтобы если что их пропустить
                int ind = syllableSections[7].Paragraphs.Count - 1;
                while (syllableSections[7].Paragraphs[ind].Text == "")
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

                //Заголовки в таблице должны совпадать с моделью
                for (int i = spaces_in_begin; i < spaces_in_begin + 4; i++)
                {
                    if (syllableSections[7].Paragraphs[i].Text != modelSections[7].Paragraphs[i].Text)
                        indsBody.Add(syllableSections[7].StartedAt + i);
                }

                //Проверяем строчки в таблице
                int ind_lesson = 0, sum = 0;
                for (int i = spaces_in_begin + 4; i < syllableSections[7].Paragraphs.Count - 4 - spaces_in_end; i += 4)
                {
                    bool isCorrect = true;

                    //Проверяем номер занятия
                    if (int.TryParse(syllableSections[7].Paragraphs[i].Text, out int x))
                    {
                        if (x != ind_lesson + 1) isCorrect = false;
                        ind_lesson = x;
                    }
                    else isCorrect = false;

                    //Проверяем номер раздела
                    if (!int.TryParse(syllableSections[7].Paragraphs[i + 1].Text, out int y))
                        isCorrect = false;

                    //Проверяем тему
                    if (syllableSections[7].Paragraphs[i + 2].Text == "") isCorrect = false;

                    //Проверяем количество часов
                    if (int.TryParse(syllableSections[7].Paragraphs[i + 3].Text, out int s))
                    {
                        sum += s;
                    }
                    else isCorrect = false;

                    //Если где-то в строке ошибка -- подсвечиваем всю строку
                    if (!isCorrect)
                    {
                        indsBody.Add(syllableSections[7].StartedAt + i);
                        indsBody.Add(syllableSections[7].StartedAt + i + 1);
                        indsBody.Add(syllableSections[7].StartedAt + i + 2);
                        indsBody.Add(syllableSections[7].StartedAt + i + 3);
                    }
                }

                if (int.TryParse(syllableSections[7].Paragraphs[syllableSections[7].Paragraphs.Count - 
                    spaces_in_end - 1].Text, out int sum_syl))
                {
                    int sum_model = int.Parse(modelSections[7].Paragraphs[modelSections[7].Paragraphs.Count -
                        spaces_in_end_model - 1].Text);

                    //Если сумма в РП не совпадает с суммой в макете
                    // или если итого по таблице не сходится = ошибко
                    if ((sum_syl != sum_model) || (sum_syl != sum))
                    {
                        indsBody.Add(syllableSections[7].StartedAt + syllableSections[7].Paragraphs.Count - spaces_in_end - 1);
                    }
                }
                else indsBody.Add(syllableSections[7].StartedAt + syllableSections[7].Paragraphs.Count - spaces_in_end - 1);
            }
            catch
            {
                indsBody.Add(syllableSections[7].StartedAt);
            }

            //Section 8 = 5 Учебно - методическое обеспечение дисциплины
            {
                bool hasText = false;
                for (int i = 1; i < syllableSections[8].Paragraphs.Count; i++)
                {
                    if (syllableSections[8].Paragraphs[i].Text != "")
                    {
                        hasText = true;
                        break;
                    }
                }

                //Т.к. в данном разделе не должно быть ничего написано,
                // подсвечиваем заголовок ошибкой, если что-то есть
                if (hasText) indsBody.Add(syllableSections[8].StartedAt);
            }

            //Section 9 = 5.1 Основная литература
            {
                bool hasText = false;
                for (int i = 1; i < syllableSections[9].Paragraphs.Count; i++)
                {
                    if (syllableSections[9].Paragraphs[i].Text != "")
                    {
                        hasText = true;
                        break;
                    }
                }

                //Т.к. в данном разделе должно быть написано хоть что-то,
                // подсвечиваем заголовок ошибкой, если пусто
                // (не можем проверить конкретное содержимое, too hard)
                if (!hasText) indsBody.Add(syllableSections[9].StartedAt);
            }

            //Section 10 = 5.2 Дополнительная литература
            {
                bool hasText = false;
                for (int i = 1; i < syllableSections[10].Paragraphs.Count; i++)
                {
                    if (syllableSections[10].Paragraphs[i].Text != "")
                    {
                        hasText = true;
                        break;
                    }
                }

                //Т.к. в данном разделе должно быть написано хоть что-то,
                // подсвечиваем заголовок ошибкой, если пусто
                // (не можем проверить конкретное содержимое, too hard)
                if (!hasText) indsBody.Add(syllableSections[10].StartedAt);
            }

            //Section 11 = 5.3 Периодические издания
            {
                bool hasText = false;
                for (int i = 1; i < syllableSections[11].Paragraphs.Count; i++)
                {
                    if (syllableSections[11].Paragraphs[i].Text != "")
                    {
                        hasText = true;
                        break;
                    }
                }

                //Т.к. в данном разделе должно быть написано хоть что-то,
                // подсвечиваем заголовок ошибкой, если пусто
                // (не можем проверить конкретное содержимое, too hard)
                if (!hasText) indsBody.Add(syllableSections[11].StartedAt);
            }

            //Section 12 = 5.4 Интернет - ресурсы
            {
                bool hasText = false;
                for (int i = 1; i < syllableSections[12].Paragraphs.Count; i++)
                {
                    if (syllableSections[12].Paragraphs[i].Text != "")
                    {
                        hasText = true;
                        break;
                    }
                }

                //Т.к. в данном разделе должно быть написано хоть что-то,
                // подсвечиваем заголовок ошибкой, если пусто
                // (не можем проверить конкретное содержимое, too hard)
                if (!hasText) indsBody.Add(syllableSections[12].StartedAt);
            }

            //Section 13 = 5.5 Программное обеспечение, профессиональные базы данных и информационные справочные системы
            {
                bool hasText = false;
                for (int i = 1; i < syllableSections[13].Paragraphs.Count; i++)
                {
                    if (syllableSections[13].Paragraphs[i].Text != "")
                    {
                        hasText = true;
                        break;
                    }
                }

                //Т.к. в данном разделе должно быть написано хоть что-то,
                // подсвечиваем заголовок ошибкой, если пусто
                // (не можем проверить конкретное содержимое, too hard)
                if (!hasText) indsBody.Add(syllableSections[13].StartedAt);
            }

            //Section 14 = 6 Материально - техническое обеспечение дисциплины
            {
                //Если не хватает абзацев, подсвечиваем заголовок
                if (syllableSections[14].Paragraphs.Count < 4)
                    indsBody.Add(syllableSections[14].StartedAt);

                //Сравниваем абзацы, которые должны совпадать с макетом
                for (int i = 0; i < 4; i++)
                {
                    if (syllableSections[14].Paragraphs[i].Text != modelSections[14].Paragraphs[i].Text)
                        indsBody.Add(syllableSections[14].StartedAt + i);
                }
            }

            return indsBody;
        }

        //Разбиение документа на секции по наименованиям разделов
        //Список наименований лежит в ресурсах
        //Результат -- список секций
        //Каждая секция -- заголовок секции + всё, что после него и до следующего заголовка
        public List<DocSection> GetDocSections(Section doc)
        {
            //Получаем имена разделов (хранятся в ресурсах)
            string names = Properties.Resources.NamesOfSections;
            List<string> namesOfSections = names.Split(new string[] { "\r\n" },
                StringSplitOptions.RemoveEmptyEntries).ToList<string>();
            int ind = 0;

            //Разбиваем док на секции
            List<DocSection> docSections = new List<DocSection>();
            for (int i = 0; i < doc.SectionParagraphs.Count; i++)
            {
                while (doc.SectionParagraphs[i].Text != namesOfSections[ind])
                {
                    i++;
                    if (i >= doc.SectionParagraphs.Count)
                    {
                        throw new Exception("Критическая ошибка: отсутствует раздел рабочей программы");
                    }
                }
                ind++;

                //Проверяем, достигли ли последнего раздела
                bool isDocEnding = false;
                if (ind >= namesOfSections.Count) isDocEnding = true;

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
                    while ((i < doc.SectionParagraphs.Count) && 
                        (doc.SectionParagraphs[i].Text != namesOfSections[ind]))
                    {
                        paragraphs.Add(doc.SectionParagraphs[i]);
                        i++;
                    }
                }
                i--;
                docSections.Add(new DocSection(startedAt, i, paragraphs));
            }

            return docSections;
        }

        public bool checkParagraphEquality()
        {
            DocX model = this.Model,
                syllabus = this.Syllable;


            bool firstCase = this.compareTwoParagraphs(model.Sections[1].SectionParagraphs[15], syllabus.Sections[1].SectionParagraphs[15]);
            return true;
        }

        //Проверка на соответствие двух абзацев
        private bool compareTwoParagraphs(Paragraph p1, Paragraph p2)
        {
            List<Tuple<string, bool>> p1Array = new List<Tuple<string, bool>>();
            string currentTextPart;
            bool currentHightlight;

            //разбираем абзац из шаблона по частям, отличающимся подсветкой
            for(int i = 0; i < p1.MagicText.Count; i++)
            {
                currentHightlight = p1.MagicText[i].formatting.Highlight == Highlight.green;
                currentTextPart = p1.MagicText[i].text;

                if (i == 0 || p1Array.Last().Item2 != currentHightlight)
                {
                    p1Array.Add(new Tuple<string, bool>(currentTextPart, currentHightlight));
                }
                else
                {
                    p1Array[p1Array.Count - 1] = new Tuple<string, bool>(
                        p1Array[p1Array.Count - 1].Item1 + currentTextPart,
                        currentHightlight
                    );
                }
            }

            string p2Text = p2.Text;

            for(int i = 0; i < p1Array.Count; i++)
            {
                if(p1Array[i].Item2) //зеленый
                {
                    if (i == p1Array.Count - 1)
                    {
                        return true;
                    }
                    else
                    {
                        continue;
                    }
                }
                else
                {
                    int index = p2Text.IndexOf(p1Array[i].Item1);

                    if(index == -1)
                    {
                        MessageBox.Show("Пропущен обязательный фрагмент: '" + p1Array[i].Item1 + "'");
                        return false;
                    }
                    else
                    {
                        if (i != 0 && p1Array[i - 1].Item2) //до этого был зелёный фрагмент
                        {
                            p2Text = p2Text.Substring(index + p1Array[i].Item1.Length);
                        }
                        else
                        {
                            if (index == 0)
                            {
                                p2Text = p2Text.Substring(p1Array[i].Item1.Length);
                            }
                            MessageBox.Show("Лишний текст: '" + p2Text.Substring(0, index) + "'");
                            return false;
                        }
                    }
                }
            }

            return false;
            //+разбить первый параграф на массив последовательностей [text, type].
            //+пробежать по массиву и
            //+  если незелёный, поискать indexOf.
            //+    Если не найден, то возврат с сообщением: пропущен обязательный фрагмент ...
            //+    Если найден и до этого обрабатывался зелёный фрагмент,
            //+      то всё перед ним отнести в зелёное. А этот фрагмент и всё до него вырезать из строки
            //+    Если найден и до этого ничего не было,
            //+      то возврат с сообщением про лишний текст (выдать его)
            //+      Если лишнего нет, то удаляем этот фрагмент
            //+  если зелёный, то
            //+    если последний фрагмент, то возвратить успех
            //+    если дальше есть, то континью

        }
    }
}
