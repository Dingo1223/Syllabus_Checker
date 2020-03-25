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
            //CreateResultDoc(IndsTitle, IndsBody, inputData);
        }

        //Создание результирующего документа (с подсветкой ошибочных мест)
        //Строится на основе Syllable
        //Параметр indsTitle -- индексы параграфов с ошибками в титульнике
        //Параметр indsBody -- индексы параграфов с ошибками в остальном документе
        public void CreateResultDoc(List<int> indsTitle, List<int> indsBody, InputData inputData)
        {
            string path = Path.Combine(inputData.ResultFolderPath,
                Path.GetFileNameWithoutExtension(inputData.SyllablePath) + "_checked.docx");
            DocX document = DocX.Create(path);

            //Переносим титульник
            document.InsertSection();
            for (int i = 0; i < Syllable.Sections[0].SectionParagraphs.Count; i++)
            {
                Paragraph p = Syllable.Sections[0].SectionParagraphs[i];
                if (indsTitle.Contains(i))
                {
                    p.Highlight(Highlight.red);
                }
                document.Sections[0].InsertParagraph(p);
            }

            //Переносим остальное
            document.InsertSection();
            for (int i = 0; i < Syllable.Sections[1].SectionParagraphs.Count; i++)
            {
                Paragraph p = Syllable.Sections[1].SectionParagraphs[i];
                if (indsBody.Contains(i))
                {
                    p.Highlight(Highlight.red);
                }
                document.Sections[1].InsertParagraph(p);
            }

            document.Save();
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
                }

                if (!hasTargets && (indOfTargets != -1)) indsBody.Add(indOfTargets);
                if (!hasGoals && (indOfGoals != -1)) indsBody.Add(indOfGoals);
            }

            //Section 2 = 2 Место дисциплины в структуре образовательной программы
            {
                int j = 1;
                for (int i = 1; i < modelSections[2].Paragraphs.Count; i++)
                {
                    if (modelSections[2].Paragraphs[i].Text == "") continue;

                    //Ищем такой же параграф в syllable
                    while (modelSections[2].Paragraphs[i].Text != syllableSections[2].Paragraphs[j].Text)
                    {
                        j++;
                        if (j >= syllableSections[2].Paragraphs.Count)
                        {
                            //Если какого-то из нужных параграфов в этом разделе нет --
                            // подсвечивается заголовок раздела
                            if (!indsBody.Contains(syllableSections[2].StartedAt))
                                indsBody.Add(syllableSections[2].StartedAt);
                            break;
                        }
                    }
                    
                }
            }

            //Section 3 = 3 Требования к результатам обучения по дисциплине


            //Section 4 = 4 Структура и содержание дисциплины


            //Section 5 = 4.1 Структура дисциплины


            //Section 6 = 4.2 Содержание разделов дисциплины


            //Section 7 = 4.3 Практические занятия(семинары)


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
                while (doc.SectionParagraphs[i].Text != namesOfSections[ind]) i++;
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
