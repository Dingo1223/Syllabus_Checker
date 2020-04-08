using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace SyllabusChecker
{
    public class HighlightHandler
    {
        private DocX Model;
        private DocX Syllabus;

        /// <summary>
        /// Проверяет два документа на соответствие (подсветка)
        /// </summary>
        /// <param name="inputData">Выбранные пользователем расположения файлов</param>
        /// <returns>Количество найденных ошибок</returns>
        public int CheckDocumentsEquality(InputData inputData)
        {
            Model = DocX.Load(inputData.ModelPath);
            Syllabus = DocX.Load(inputData.DocumentPath);
            List<Error> errors = CheckDocumentsTextEquality();
            CreateResultDoc(errors, inputData);

            return errors.Count;
        }

        /// <summary>
        /// Создание результирующего документа, с отмеченными ошибками
        /// </summary>
        /// <param name="errors">Список ошибок</param>
        /// <param name="inputData">Выбранные пользователем расположения файлов</param>
        public void CreateResultDoc(List<Error> errors, InputData inputData)
        {
            string path = Path.Combine(inputData.ResultFolderPath,
                Path.GetFileNameWithoutExtension(inputData.DocumentPath) + "_checked.docx");

            //Сохраняем проверенный документ
            Syllabus.SaveAs(path);

            Dictionary<int, string> errors_dict = new Dictionary<int, string>();
            foreach (Error err in errors)
            {
                if (!errors_dict.ContainsKey(err.Index))
                {
                    errors_dict.Add(err.Index, err.Type);
                }
            }

            //Добавляем комментарии об ошибках
            DocComments.AddComments(errors_dict, path);
        }

        /// <summary>
        /// Проверяет документы на сходство (?)
        /// </summary>
        /// <returns>Список найденных ошибок с описанием</returns>
        public List<Error> CheckDocumentsTextEquality()
        {
            List<Error> errors = new List<Error>();
            List<ModelParagraph> usefulModelParagraphs = new List<ModelParagraph>();
            //Paragraph currentModelParagraph;
            //string currentParagraph;
            //int currentModelIndex;

            for (int i = 0; i < Model.Paragraphs.Count; i++)
            {
                //currentModelParagraph = model.Paragraphs[i];
                //if (HasSomeContent(Model.Paragraphs[i].Text) && HasWhiteSegments(i))
                if ((Model.Paragraphs[i].Text.Length > 0) && HasWhiteSegments(i))
                {
                    usefulModelParagraphs.Add(new ModelParagraph(i, Model.Paragraphs[i].Text));
                }
            }

            int k = 0; //usefulModelParagraphs counter
            int lastSyllabusIndex = Syllabus.Paragraphs.Count - 1;

            for (int i = 0; i < Syllabus.Paragraphs.Count; i++)
            {
                //currentParagraph = syllabus.Paragraphs[i].Text;

                //if (!HasSomeContent(Syllabus.Paragraphs[i].Text)) continue;
                if (Syllabus.Paragraphs[i].Text.Length == 0) continue;

                //currentModelParagraph = usefulModelParagraphs[k].paragraph; //paragraph
                //currentModelIndex = usefulModelParagraphs[k].sourceIndex;

                //bool paragraphsMatchUp = CompareTwoParagraphs(usefulModelParagraphs[k].paragraph, syllabus.Paragraphs[i].Text);
                if (CompareTwoParagraphs(Model.Paragraphs[usefulModelParagraphs[k].SourceIndex], Syllabus.Paragraphs[i].Text))
                {
                    k++;
                    if (k == usefulModelParagraphs.Count) //model ends
                    {
                        lastSyllabusIndex = i;
                        break;
                    }
                    continue;
                }
                else if (!StartsWithGreenHighligth(usefulModelParagraphs[k].SourceIndex))
                {
                    bool previousNonemptyIsGreen = false;
                    int counter = usefulModelParagraphs[k].SourceIndex - 1;
                    //while (counter >= 0 && !HasSomeContent(Model.Paragraphs[counter].Text))
                    while (counter >= 0 && (Model.Paragraphs[counter].Text.Length == 0))
                    {
                        counter--;
                    }
                    if (counter >= 0 && EndsWithGreenHighligth(counter))
                    {
                        previousNonemptyIsGreen = true;
                    }

                    if (usefulModelParagraphs[k].SourceIndex == 0 || !previousNonemptyIsGreen)
                    {
                        errors.Add(new Error(i, "Лишний текст в РП! Абзац: " + i.ToString()));
                    }
                }
                else continue;
            }

            //model doest'n ends
            if (k != usefulModelParagraphs.Count)
            {
                errors.Add(new Error(usefulModelParagraphs[k].SourceIndex, 
                    "В РП нет обязательного фрагмента: " + usefulModelParagraphs[k].Paragraph));
            }
            else if ((lastSyllabusIndex != Syllabus.Paragraphs.Count - 1) // syllabus doesn't ends
                && !EndsWithGreenHighligth(usefulModelParagraphs.Last().SourceIndex))
            {
                bool nextNonemptyIsGreen = false;
                int counter = usefulModelParagraphs.Last().SourceIndex + 1;
                //while (counter <= Syllabus.Paragraphs.Count && !HasSomeContent(Model.Paragraphs[counter].Text))
                while (counter <= Syllabus.Paragraphs.Count && (Model.Paragraphs[counter].Text.Length == 0))
                {
                    counter++;
                }
                if (counter <= Syllabus.Paragraphs.Count && StartsWithGreenHighligth(counter))
                {
                    nextNonemptyIsGreen = true;
                }


                bool hasNonemptyParagraph = false;
                for (int i = lastSyllabusIndex + 1; i < Syllabus.Paragraphs.Count; i++)
                {
                    //if (HasSomeContent(Syllabus.Paragraphs[i].Text))
                    if (Syllabus.Paragraphs[i].Text.Length > 0)
                    {
                        hasNonemptyParagraph = true;
                        break;
                    }
                }

                if (hasNonemptyParagraph && !nextNonemptyIsGreen)
                {
                    errors.Add(new Error(lastSyllabusIndex, "Лишний текст в документе: " + Syllabus.Paragraphs[lastSyllabusIndex].Text));
                }
            }

            return errors;
        }

        /// <summary>
        /// Проверка, есть ли белые части в параграфе
        /// </summary>
        /// <param name="p">Параграф для проверки</param>
        /// <returns></returns>
        private bool HasWhiteSegments(int p)
        {
            if (Model.Paragraphs[p].MagicText.Count == 0) return false;
            for (int i = 0; i < Model.Paragraphs[p].MagicText.Count; i++)
            {
                if (Model.Paragraphs[p].MagicText[i].formatting == null || 
                    Model.Paragraphs[p].MagicText[i].formatting.Highlight != Highlight.green)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Проверка, есть ли содеримое в параграфе
        /// </summary>
        /// <param name="text">Текст параграфа</param>
        /// <returns></returns>
        //private static bool HasSomeContent(string text) => text.Trim().Length > 0;

        private bool StartsWithGreenHighligth(int p) => 
            Model.Paragraphs[p].MagicText.Count > 0 && Model.Paragraphs[p].MagicText[0].formatting != null &&
            Model.Paragraphs[p].MagicText[0].formatting.Highlight == Highlight.green;

        private bool EndsWithGreenHighligth(int p) =>
            Model.Paragraphs[p].MagicText.Count > 0 && Model.Paragraphs[p].MagicText.Last().formatting != null &&
            Model.Paragraphs[p].MagicText.Last().formatting.Highlight == Highlight.green;

        /// <summary>
        /// Проверка на соответствие двух параграфов
        /// </summary>
        /// <param name="p1">Первый параграф</param>
        /// <param name="p2">Текст второго параграфа</param>
        /// <returns></returns>
        private bool CompareTwoParagraphs(Paragraph p1, string p2)
        {
            List<ElementaryFragment> modelFragments = new List<ElementaryFragment>();
            string currentTextPart;
            bool currentHightlight;
            ElementaryFragment currentFragment;

            //Разбираем абзац из шаблона по частям, отличающимся подсветкой
            for (int i = 0; i < p1.MagicText.Count; i++)
            {
                currentHightlight = p1.MagicText[i].formatting != null && p1.MagicText[i].formatting.Highlight == Highlight.green;
                currentTextPart = p1.MagicText[i].text;

                if (currentTextPart.Length == 0) continue;

                currentFragment.isGreen = currentHightlight;

                if (modelFragments.Count == 0 || modelFragments.Last().isGreen != currentHightlight)
                {
                    currentFragment.text = currentTextPart;
                    modelFragments.Add(currentFragment);
                }
                else
                {
                    currentFragment.text = modelFragments.Last().text + currentTextPart;
                    modelFragments[modelFragments.Count - 1] = currentFragment;
                }
            }

            for (int i = 0; i < modelFragments.Count; i++)
            {
                currentFragment.isGreen = modelFragments[i].isGreen;
                currentFragment.text = modelFragments[i].text.Trim();
                modelFragments[i] = currentFragment;
            }

            p2 = p2.Trim();

            for (int i = 0; i < modelFragments.Count; i++)
            {
                if (modelFragments[i].isGreen)
                {
                    if (i == modelFragments.Count - 1) return true;
                    else continue;
                }
                else
                {
                    int index = p2.IndexOf(modelFragments[i].text);

                    if (index == -1) return false;
                    else
                    {
                        if (i != 0 && modelFragments[i - 1].isGreen) //до этого был зелёный фрагмент
                        {
                            p2 = p2.Substring(index + modelFragments[i].text.Length);
                            if (p2.Trim().Length == 0) return true;
                        }
                        else
                        {
                            if (index == 0)
                            {
                                p2 = p2.Substring(modelFragments[i].text.Length);
                            }
                            if (p2.Length == 0) //Syllabus ends
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            return false;
        }
    }
}
