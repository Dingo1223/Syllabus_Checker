using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace SyllabusChecker
{
    public class HighlightHandler
    {
        //private DocX Model;
        private List<Paragraph> ModParagraphs;
        private DocX Syllabus;
        private List<Paragraph> SylParagraphs;

        /// <summary>
        /// Проверяет два документа на соответствие (подсветка)
        /// </summary>
        /// <param name="inputData">Выбранные пользователем расположения файлов</param>
        /// <returns>Количество найденных ошибок</returns>
        public int CheckDocumentsEquality(InputData inputData)
        {
            //Model = DocX.Load(inputData.ModelPath);
            Syllabus = DocX.Load(inputData.DocumentPath);

            ModParagraphs = DocX.Load(inputData.ModelPath).Paragraphs.ToList();
            SylParagraphs = Syllabus.Paragraphs.ToList();
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
            List<int> usefulModelParagraphs = new List<int>();

            for (int i = 0; i < ModParagraphs.Count; i++)
            {
                if ((ModParagraphs[i].Text != "") && HasWhiteSegments(i))
                {
                    usefulModelParagraphs.Add(i);
                }
            }

            int k = 0; //usefulModelParagraphs counter
            int lastSyllabusIndex = SylParagraphs.Count - 1;

            for (int i = 0; i < SylParagraphs.Count; i++)
            {
                if (!HasSomeContent(SylParagraphs[i].Text)) continue;

                if (CompareTwoParagraphs(ModParagraphs[usefulModelParagraphs[k]], SylParagraphs[i].Text))
                {
                    k++;
                    if (k == usefulModelParagraphs.Count) //model ends
                    {
                        lastSyllabusIndex = i;
                        break;
                    }
                    continue;
                }
                //else if (!StartsWithGreenHighligth(usefulModelParagraphs[k]))
                //{
                //    bool previousNonemptyIsGreen = false;
                //    int counter = usefulModelParagraphs[k] - 1;
                //    while (counter >= 0 && !HasSomeContent(ModParagraphs[counter].Text))
                //    {
                //        counter--;
                //    }
                //    if (counter >= 0 && EndsWithGreenHighligth(counter))
                //    {
                //        previousNonemptyIsGreen = true;
                //    }

                //    if (usefulModelParagraphs[k] == 0 || !previousNonemptyIsGreen)
                //    {
                //        errors.Add(new Error(i, "Лишний текст в документе; абзац: \"" + i.ToString() + "\""));
                //    }
                //}
                else continue;
            }

            //model doest'n ends
            if (k != usefulModelParagraphs.Count)
            {
                errors.Add(new Error(usefulModelParagraphs[k], 
                    "В документе нет обязательного фрагмента: \"" + ModParagraphs[usefulModelParagraphs[k]].Text + "\""));
            }
            else if ((lastSyllabusIndex != SylParagraphs.Count - 1) && 
                !EndsWithGreenHighligth(usefulModelParagraphs.Last())) // syllabus doesn't ends
            {
                bool nextNonemptyIsGreen = false;
                int counter = usefulModelParagraphs.Last() + 1;
                while (counter <= SylParagraphs.Count && !HasSomeContent(SylParagraphs[counter].Text))
                {
                    counter++;
                }
                if (counter <= SylParagraphs.Count && StartsWithGreenHighligth(counter))
                {
                    nextNonemptyIsGreen = true;
                }


                bool hasNonemptyParagraph = false;
                for (int i = lastSyllabusIndex + 1; i < SylParagraphs.Count; i++)
                {
                    if (HasSomeContent(SylParagraphs[i].Text))
                    {
                        hasNonemptyParagraph = true;
                        break;
                    }
                }

                if (hasNonemptyParagraph && !nextNonemptyIsGreen)
                {
                    errors.Add(new Error(lastSyllabusIndex, "Лишний текст в документе: \"" +
                        SylParagraphs[lastSyllabusIndex].Text + "\""));
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
            if (ModParagraphs[p].MagicText.Count == 0) return false;
            for (int i = 0; i < ModParagraphs[p].MagicText.Count; i++)
            {
                if (ModParagraphs[p].MagicText[i].formatting == null || 
                    ModParagraphs[p].MagicText[i].formatting.Highlight != Highlight.green)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Проверка, есть ли содеримое в параграфе
        /// </summary>
        /// <param name="text">Текст параграфа</param>
        /// <returns></returns>
        private static bool HasSomeContent(string text) => text.Trim().Length > 0;

        private bool StartsWithGreenHighligth(int p) => 
            ModParagraphs[p].MagicText.Count > 0 && ModParagraphs[p].MagicText[0].formatting != null &&
            ModParagraphs[p].MagicText[0].formatting.Highlight == Highlight.green;

        private bool EndsWithGreenHighligth(int p) =>
            ModParagraphs[p].MagicText.Count > 0 && ModParagraphs[p].MagicText.Last().formatting != null &&
            ModParagraphs[p].MagicText.Last().formatting.Highlight == Highlight.green;

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

                currentFragment.IsGreen = currentHightlight;

                if (modelFragments.Count == 0 || modelFragments.Last().IsGreen != currentHightlight)
                {
                    currentFragment.Text = currentTextPart;
                    modelFragments.Add(currentFragment);
                }
                else
                {
                    currentFragment.Text = modelFragments.Last().Text + currentTextPart;
                    modelFragments[modelFragments.Count - 1] = currentFragment;
                }
            }

            for (int i = 0; i < modelFragments.Count; i++)
            {
                currentFragment.IsGreen = modelFragments[i].IsGreen;
                currentFragment.Text = modelFragments[i].Text.Trim();
                modelFragments[i] = currentFragment;
            }

            p2 = p2.Trim();

            for (int i = 0; i < modelFragments.Count; i++)
            {
                if (modelFragments[i].IsGreen)
                {
                    if (i == modelFragments.Count - 1) return true;
                    else continue;
                }
                else
                {
                    int index = p2.IndexOf(modelFragments[i].Text);

                    if (index == -1) return false;
                    else
                    {
                        if (i != 0 && modelFragments[i - 1].IsGreen) //до этого был зелёный фрагмент
                        {
                            p2 = p2.Substring(index + modelFragments[i].Text.Length);
                            if (p2.Trim().Length == 0) return true;
                        }
                        else
                        {
                            if (index == 0)
                            {
                                p2 = p2.Substring(modelFragments[i].Text.Length);
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
