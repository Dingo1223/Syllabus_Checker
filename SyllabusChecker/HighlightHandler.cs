using System.Collections.Generic;
using System.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;
using MessageBox = System.Windows.MessageBox;

namespace SyllabusChecker
{
    struct ElementaryFragment
    {
        public string text;
        public bool isGreen;
    }

    struct ModelParagraph
    {
        public int sourceIndex; //index in original
        public Paragraph paragraph;
        //public string text; //full text
        //public ElementaryFragment[] fragments; //magic text
    };

    struct Document
    {
        public ModelParagraph[] paragraphs;
    };

    struct Error
    {
        public string paragraph;
        public int index;
        public string type;
    }

    static class HighlightHandler
    {
        /// <summary>
        /// Проверяет два документа на соответствие (подсветка)
        /// </summary>
        /// <param name="inputData">Выбранные пользователем расположения файлов</param>
        /// <returns></returns>
        public static bool CheckDocumentsEquality(InputData inputData)
        {
            DocX Model = DocX.Load(inputData.ModelPath);
            DocX Syllable = DocX.Load(inputData.SyllablePath);
            return HighlightHandler.CheckDocumentsTextEquality(Model, Syllable);
        }

        /// <summary>
        /// Проверяет документы на сходство (?)
        /// </summary>
        /// <param name="model">Образец документа</param>
        /// <param name="syllabus">Проверяемый документ</param>
        /// <returns></returns>
        public static bool CheckDocumentsTextEquality(DocX model, DocX syllabus)
        {
            List<Error> errors = new List<Error>();
            List<ModelParagraph> usefulModelParagraphs = new List<ModelParagraph>();
            Paragraph currentModelParagraph;
            string currentParagraph;
            int currentModelIndex;

            for (int i = 0; i < model.Paragraphs.Count; i++)
            {
                currentModelParagraph = model.Paragraphs[i];
                if (HasSomeContent(currentModelParagraph.Text) && HasWhiteSegments(currentModelParagraph))
                {
                    ModelParagraph mp = new ModelParagraph
                    {
                        paragraph = currentModelParagraph,
                        sourceIndex = i
                    };
                    usefulModelParagraphs.Add(mp);
                }
            }

            int k = 0; //usefulModelParagraphs counter
            int lastSyllabusIndex = syllabus.Paragraphs.Count - 1;

            for (int i = 0; i < syllabus.Paragraphs.Count; i++)
            {
                currentParagraph = syllabus.Paragraphs[i].Text;

                if (!HasSomeContent(currentParagraph)) continue;

                currentModelParagraph = usefulModelParagraphs[k].paragraph; //paragraph
                currentModelIndex = usefulModelParagraphs[k].sourceIndex;

                bool paragraphsMatchUp = CompareTwoParagraphs(currentModelParagraph, currentParagraph);
                if (paragraphsMatchUp)
                {
                    k++;
                    if (k == usefulModelParagraphs.Count) //model ends
                    {
                        lastSyllabusIndex = i;
                        break;
                    }
                    continue;
                }
                else if (!StartsWithGreenHighligth(currentModelParagraph))
                {
                    bool previousNonemptyIsGreen = false;
                    int counter = currentModelIndex - 1;
                    while (counter >= 0 && !HasSomeContent(model.Paragraphs[counter].Text))
                    {
                        counter--;
                    }
                    if (counter >= 0 && EndsWithGreenHighligth(model.Paragraphs[counter]))
                    {
                        previousNonemptyIsGreen = true;
                    }

                    if (currentModelIndex == 0 || !previousNonemptyIsGreen)
                    {
                        MessageBox.Show("Лишний текст в РП! Абзац: " + i.ToString());
                        Error error;
                        error.paragraph = currentParagraph;
                        error.index = i;
                        error.type = "Undeclared text in Syllabus";
                        errors.Add(error);
                    }
                }
                else continue;
            }

            //model doest'n ends
            if (k != usefulModelParagraphs.Count)
            {
                MessageBox.Show("В РП нет обязательного фрагмента! Индекс фрагмента в модели: " + k.ToString() + " -- " + usefulModelParagraphs[k].paragraph.Text);
                Error error;
                error.paragraph = usefulModelParagraphs[k].paragraph.Text;
                error.index = usefulModelParagraphs[k].sourceIndex;
                error.type = "Missing required paragraph";
                errors.Add(error);
            }
            else if (lastSyllabusIndex != syllabus.Paragraphs.Count - 1 // syllabus doesn't ends
                && !EndsWithGreenHighligth(usefulModelParagraphs.Last().paragraph))
            {
                bool nextNonemptyIsGreen = false;
                int counter = usefulModelParagraphs.Last().sourceIndex + 1;
                while (counter <= syllabus.Paragraphs.Count && !HasSomeContent(model.Paragraphs[counter].Text))
                {
                    counter++;
                }
                if (counter <= syllabus.Paragraphs.Count && StartsWithGreenHighligth(model.Paragraphs[counter]))
                {
                    nextNonemptyIsGreen = true;
                }


                bool hasNonemptyParagraph = false;
                for (int i = lastSyllabusIndex + 1; i < syllabus.Paragraphs.Count; i++)
                {
                    if (HasSomeContent(syllabus.Paragraphs[i].Text))
                    {
                        hasNonemptyParagraph = true;
                        break;
                    }
                }

                if (hasNonemptyParagraph && !nextNonemptyIsGreen)
                {
                    MessageBox.Show("Лишний текст в документе! Абзац: " + lastSyllabusIndex.ToString() + " -- " + syllabus.Paragraphs[lastSyllabusIndex].Text);
                    Error error;
                    error.paragraph = syllabus.Paragraphs[lastSyllabusIndex].Text;
                    error.index = lastSyllabusIndex;
                    error.type = "Undeclared text in Syllabus";
                    errors.Add(error);
                }
            }

            //пока так выводятся ошибки, потом можно как-то получше выводить массив errors

            if (errors.Count == 0) return true;
            else return false;
        }

        /// <summary>
        /// Проверка, есть ли белые части в параграфе
        /// </summary>
        /// <param name="p">Параграф для проверки</param>
        /// <returns></returns>
        private static bool HasWhiteSegments(Paragraph p)
        {
            if (p.MagicText.Count == 0) return false;
            for (int i = 0; i < p.MagicText.Count; i++)
            {
                if (p.MagicText[i].formatting == null || p.MagicText[i].formatting.Highlight != Highlight.green)
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

        private static bool StartsWithGreenHighligth(Paragraph p) => 
            p.MagicText.Count > 0 && p.MagicText[0].formatting != null && p.MagicText[0].formatting.Highlight == Highlight.green;

        private static bool EndsWithGreenHighligth(Paragraph p) => 
            p.MagicText.Count > 0 && p.MagicText.Last().formatting != null && p.MagicText.Last().formatting.Highlight == Highlight.green;

        /// <summary>
        /// Проверка на соответствие двух параграфов
        /// </summary>
        /// <param name="p1">Первый параграф</param>
        /// <param name="p2">Текст второго параграфа</param>
        /// <returns></returns>
        private static bool CompareTwoParagraphs(Paragraph p1, string p2)
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

                    if (index == -1)
                    {
                        //MessageBox.Show("Пропущен обязательный фрагмент: '" + modelFragments[i].text + "'");
                        return false;
                    }
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
                            //MessageBox.Show("Лишний текст: '" + p2.Substring(0, index) + "'");
                            //return false;
                        }
                    }
                }
            }
            return false;
        }
    }
}
