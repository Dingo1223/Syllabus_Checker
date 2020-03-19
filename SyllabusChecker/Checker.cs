using System;
using System.Collections.Generic;
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
            CheckTitlePage();

            //Получаем разбитые на секции модель и РП
            List<DocSection> ModelSections = GetDocSections(Model.Sections[1]);
            List<DocSection> SyllableSections = GetDocSections(Syllable.Sections[1]);

            //Дальше надо их проверять
        }

        //Проверка титульника
        public void CheckTitlePage()
        {
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

                    //Т.к. важен только контент, проверку стилей можно временно убрать
                    /*if ((par_model.Text != par_syllable.Text) ||
                        (par_model.Alignment != par_syllable.Alignment) ||
                        (par_model.IndentationAfter != par_syllable.IndentationAfter) ||
                        (par_model.IndentationBefore != par_syllable.IndentationBefore) ||
                        (par_model.IndentationFirstLine != par_syllable.IndentationFirstLine) ||
                        (par_model.IndentationHanging != par_syllable.IndentationHanging) ||
                        (par_model.IsKeepWithNext != par_syllable.IsKeepWithNext) ||
                        (par_model.LineSpacing != par_syllable.LineSpacing) ||
                        (par_model.LineSpacingAfter != par_syllable.LineSpacingAfter) ||
                        (par_model.LineSpacingBefore != par_syllable.LineSpacingBefore) ||
                        (par_model.StyleName != par_syllable.StyleName))*/

                    if (par_model.Text != par_syllable.Text)
                    {
                        MessageBox.Show("Несоответствие в параграфе №" + i.ToString() +
                            "\nТекст в макете: " + par_model.Text +
                            "\nТекст в проверяемой программе: " + par_syllable.Text);
                    }
                    break;
                }
            }
        }

        //Разбиение документа на секции по наименованиям разделов
        //Список наименований лежит в ресурсах
        //Результат -- список секций
        //Каждая секция -- заголовок секции + всё, что после него и до следующего заголовка
        public List<DocSection> GetDocSections(Section doc)
        {
            //Получаем имена разделов (хранятся в ресурсах)
            string names = Properties.Resources.NamesOfSections;
            List<string> namesOfSections = names.Split(new string[] {"\r\n"}, 
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
                    while (doc.SectionParagraphs[i].Text != namesOfSections[ind])
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
    }
}
