using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using Word_kazakov.Models;
using Word = Microsoft.Office.Interop.Word;

namespace Word_kazakov.Context
{
    public class OwnerContext : Owner
    {
        public OwnerContext(string FirstName, string LastName, string SureName, int NumberRoom) : base (FirstName, LastName, SureName, NumberRoom) { }

        public static List<OwnerContext> AllOwners()
        {
            List<OwnerContext> allOwners = new List<OwnerContext>();

            allOwners.Add(new OwnerContext("Елена", "Иванова", "Петровна", 1));
            allOwners.Add(new OwnerContext("Алексей", "Смирнов", "Владимирович", 2));
            allOwners.Add(new OwnerContext("Анна", "Кузнецова", "Сергеевна", 3));
            allOwners.Add(new OwnerContext("Дмитрий", "Павлов", "Александрович", 3));
            allOwners.Add(new OwnerContext("Ольга", "Михайлова", "Ивановна", 4));
            allOwners.Add(new OwnerContext("Артем", "Козлов", "Олегович", 5));
            allOwners.Add(new OwnerContext("Наталья", "Лебедева", "Андреевна", 6));
            allOwners.Add(new OwnerContext("Игорь", "Федоров", "Дмитриевич", 7));
            allOwners.Add(new OwnerContext("Екатерина", "Александрова", "Игоревна", 7));
            allOwners.Add(new OwnerContext("Андрей", "Степанов", "Николаевич", 8));
            allOwners.Add(new OwnerContext("Оксана", "Никитина", "Васильевна", 9));
            allOwners.Add(new OwnerContext("Сергей", "Ковалев", "Александрович", 10));
            allOwners.Add(new OwnerContext("Мария", "Фролова", "Михайловна", 11));
            allOwners.Add(new OwnerContext("Павел", "Белов", "Александрович", 12));
            allOwners.Add(new OwnerContext("Елена", "Полякова", "Даниловна", 13));
            allOwners.Add(new OwnerContext("Илья", "Гаврилов", "Валерьевич", 14));
            allOwners.Add(new OwnerContext("Анастасия", "Орлова", "Владимировна", 15));
            allOwners.Add(new OwnerContext("Денис", "Киселев", "Сергеевич", 16));
            allOwners.Add(new OwnerContext("Алина", "Ткаченко", "Викторовна", 16));
            allOwners.Add(new OwnerContext("Артем", "Романов", "Павлович", 16));
            allOwners.Add(new OwnerContext("Валерия", "Максимова", "Юрьевна", 17));
            allOwners.Add(new OwnerContext("Александр", "Сидоров", "Игоревич", 17));
            allOwners.Add(new OwnerContext("Евгения", "Антонова", "Алексеевна", 18));
            allOwners.Add(new OwnerContext("Никита", "Дмитриев", "Владимирович", 19));

            return allOwners;
        }

        public static void Report(string fileName)
        {
            Word.Application app = new Word.Application();
            Word.Document doc = app.Documents.Add();

            Word.Paragraph paraHeader = doc.Paragraphs.Add();
            paraHeader.Range.Font.Size = 16;
            paraHeader.Range.Font.Bold = 1;
            paraHeader.Range.Text = "Список жильцов дома";
            paraHeader.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paraHeader.Range.InsertParagraphAfter();

            Word.Paragraph paraAddress = doc.Paragraphs.Add();
            paraAddress.Range.Font.Size = 14;
            paraAddress.Range.Font.Bold = 0;
            paraAddress.Range.Text = "по адресу: Пермь, ул. Луначарского, д. 24";
            paraAddress.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paraAddress.Range.ParagraphFormat.SpaceAfter = 20;
            paraAddress.Range.InsertParagraphAfter();

            var ownersList = AllOwners();
            Word.Paragraph paraCount = doc.Paragraphs.Add();
            paraCount.Range.Font.Size = 14;
            paraCount.Range.Text = $"Всего жильцов: {ownersList.Count}";
            paraCount.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paraCount.Range.InsertParagraphAfter();

            Word.Paragraph tablePara = doc.Paragraphs.Add();
            Word.Table paymentsTable = doc.Tables.Add(tablePara.Range, ownersList.Count + 1, 4);
            paymentsTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            paymentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            FillCell(paymentsTable.Cell(1, 1), "№", true);
            FillCell(paymentsTable.Cell(1, 2), "Фамилия", true);
            FillCell(paymentsTable.Cell(1, 3), "Имя", true);
            FillCell(paymentsTable.Cell(1, 4), "Отчество", true);

            for (int i = 0; i < ownersList.Count; i++)
            {
                OwnerContext owner = ownersList[i];
                int row = i + 2;

                FillCell(paymentsTable.Cell(row, 1), (i + 1).ToString());
                FillCell(paymentsTable.Cell(row, 2), owner.LastName);
                FillCell(paymentsTable.Cell(row, 3), owner.FirstName);
                FillCell(paymentsTable.Cell(row, 4), owner.SureName);
            }

            doc.SaveAs2(fileName);
            doc.Close();
            app.Quit();
        }

        private static void FillCell(Word.Cell cell, string text, bool isBold = false)
        {
            cell.Range.Text = text;
            cell.Range.Font.Bold = isBold ? 1 : 0;
            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
        }
    }
}