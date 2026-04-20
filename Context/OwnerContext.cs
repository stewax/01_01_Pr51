using Microsoft.Office.Interop.Word;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
        public class MyFontResolver : IFontResolver
        {
            public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
            {
                if (familyName.Equals("Arial", StringComparison.OrdinalIgnoreCase))
                {
                    if (isBold) return new FontResolverInfo("Arial#b");
                    return new FontResolverInfo("Arial#");
                }
                return null;
            }

            public byte[] GetFont(string faceName)
            {
                string fontPath = faceName.Contains("b")
                    ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arialbd.ttf")
                    : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");
                return File.ReadAllBytes(fontPath);
            }
        }

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

        public static void ReportPDF(string fileName)
        {
            if (GlobalFontSettings.FontResolver == null)
                GlobalFontSettings.FontResolver = new MyFontResolver();

            string fullPath = Path.GetFullPath(fileName);

            using (PdfDocument document = new PdfDocument())
            {
                document.Info.Title = "Отчёт по жильцам дома";

                PdfPage page = document.AddPage();

                using (XGraphics gfx = XGraphics.FromPdfPage(page))
                {
                    int marginTop = 20;
                    int marginLeft = 50;

                    XFont fontHeader = new XFont("Arial", 16, XFontStyleEx.Bold);
                    XFont font = new XFont("Arial", 12, XFontStyleEx.Regular);

                    gfx.DrawString("Список жильцов дома", fontHeader, XBrushes.Black,
                        new XRect(0, marginTop, page.Width, 15),
                        XStringFormats.Center);

                    gfx.DrawString("по адресу: г. Пермь, ул. Луначарского, д. 24", font, XBrushes.Black,
                        new XRect(0, marginTop + 30, page.Width, 10),
                        XStringFormats.Center);

                    var owners = AllOwners();

                    gfx.DrawString($"Всего жильцов: {owners.Count}", font, XBrushes.Black,
                        new XRect(marginLeft, marginTop + 70, page.Width, 10),
                        XStringFormats.CenterLeft);

                    int Width = (Convert.ToInt32(page.Width.Value) - marginLeft * 2 - 30) / 4;

                    XSolidBrush headerBrush = new XSolidBrush(XColors.LightGray);

                    gfx.DrawRectangle(headerBrush, marginLeft, marginTop + 100, Width, 20);
                    gfx.DrawRectangle(headerBrush, marginLeft + Width + 10, marginTop + 100, Width, 20);
                    gfx.DrawRectangle(headerBrush, marginLeft + (Width + 10) * 2, marginTop + 100, Width, 20);
                    gfx.DrawRectangle(headerBrush, marginLeft + (Width + 10) * 3, marginTop + 100, Width, 20);

                    gfx.DrawString("№", font, XBrushes.Black,
                        new XRect(marginLeft, marginTop + 100, Width, 20), XStringFormats.Center);
                    gfx.DrawString("Фамилия", font, XBrushes.Black,
                        new XRect(marginLeft + Width + 10, marginTop + 100, Width, 20), XStringFormats.Center);
                    gfx.DrawString("Имя", font, XBrushes.Black,
                        new XRect(marginLeft + (Width + 10) * 2, marginTop + 100, Width, 20), XStringFormats.Center);
                    gfx.DrawString("Отчество", font, XBrushes.Black,
                        new XRect(marginLeft + (Width + 10) * 3, marginTop + 100, Width, 20), XStringFormats.Center);

                    for (int i = 0; i < owners.Count; i++)
                    {
                        int yPos = marginTop + 100 + 25 * (i + 1);

                        gfx.DrawRectangle(headerBrush, marginLeft, yPos, Width, 20);
                        gfx.DrawRectangle(headerBrush, marginLeft + Width + 10, yPos, Width, 20);
                        gfx.DrawRectangle(headerBrush, marginLeft + (Width + 10) * 2, yPos, Width, 20);
                        gfx.DrawRectangle(headerBrush, marginLeft + (Width + 10) * 3, yPos, Width, 20);

                        gfx.DrawString((i + 1).ToString(), font, XBrushes.Black,
                            new XRect(marginLeft, yPos, Width, 20), XStringFormats.Center);

                        gfx.DrawString(owners[i].LastName, font, XBrushes.Black,
                            new XRect(marginLeft + Width + 10, yPos, Width, 20), XStringFormats.Center);

                        gfx.DrawString(owners[i].FirstName, font, XBrushes.Black,
                            new XRect(marginLeft + (Width + 10) * 2, yPos, Width, 20), XStringFormats.Center);

                        gfx.DrawString(owners[i].SureName, font, XBrushes.Black,
                            new XRect(marginLeft + (Width + 10) * 3, yPos, Width, 20), XStringFormats.Center);
                    }
                }
                document.Save(fullPath);
            } 
        }

        private static void FillCell(Word.Cell cell, string text, bool isBold = false)
        {
            cell.Range.Text = text;
            cell.Range.Font.Bold = isBold ? 1 : 0;
            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
        }
    }
}