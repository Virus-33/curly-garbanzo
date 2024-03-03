using System.ComponentModel;
using System.IO;
using System;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Logic
{
    /// <summary>
    /// Этот класс отвечает за вывод готового отчёта в файл.
    /// </summary>
    public class FileWriter
    {
        /// <summary>
        /// Записывает отчёт в файл
        /// </summary>
        /// <param name="report">Отчёт</param>
        /// <param name="path">Путь к выходному файлу</param>
        public static void WriteFile(Report report, string path)
        {
            throw new System.NotImplementedException();
        }

        void GenerateEmptyExcelReportWith2Groups(string path)
        {
            //Проверяет на существующий файл и даёт новое имя
            var desktopPath = path;
            string FilePath = GetReportPath(desktopPath);

            //Это надо, иначе NuGet пакет ругается
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            // Создание нового документа Excel
            using (ExcelPackage package = new ExcelPackage())
            {
                // Добавление нового листа
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Лист1");

                DrawMarkingInReport(worksheet);
                WriteInReport(worksheet);

                package.SaveAs(new FileInfo(FilePath));
            }

        }



        ExcelWorksheet DrawMarkingInReport(ExcelWorksheet worksheet)
        {
            ChangeColumnAllWidth(worksheet);
            ChangeRowAllHeight(worksheet);
            JoinAllCells(worksheet);
            DrawAllBorders(worksheet);
            return worksheet;
        }

        ExcelWorksheet WriteInReport(ExcelWorksheet worksheet)
        {
            WriteTextInCurrentCell(worksheet.Cells["A2:P2"], "Отчет о выполненной работе", 18);
            worksheet.Cells["A2:P2"].Style.Font.Bold = true;

            WriteTextInCurrentCell(worksheet.Cells["D3"], "за", 14);
            WriteTextInCurrentCell(worksheet.Cells["I3:J3"], "2023-2024 г.", 14);
            WriteTextInCurrentCell(worksheet.Cells["C4:D4"], "преподавателя", 14);
            WriteTextInCurrentCell(worksheet.Cells["C5:F5"], "Учебная нагрузка в часах", 14);

            WriteTextInCurrentCell(worksheet.Cells["C7:O7"], "Виды занятий", 12);
            WriteTextInCurrentCell(worksheet.Cells["B7"], "Группа", 12);
            WriteTextInCurrentCell(worksheet.Cells["C8"], "Лекции", 12);
            WriteTextInCurrentCell(worksheet.Cells["D8"], "Практ. зан.", 12);
            WriteTextInCurrentCell(worksheet.Cells["E8"], "Лаб. занятия", 12);
            WriteTextInCurrentCell(worksheet.Cells["F8"], "Консульт.", 12);
            WriteTextInCurrentCell(worksheet.Cells["G8"], "Зачёты", 12);
            WriteTextInCurrentCell(worksheet.Cells["H8"], "Экзамены", 12);
            WriteTextInCurrentCell(worksheet.Cells["I8"], "Курс. пр.", 12);
            WriteTextInCurrentCell(worksheet.Cells["J8"], "РГР", 12);
            WriteTextInCurrentCell(worksheet.Cells["K8"], "Практика", 12);
            WriteTextInCurrentCell(worksheet.Cells["L8"], "Дипл. пр.", 12);
            WriteTextInCurrentCell(worksheet.Cells["M8"], "ГЭК", 12);
            WriteTextInCurrentCell(worksheet.Cells["N8"], "Рук. магистр.", 12);
            WriteTextInCurrentCell(worksheet.Cells["O8"], "Рук. аспирант-стажерами", 12);
            WriteTextInCurrentCell(worksheet.Cells["P7"], "Всего часов", 12);

            WriteTextInCurrentCell(worksheet.Cells["A9"], "Очная форма обучения", 12);
            worksheet.Cells["A9"].Style.TextRotation = 90;

            WriteTextInCurrentCell(worksheet.Cells["A11"], "Итого", 12);
            worksheet.Cells["A11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            WriteTextInCurrentCell(worksheet.Cells["A12"], "Заочная форма обучения", 12);
            worksheet.Cells["A12"].Style.TextRotation = 90;

            WriteTextInCurrentCell(worksheet.Cells["A14"], "Итого", 12);
            worksheet.Cells["A14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            WriteTextInCurrentCell(worksheet.Cells["A15"], "Всего за месяц", 12);
            worksheet.Cells["A15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            WriteTextInCurrentCell(worksheet.Cells["A16"], "Всего от начала года", 12);
            worksheet.Cells["A16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            WriteTextInCurrentCell(worksheet.Cells["A19"], "Учебно-методическая работа", 12);
            worksheet.Cells["A19"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells["A19"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

            WriteTextInCurrentCell(worksheet.Cells["A24"], "Научно-исследовательская работа", 12);
            worksheet.Cells["A24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells["A24"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

            WriteTextInCurrentCell(worksheet.Cells["A29"], "Воспитательная работа", 12);
            worksheet.Cells["A29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells["A29"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

            WriteTextInCurrentCell(worksheet.Cells["A34"], "Прочая работа", 12);
            worksheet.Cells["A34"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells["A34"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

            WriteTextInCurrentCell(worksheet.Cells["J42"], "подпись преподавателя", 12);
            worksheet.Cells["J42"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

            return worksheet;
        }

        string GetReportPath(string desktopPath)
        {
            string filePath = Path.Combine(desktopPath, "Отчёт.xlsx");

            int count = 1;
            while (File.Exists(filePath))
            {
                filePath = Path.Combine(desktopPath, $"Отчёт{count}.xlsx");
                count++;
            }
            return filePath;
        }

        ExcelWorksheet ChangeColumnAllWidth(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 7.14;
            worksheet.Column(2).Width = 12.86;

            for (int i = 3; i < 15; i++)
            {
                worksheet.Column(i).Width = 10.43;
            }

            worksheet.Column(15).Width = 12.71;
            worksheet.Column(16).Width = 10.43;
            return worksheet;
        }

        ExcelWorksheet ChangeRowAllHeight(ExcelWorksheet worksheet)
        {
            worksheet.Row(1).Height = 23;
            worksheet.Row(2).Height = 24.5;
            worksheet.Row(3).Height = 21;
            worksheet.Row(4).Height = 27;
            worksheet.Row(5).Height = 23;
            worksheet.Row(6).Height = 17;
            worksheet.Row(7).Height = 17;

            for (int i = 8; i < 11; i++)
            {
                worksheet.Row(i).Height = 68;
            }

            worksheet.Row(11).Height = 29;
            worksheet.Row(12).Height = 47;
            worksheet.Row(13).Height = 47;
            worksheet.Row(14).Height = 24;

            for (int i = 15; i < 18; i++)
            {
                worksheet.Row(i).Height = 30;
            }

            worksheet.Row(18).Height = 16;

            for (int i = 19; i < 38; i++)
            {
                worksheet.Row(i).Height = 22;
            }

            worksheet.Row(41).Height = 17;
            worksheet.Row(42).Height = 17;

            return worksheet;
        }

        ExcelWorksheet JoinAllCells(ExcelWorksheet worksheet)
        {
            JoinCells(worksheet, "A2:P2");
            JoinCells(worksheet, "E3:H3");
            JoinCells(worksheet, "I3:J3");
            JoinCells(worksheet, "C4:D4");
            JoinCells(worksheet, "E4:L4");
            JoinCells(worksheet, "C5:F5");
            JoinCells(worksheet, "G5:I5");
            JoinCells(worksheet, "A7:A8");
            JoinCells(worksheet, "B7:B8");
            JoinCells(worksheet, "C7:O7");
            JoinCells(worksheet, "P7:P8");
            JoinCells(worksheet, "A9:A10");
            JoinCells(worksheet, "A11:B11");
            JoinCells(worksheet, "A12:A13");
            JoinCells(worksheet, "A14:B14");
            JoinCells(worksheet, "A15:B15");
            JoinCells(worksheet, "A16:B16");

            for (int i = 19; i < 39; i = i + 5)
            {
                JoinCells(worksheet, $"A{i}:D{i}");
                JoinCells(worksheet, $"A{i + 1}:P{i + 1}");
                JoinCells(worksheet, $"A{i + 2}:P{i + 2}");
                JoinCells(worksheet, $"A{i + 3}:P{i + 3}");
            }

            JoinCells(worksheet, "J41:P41");
            JoinCells(worksheet, "J42:P42");

            return worksheet;
        }

        ExcelWorksheet JoinCells(ExcelWorksheet worksheet, string CellsCoordinate)
        {
            ExcelRange cells = worksheet.Cells[CellsCoordinate];
            cells.Merge = true;
            return worksheet;
        }

        ExcelWorksheet DrawAllBorders(ExcelWorksheet worksheet)
        {
            DrawBottomBorders(worksheet, "E3:H3");
            DrawBottomBorders(worksheet, "J3");
            DrawBottomBorders(worksheet, "E4:L4");
            DrawBottomBorders(worksheet, "G5:I5");

            for (int i = 19; i < 39; i = i + 5)
            {
                DrawBottomBorders(worksheet, $"E{i}:$P{i}");
                DrawBottomBorders(worksheet, $"A{i + 1}:P{i + 1}");
                DrawBottomBorders(worksheet, $"A{i + 2}:P{i + 2}");
                DrawBottomBorders(worksheet, $"A{i + 3}:P{i + 3}");
            }

            DrawBottomBorders(worksheet, "J41:P41");

            DrawAroundBorders(worksheet, 7, 1, 8, 1);
            DrawAroundBorders(worksheet, 7, 2, 8, 2);
            DrawAroundBorders(worksheet, 7, 3, 7, 15);
            DrawAroundBorders(worksheet, 7, 16, 8, 16);

            for (int i = 3; i < 16; i++)
            {
                DrawAroundBorders(worksheet, 8, i, 8, i);
            }

            DrawAroundBorders(worksheet, 9, 1, 10, 1);

            for (int i = 2; i < 17; i++)
            {
                DrawAroundBorders(worksheet, 9, i, 9, i);
                DrawAroundBorders(worksheet, 10, i, 10, i);
            }

            DrawAroundBorders(worksheet, 11, 1, 11, 2);

            for (int i = 3; i < 17; i++)
            {
                DrawAroundBorders(worksheet, 11, i, 11, i);
            }

            DrawAroundBorders(worksheet, 12, 1, 13, 1);

            for (int i = 2; i < 17; i++)
            {
                DrawAroundBorders(worksheet, 12, i, 12, i);
                DrawAroundBorders(worksheet, 13, i, 13, i);
            }

            for (int i = 14; i < 17; i++)
            {
                for (int j = 3; j < 17; j++)
                {
                    DrawAroundBorders(worksheet, i, 1, i, 2);
                    DrawAroundBorders(worksheet, i, j, i, j);
                }
            }

            return worksheet;
        }

        ExcelWorksheet DrawBottomBorders(ExcelWorksheet worksheet, string CellsCoordinate)
        {
            worksheet.Cells[CellsCoordinate].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            return worksheet;
        }

        ExcelWorksheet DrawAroundBorders(ExcelWorksheet worksheet, int RowFirstCoordinate, int ColFirstCoordinate, int RowSecondCoordinate, int ColSecondCoordinate)
        {
            worksheet.Cells[RowFirstCoordinate, ColFirstCoordinate, RowSecondCoordinate, ColSecondCoordinate].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[RowFirstCoordinate, ColFirstCoordinate, RowSecondCoordinate, ColSecondCoordinate].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[RowFirstCoordinate, ColFirstCoordinate, RowSecondCoordinate, ColSecondCoordinate].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[RowFirstCoordinate, ColFirstCoordinate, RowSecondCoordinate, ColSecondCoordinate].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            return worksheet;
        }

        ExcelRange WriteTextInCurrentCell(ExcelRange cell, string ValueText, int FontSize)
        {
            cell.Value = ValueText;
            cell.Style.Font.Size = FontSize;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            cell.Style.Font.Name = "Times New Roman";
            cell.Style.WrapText = true;

            return cell;
        }
    }
}
