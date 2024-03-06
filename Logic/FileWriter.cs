using System.ComponentModel;
using System.IO;
using System;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Linq;
using System.Collections.Generic;

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
            FileWriter fileWriter = new FileWriter();
            fileWriter.GenerateExcelReport(path,report);
        }

        void GenerateExcelReport(string path, Report report)
        {
            //Проверяет на существующий файл и даёт новое имя
            var desktopPath = path;
            string FilePath = GetReportPath(desktopPath);

            //Это надо, иначе NuGet пакет ругается
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            // Создание нового документа Excel
            using (ExcelPackage package = new ExcelPackage())
            {
                int CountIntramuralGroups = report.intramuralGroups.Count;
                int CountAbsentiaGroups = report.absentiaGroups.Count;

                // Добавление нового листа
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Лист1");

                DrawMarkingInReport(worksheet, CountIntramuralGroups, CountAbsentiaGroups);
                WriteInReport(worksheet, CountIntramuralGroups, CountAbsentiaGroups);

                WriteDataInReport(worksheet, report);

                package.SaveAs(new FileInfo(FilePath));
            }

        }

        ExcelWorksheet WriteDataInReport(ExcelWorksheet worksheet, Report report)
        {
            WriteTextInCurrentCell(worksheet.Cells["E3"], report.month, 14);
            WriteTextInCurrentCell(worksheet.Cells["E4"], report.teacher, 14);
            WriteTextInCurrentCell(worksheet.Cells["G5"], Convert.ToString(report.totalWorkload), 14);

            int CurrentCell = 9;

            //очники
            for (int i = 0; i < report.intramuralGroups.Count(); i++)
            {
                Group currentGroup = report.intramuralGroups[i];
                WriteTextInCurrentCell(worksheet.Cells[$"B{CurrentCell}"], currentGroup.code, 12);

                for (int j = 3; j < 17; j++)
                {
                    foreach (KeyValuePair<string, int> element in currentGroup.workload)
                    {
                        if (Convert.ToString(worksheet.Cells[8, j]) == element.Key)
                        {
                            WriteTextInCurrentCell(worksheet.Cells[CurrentCell,j], Convert.ToString(element.Value),12);
                        }
                    }
                }
                CurrentCell++;                                
            }
            WriteTextInCurrentCell(worksheet.Cells[$"P{CurrentCell}"], Convert.ToString(report.intramuralSummary), 12);
            CurrentCell++;

            //заочники
            for (int i = 0; i < report.absentiaGroups.Count(); i++)
            {
                Group currentGroup = report.absentiaGroups[i];
                WriteTextInCurrentCell(worksheet.Cells[$"B{CurrentCell}"], currentGroup.code, 12);

                for (int j = 3; j < 17; j++)
                {
                    foreach (KeyValuePair<string, int> element in currentGroup.workload)
                    {
                        if (Convert.ToString(worksheet.Cells[8, j]) == element.Key)
                        {
                            WriteTextInCurrentCell(worksheet.Cells[CurrentCell, j], Convert.ToString(element.Value), 12);
                        }
                    }
                }
                CurrentCell++;
            }
            WriteTextInCurrentCell(worksheet.Cells[$"P{CurrentCell}"], Convert.ToString(report.absentiaSummary), 12);
            CurrentCell++;

            WriteTextInCurrentCell(worksheet.Cells[$"P{CurrentCell}"], Convert.ToString(report.monthlySummary), 12);
            CurrentCell++;

            WriteTextInCurrentCell(worksheet.Cells[$"P{CurrentCell}"], Convert.ToString(report.yearlySummary), 12);
            CurrentCell= CurrentCell + 4;

            return worksheet;
        }

        ExcelWorksheet DrawMarkingInReport(ExcelWorksheet worksheet, int CountIntramuralGroups, int CountAbsentiaGroups)
        {
            ChangeColumnAllWidth(worksheet);

            ChangeRowAllHeight(worksheet, CountIntramuralGroups, CountAbsentiaGroups);
            JoinAllCells(worksheet, CountIntramuralGroups, CountAbsentiaGroups);
            DrawAllBorders(worksheet, CountIntramuralGroups, CountAbsentiaGroups);
            return worksheet;
        }

        ExcelWorksheet WriteInReport(ExcelWorksheet worksheet, int CountIntramuralGroups, int CountAbsentiaGroups)
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

            int CurrentCoordinate = 9;

            //Очные 
            WriteTextInCurrentCell(worksheet.Cells[$"A{CurrentCoordinate}"], "Очная форма обучения", 12);
            worksheet.Cells[$"A{CurrentCoordinate}"].Style.TextRotation = 90;
            WriteTextInCurrentCell(worksheet.Cells[$"A{CurrentCoordinate + CountIntramuralGroups}"], "Итого", 12);
            worksheet.Cells[$"A{CurrentCoordinate + CountIntramuralGroups}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            CurrentCoordinate = CurrentCoordinate + CountIntramuralGroups + 1;

            //Заочные 
            WriteTextInCurrentCell(worksheet.Cells[$"A{CurrentCoordinate}"], "Заочная форма обучения", 12);
            worksheet.Cells[$"A{CurrentCoordinate}"].Style.TextRotation = 90;
            WriteTextInCurrentCell(worksheet.Cells[$"A{CurrentCoordinate + CountAbsentiaGroups}"], "Итого", 12);
            worksheet.Cells[$"A{CurrentCoordinate + CountAbsentiaGroups}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            CurrentCoordinate = CurrentCoordinate + CountAbsentiaGroups + 1;


            WriteTextInCurrentCell(worksheet.Cells[$"A{CurrentCoordinate}"], "Всего за месяц", 12);
            worksheet.Cells[$"A{CurrentCoordinate}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            WriteTextInCurrentCell(worksheet.Cells[$"A{CurrentCoordinate + 1}"], "Всего от начала года", 12);
            worksheet.Cells[$"A{CurrentCoordinate + 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            WriteTextInCurrentCell(worksheet.Cells[$"A{CurrentCoordinate + 4}"], "Учебно-методическая работа", 12);
            worksheet.Cells[$"A{CurrentCoordinate + 4}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[$"A{CurrentCoordinate + 4}"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

            WriteTextInCurrentCell(worksheet.Cells[$"A{CurrentCoordinate + 9}"], "Научно-исследовательская работа", 12);
            worksheet.Cells[$"A{CurrentCoordinate + 9}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[$"A{CurrentCoordinate + 9}"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

            WriteTextInCurrentCell(worksheet.Cells[$"A{CurrentCoordinate + 14}"], "Воспитательная работа", 12);
            worksheet.Cells[$"A{CurrentCoordinate + 14}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[$"A{CurrentCoordinate + 14}"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

            WriteTextInCurrentCell(worksheet.Cells[$"A{CurrentCoordinate + 19}"], "Прочая работа", 12);
            worksheet.Cells[$"A{CurrentCoordinate + 19}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[$"A{CurrentCoordinate + 19}"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

            WriteTextInCurrentCell(worksheet.Cells[$"J{CurrentCoordinate + 27}"], "подпись преподавателя", 12);
            worksheet.Cells[$"J{CurrentCoordinate + 27}"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

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

        ExcelWorksheet ChangeRowAllHeight(ExcelWorksheet worksheet, int CountIntramuralGroups, int CountAbsentiaGroups)
        {
            worksheet.Row(1).Height = 23;
            worksheet.Row(2).Height = 24.5;
            worksheet.Row(3).Height = 21;
            worksheet.Row(4).Height = 27;
            worksheet.Row(5).Height = 23;
            worksheet.Row(6).Height = 17;
            worksheet.Row(7).Height = 17;
            worksheet.Row(8).Height = 68;

            int CurrentCoordinate = 9;

            //Высота ячеек для Очных групп
            for (int i = CurrentCoordinate; i < CurrentCoordinate + CountIntramuralGroups; i++)
            {
                worksheet.Row(i).Height = 68;
            }
            worksheet.Row(CurrentCoordinate + CountIntramuralGroups).Height = 30;
            CurrentCoordinate = CurrentCoordinate + CountIntramuralGroups + 1;

            //Высота ячеек для заочных групп
            for (int i = CurrentCoordinate; i < CurrentCoordinate + CountAbsentiaGroups; i++)
            {
                worksheet.Row(i).Height = 68;
            }
            worksheet.Row(CurrentCoordinate + CountAbsentiaGroups).Height = 30;
            CurrentCoordinate = CurrentCoordinate + CountAbsentiaGroups + 1;


            for (int i = CurrentCoordinate; i < CurrentCoordinate + 3; i++)
            {
                worksheet.Row(i).Height = 30;
            }

            worksheet.Row(CurrentCoordinate + 3).Height = 16;
            CurrentCoordinate = CurrentCoordinate + 4;

            for (int i = CurrentCoordinate; i < CurrentCoordinate + 19; i++)
            {
                worksheet.Row(i).Height = 22;
            }

            CurrentCoordinate = CurrentCoordinate + 22;
            worksheet.Row(CurrentCoordinate).Height = 17;
            worksheet.Row(CurrentCoordinate + 1).Height = 17;

            return worksheet;
        }

        ExcelWorksheet JoinAllCells(ExcelWorksheet worksheet, int CountIntramuralGroups, int CountAbsentiaGroups)
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

            int CurrentCoordinate = 9;

            //Ячейка под Очные группы
            JoinCells(worksheet, $"A{CurrentCoordinate}:A{CurrentCoordinate + CountIntramuralGroups - 1}");
            CurrentCoordinate = CurrentCoordinate + CountIntramuralGroups;

            //Итог очных
            JoinCells(worksheet, $"A{CurrentCoordinate}:B{CurrentCoordinate}");
            CurrentCoordinate = CurrentCoordinate + 1;

            //Ячейка под заочные группы
            JoinCells(worksheet, $"A{CurrentCoordinate}:A{CurrentCoordinate + CountAbsentiaGroups - 1}");
            CurrentCoordinate = CurrentCoordinate + CountAbsentiaGroups;

            //Итог заочных
            JoinCells(worksheet, $"A{CurrentCoordinate}:B{CurrentCoordinate}");
            CurrentCoordinate = CurrentCoordinate + 1;


            JoinCells(worksheet, $"A{CurrentCoordinate}:B{CurrentCoordinate}");
            JoinCells(worksheet, $"A{CurrentCoordinate + 1}:B{CurrentCoordinate + 1}");

            CurrentCoordinate = CurrentCoordinate + 4;

            for (int i = CurrentCoordinate; i < CurrentCoordinate + 19; i = i + 5)
            {
                JoinCells(worksheet, $"A{i}:D{i}");
                JoinCells(worksheet, $"A{i + 1}:P{i + 1}");
                JoinCells(worksheet, $"A{i + 2}:P{i + 2}");
                JoinCells(worksheet, $"A{i + 3}:P{i + 3}");
            }

            CurrentCoordinate = CurrentCoordinate + 22;

            JoinCells(worksheet, $"J{CurrentCoordinate}:P{CurrentCoordinate}");
            JoinCells(worksheet, $"J{CurrentCoordinate + 1}:P{CurrentCoordinate + 1}");

            return worksheet;
        }

        ExcelWorksheet JoinCells(ExcelWorksheet worksheet, string CellsCoordinate)
        {
            ExcelRange cells = worksheet.Cells[CellsCoordinate];
            cells.Merge = true;
            return worksheet;
        }

        ExcelWorksheet DrawAllBorders(ExcelWorksheet worksheet, int CountIntramuralGroups, int CountAbsentiaGroups)
        {
            DrawBottomBorders(worksheet, "E3:H3");
            DrawBottomBorders(worksheet, "J3");
            DrawBottomBorders(worksheet, "E4:L4");
            DrawBottomBorders(worksheet, "G5:I5");


            DrawAroundBorders(worksheet, 7, 1, 8, 1);
            DrawAroundBorders(worksheet, 7, 2, 8, 2);
            DrawAroundBorders(worksheet, 7, 3, 7, 15);
            DrawAroundBorders(worksheet, 7, 16, 8, 16);

            for (int i = 3; i < 16; i++)
            {
                DrawAroundBorders(worksheet, 8, i, 8, i);
            }

            int CurrentCoordinate = 9;

            //Для Очных
            DrawAroundBorders(worksheet, CurrentCoordinate, 1, CurrentCoordinate + CountIntramuralGroups - 1, 1);

            for (int j = 0; j < CountIntramuralGroups; j++)
            {
                for (int i = 2; i < 17; i++)
                {
                    DrawAroundBorders(worksheet, CurrentCoordinate, i, CurrentCoordinate, i);
                }
                CurrentCoordinate++;

            }
            //Итог очных
            DrawAroundBorders(worksheet, CurrentCoordinate, 1, CurrentCoordinate, 2);

            for (int i = 3; i < 17; i++)
            {
                DrawAroundBorders(worksheet, CurrentCoordinate, i, CurrentCoordinate, i);
            }

            CurrentCoordinate = CurrentCoordinate + 1;

            //Для заочных
            DrawAroundBorders(worksheet, CurrentCoordinate, 1, CurrentCoordinate + CountAbsentiaGroups - 1, 1);

            for (int j = 0; j < CountAbsentiaGroups; j++)
            {
                for (int i = 2; i < 17; i++)
                {
                    DrawAroundBorders(worksheet, CurrentCoordinate, i, CurrentCoordinate, i);
                }
                CurrentCoordinate++;

            }

            //Итог заочных
            DrawAroundBorders(worksheet, CurrentCoordinate, 1, CurrentCoordinate, 2);

            for (int i = 3; i < 17; i++)
            {
                DrawAroundBorders(worksheet, CurrentCoordinate, i, CurrentCoordinate, i);
            }

            CurrentCoordinate = CurrentCoordinate + 1;


            for (int i = CurrentCoordinate; i < CurrentCoordinate + 2; i++)
            {
                DrawAroundBorders(worksheet, i, 1, i, 2);
                for (int j = 3; j < 17; j++)
                {
                    DrawAroundBorders(worksheet, i, j, i, j);
                }
            }

            CurrentCoordinate = CurrentCoordinate + 4;

            for (int i = CurrentCoordinate; i < CurrentCoordinate + 19; i = i + 5)
            {
                DrawBottomBorders(worksheet, $"E{i}:$P{i}");
                DrawBottomBorders(worksheet, $"A{i + 1}:P{i + 1}");
                DrawBottomBorders(worksheet, $"A{i + 2}:P{i + 2}");
                DrawBottomBorders(worksheet, $"A{i + 3}:P{i + 3}");
            }

            CurrentCoordinate = CurrentCoordinate + 22;

            DrawBottomBorders(worksheet, $"J{CurrentCoordinate}:P{CurrentCoordinate}");

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
