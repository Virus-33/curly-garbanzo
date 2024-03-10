using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Logic.Utility;
using NPOI.HSSF.Record;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace Logic.WorkloadParser
{
    public static class WorkloadParser
    {
        const int groupCapacity = 30;
        //Тоже бы в конфиг вынести.... здесь колонки которые мы не учитываем при парсинге
        static List<string> cellsToSkip = new List<string>()
        {
            "данные к расчету",
            "контингент студентов",
            "всего часов",
            "доля ставки",
            "примечание",
            "кол- во гру пп",
            "кол- во не- дель",
            "колич. часов на группу в день",
            "часов по  плану на группу",
            "колич. часов на одно го дипл.",
            "колич. дип-\nлом-\nни-\nков",
            "число членов гэк",
            "кол- во по- то-ков",
            "колич. дней",
            "ка-\nфед-\nрой",
            "факу-\nльте-\nтом"

        };
        /// <summary>
        /// Достаёт данные из календарного учебного графика
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        /// <param name="teacherName">Имя преподавателя, по которому ищется активность</param>
        /// <returns>Список групп, у которых преподаватель что-либо вёл. Свойства групп должны быть заполнены</returns>
        public static List<Group> Parse(string path, string teacherName)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            List<IRow> rowsToResearch = new List<IRow>();
            IRow headerRow = null;
            IRow subheaderRow = null;
            XSSFWorkbook workbook;
            using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(fileStream);
            }
            var sheet = workbook.GetSheetAt(0);
            foreach (IRow row in sheet)
            {
                if (row.Cells[0].StringCellValue == "Дисциплины")
                {
                    headerRow = row;
                    subheaderRow = sheet.GetRow(row.RowNum + 1);
                }
                if (row.Cells[0].StringCellValue.ToLower().Contains(teacherName.ToLower()))
                {
                    rowsToResearch.AddRange(GetRowsToResearch(sheet, row));
                }
            }
            return GetGroups(headerRow, subheaderRow, rowsToResearch);
        }

        private static List<IRow> GetRowsToResearch(ISheet sheet, IRow startRow)
        {
            // Константа, можно вынести в конфиг. Выбираем из таблицы набор строк, которые относятся к нужному преподу, чтобы дальше уже считывать из них все.
            string endRowName = "ИТОГО";
            //Тоже константа, можно вынести в конфиг. нужна чтобы лишние строчки отсекать.
            string rowNamesToSkip = "Всего по 1 семестру Всего по 2 семестру Итого по очной ф. о. Итого по зачной ф. о.";
            List<IRow> rowsToResearch = new List<IRow>();
            for (int i = startRow.RowNum; i < sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row.Cells[0].StringCellValue == endRowName)
                {
                    break;
                }
                if ((row.Cells.Where(x => x.CellType == CellType.Numeric).ToList().Count > 1 && !rowNamesToSkip.Contains(row.Cells[0].StringCellValue)) || row.Cells[0].StringCellValue.Contains("форма обучения"))
                {

                    rowsToResearch.Add(row);
                }
            }
            return rowsToResearch;
        }

        private static List<Group> GetGroups(IRow headerRow, IRow subheaderRow, List<IRow> rowsToResearch)
        {
            var groups = new List<Group>();
            foreach(var row in rowsToResearch)
            {
                var cells = row.Cells.Where(x => !cellsToSkip.Contains(subheaderRow.Cells[x.ColumnIndex].StringCellValue.ToLower()) && !cellsToSkip.Contains(headerRow.Cells[x.ColumnIndex].StringCellValue.ToLower()) && x.CellType == CellType.Numeric);
                if(cells.ToList().Count==0)
                {
                    continue;
                }
                Group group = new Group();
                group.code = row.Cells[headerRow.Cells.Find(x => x.StringCellValue.ToString() == "Наименование направ-\nления подготовки, \nспе-\nциаль-\nности").ColumnIndex].StringCellValue;
                group.grade = group.code.ToGrade();
                group.course = CorrectCourse(headerRow, row);
                string groupType = rowsToResearch
                    .Where(x => x.RowNum < row.RowNum)
                    .Where(x => x.Cells[0].StringCellValue.Contains("форма обучения"))
                    .First().Cells[0].StringCellValue == "Заочная форма обучения" ? "Заочная форма обучения" : "Очная форма обучения";
                
                group.type = groupType.ToEnum();

                Dictionary<string, int> workLoad = new Dictionary<string, int>();
                foreach (var cell in cells)
                {
                    workLoad.Add(headerRow.Cells.Find(x => x.Address.Column == cell.Address.Column).ToFullForm(headerRow), Convert.ToInt32(cell.NumericCellValue));
                }
                group.workload = workLoad;
                groups.Add(group);
            }
            return groups;
        }

        private static int CorrectCourse(IRow headerRow, IRow row)
        {
            string courceCellValue = row.Cells[headerRow.Cells.Find(x => x.StringCellValue == "Курс").ColumnIndex].StringCellValue;
            if (courceCellValue == string.Empty)
            {
                return 0;
            }
            if(courceCellValue.Contains('/'))
            {
                return Convert.ToInt32(courceCellValue[0]);
            }
            return Convert.ToInt32(courceCellValue);
        }
    }
}
