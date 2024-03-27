using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using Logic.Utility;
using MathNet.Numerics;
using NPOI.OpenXmlFormats.Dml.Diagram;
using NPOI.POIFS.Storage;
using NPOI.SS.Formula.Functions;

namespace Logic
{
    public static class CalendarParser
    {
        private const int _startPositionRow = 2;
        private const int _startPositionColumn = 3;


        /// <summary>
        /// Достаёт данные из календарного учебного графика
        /// </summary>
        /// <param name="path">Путь к файлу с календарным учебным графиком</param>
        /// <returns>Список, состоящий из направлений и нагрузок</returns>
        public static List<Group> Parse(string path, string month) //ПОМЕНЯЙТЕ ДЕЛАЛ ДЛЯ ТЕСТА
        {
            List<Group> res = new();

            var wb = new XLWorkbook(path);
            wb.Worksheets.Delete(1);//Удаляем страницу без таблиц
            List<List<Month>> monthsFromAllTables = new(); //месяца для каждой таблицы
            List<List<int>> facultsTitleRows = new(); //плашки с названием специальности для каждой таблицы (их координата по вертикали)
            foreach (var currentWorksheet in wb.Worksheets)
            {
                facultsTitleRows.Add(FindFacults(currentWorksheet));
                monthsFromAllTables.Add(FillMonths(currentWorksheet));
            }
           
            var coordinatesForOneMonth = (from t in monthsFromAllTables
                                          from m in t
                                          where m.Name == month
                                          select
                                          m).ToList();

            List<FacultLoad> loadsForAllGroups = new();
            foreach (var table in Enumerable.Range(0,coordinatesForOneMonth.Count))
            {
                var currentWS = wb.Worksheet(table + 1);
                var oneTableMonth = coordinatesForOneMonth[table];
                var oneTableFacults = facultsTitleRows[table];
                
                foreach (var facult in Enumerable.Range(0,oneTableFacults.Count-1))
                {
                    var countCourses = (oneTableFacults[facult + 1] - oneTableFacults[facult] - 1) / 6;
                    foreach (var course in Enumerable.Range(0, countCourses))
                    {
                        int lastCol = GetLastColNumber(currentWS);
                        Dictionary<string, byte> loadOneCourse = new();
                        var startRowInCourse = oneTableFacults[facult] + 6 * course;
                       
                        var courseLoads = currentWS.Range(currentWS.Cell(startRowInCourse, 1), currentWS.Cell(startRowInCourse + 6, lastCol));
                        foreach (var day in oneTableMonth.Coordinates)
                        {
                            if (!loadOneCourse.Keys.Contains(courseLoads.Cell(day.row,day.col).GetString()))
                            {
                                loadOneCourse.Add(courseLoads.Cell(day.row, day.col).GetString(), 1);
                            }
                            else
                            {
                                loadOneCourse[courseLoads.Cell(day.row, day.col).GetString()] += 1;
                            }
                        }

                        FacultLoad facultLoad = new( (byte)Int32.Parse(currentWS.Cell(startRowInCourse + 1, 1).GetString()), currentWS.Cell(oneTableFacults[facult], 1).GetString() ,loadOneCourse);
                        loadsForAllGroups.Add(facultLoad);
                    }
                }
            }

            foreach (FacultLoad fl in loadsForAllGroups)
            {
                Group temp = new();
                string[] tempS = fl.facultName.Split(' ');
                temp.course = fl.course;
                temp.workload = fl.load;
                temp.code = tempS[0];
                GroupPreprocessor.GenerateGroupCode(temp);
                res.Add(temp);
            }

            return res;
        }

        private static List<int> FindFacults(IXLWorksheet worksheet)
        {
            List<int> facultsTitleRows = new List<int>();
            foreach (int row in Enumerable.Range(1, worksheet.LastRowUsed().RowNumber()))
            {
                if (worksheet.Cell(row, 1).GetString().Length > 10)
                {
                    facultsTitleRows.Add(row);
                }
            }
            facultsTitleRows.Add(worksheet.LastRowUsed().RowNumber()+1);
            return facultsTitleRows;
        }
        private static List<Month> FillMonths(IXLWorksheet worksheet)
        {
            List<Month> months = new List<Month>();
            int lastCol = GetLastColNumber(worksheet);
            bool cancelOuterLoop = false;
            int monthLastCol = _startPositionColumn;
            int monthLastRow = _startPositionRow;

            for (int month = 0;month<12;month++)
            {
                Month currentMonth;
                if (monthLastCol != _startPositionColumn || monthLastRow != _startPositionRow)
                {
                    if (worksheet.Cell(1, monthLastCol).GetString() == "")
                    {
                        var offsetHorisontal = 1;
                        while (worksheet.Cell(1, monthLastCol + offsetHorisontal).GetString() == "")
                        {
                            offsetHorisontal++;
                        }
                        currentMonth = new Month(worksheet.Cell(1, monthLastCol+offsetHorisontal).GetString());
                    }
                    else
                    {
                        currentMonth = new Month(worksheet.Cell(1, monthLastCol).GetString());
                    }

                }
                else currentMonth = new Month(worksheet.Cell(1, monthLastCol).GetString());
                for (int col = monthLastCol; col <= lastCol; col++)
                {
                    for (int row = monthLastRow; row < 8; row++)
                    {
                        
                        if (col == lastCol && worksheet.Cell(row+1, col).GetString() == "")
                        {
                            currentMonth.AddCoordinate(new Coordinate(row, col));
                            break;
                        }
                        else
                        if (col != lastCol && worksheet.Cell(row, col).GetString() == "")
                        {
                            continue;
                        }
                        currentMonth.AddCoordinate(new Coordinate(row, col));


                        if (Convert.ToInt32(worksheet.Cell(row + 1, col).GetString()) < Convert.ToInt32(worksheet.Cell(row, col).GetString())) 
                        {
                            cancelOuterLoop = true;
                            if (row+1>7)
                            {
                                monthLastCol = col+1;
                                monthLastRow = _startPositionRow;
                            }
                            else
                            {
                                monthLastCol = col;
                                monthLastRow = row + 1;
                            }
                            break; 
                        }
                        
                        if (worksheet.Cell(_startPositionRow, col + 1).GetString() == "" && (col+1 < lastCol) && row == 7)
                        {
                            var offsetHorizontal = 1;
                            while (worksheet.Cell(_startPositionRow, col + offsetHorizontal).GetString() == "" && col+offsetHorizontal <= lastCol)
                            {  
                                offsetHorizontal++; 
                            }
                            if (row == 7 && (Convert.ToInt32(worksheet.Cell(_startPositionRow, col + offsetHorizontal).GetString()) < Convert.ToInt32(worksheet.Cell(row, col).GetString())))
                            {
                                cancelOuterLoop = true;
                                monthLastCol = col + offsetHorizontal;
                                monthLastRow = _startPositionRow;
                                break;
                            }

                            monthLastCol = col + offsetHorizontal;
                            break;
                        }
                        if (row == 7 && (Convert.ToInt32(worksheet.Cell(_startPositionRow, col+1).GetString()) < Convert.ToInt32(worksheet.Cell(row, col).GetString())))
                        {
                            cancelOuterLoop = true;
                            monthLastCol = col + 1;
                            monthLastRow = _startPositionRow;
                            break;
                        }
                    }
                    if (cancelOuterLoop) 
                    { 
                        cancelOuterLoop = false; 
                        break; 
                    }
                    else
                    {
                        monthLastRow = _startPositionRow;
                    }
                }
                months.Add(currentMonth);
            }   
            return months;
        }

        private static int GetLastColNumber(IXLWorksheet worksheet)
        {
            int lastCell = 0;
            int res = -1;
            foreach (var col in worksheet.Row(2).CellsUsed())
            {
                int.TryParse(col.GetString(), out res);
                if (res != 0) lastCell = col.Address.ColumnNumber;
            }
            return lastCell;
        }

        private class Month
        {
            string _name;
            List<Coordinate> coordinates;

            public Month(string name)
            {
                _name = name;
                coordinates = new List<Coordinate>();
            }

            public void AddCoordinate(Coordinate coordinate)
            {
                coordinates.Add(coordinate);
            }

            public string Name => _name;
            public List<Coordinate> Coordinates => coordinates;
        }

        private struct Coordinate
        {
            public int row;
            public int col;
            public Coordinate(int row, int col)
            {
                this.row = row;
                this.col = col;
            }
        }

        private class FacultLoad
        {
            public string facultName;
            public Dictionary<string, byte> load;
            public byte course;

            public FacultLoad(byte course, string facultName, Dictionary<string, byte> load)
            {
                this.course = course;
                this.facultName = facultName;
                this.load = load;
            }

        }

    }
}
