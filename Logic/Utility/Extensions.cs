using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Logic.Utility
{
    public static class Extensions
    {
        static List<string> _bachelorGroups = new List<string>() { "БПИ", "БИСТ", "МЛ" };
        static List<string> _magistracyGroups = new List<string>() { "МИСТ", "ММЛ" };
        public static GroupGrade ToGrade(this string value)
        {
            if(_bachelorGroups.Contains(value))
            {
                return GroupGrade.bachelor;
            }
            if (_magistracyGroups.Contains(value))
            {
                return GroupGrade.magistracy;
            }
            return GroupGrade.aspirant;
        }
        public static GroupType ToEnum(this string value)
        {
            if(value == "Заочная форма обучения")
            {
                return GroupType.absentia;
            }
            return GroupType.intramural;
        }
        public static string ToFullForm(this ICell cell, IRow headerRow)
        {
            string fullForm = cell.StringCellValue;
            if (headerRow.Cells[cell.ColumnIndex].StringCellValue == string.Empty)
            {
                for (int i = cell.ColumnIndex; i > 0; i--)
                {
                    if (headerRow.Cells[i].StringCellValue != string.Empty)
                    {
                        fullForm = fullForm.Insert(0, headerRow.Cells[i].StringCellValue);
                        break;
                    }
                }
            }
            return fullForm;
        }
    }
}
