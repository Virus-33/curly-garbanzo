using System;
using System.Collections.Generic;

namespace Logic
{
    // Я не знал куда будет лучше прописать перечисления, чтобы они были доступны из любой части сборки.
    // Пусть пока будут здесь.

    enum WorkType
    {
        lecture,
        practiceWork,
        lab,
        consultation,
        unmarkedExam,
        markedExam,
        courseProject,
        CGW,
        practice,
        graduationWork,
        GEC,
        magistracyLeader,
        internLeader
    }

    /// <summary>
    /// Класс, отвечающий за формирование отчёта и заполнение его полей на основе остальной информации.
    /// </summary>
    public static class ReportShaper
    {
        /// <summary>
        /// Какая именно группа по итогу там окажется (11 или 12) не принципиально
        /// </summary>
        /// <param name="teacherWorkload">Принимает нагрузку из файла по преподавателям</param>
        /// <param name="plannedWorkload">Принимает нагрузку из файла по студентам</param>
        /// <returns>Новый экземпляр класса Report с заполненными полями</returns>
        public static Report Shape(DateTime month, string teacher, List<Group> teacherWorkload, List<Group> plannedWorkload)
        {
            throw new System.NotImplementedException();
        }
    }
}
