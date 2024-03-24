using System;
using System.Collections.Generic;
using System.Linq;

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
        /// <returns>Новый экземпляр класса Report с заполненными полями</returns>
        public static Report Shape(DateTime month, string teacher, List<Group> teacherWorkload)
        {
            Report shaper = new Report();
            shaper.month = month.ToString("MMMM");
            shaper.teacher = teacher;
            shaper.totalWorkload = teacherWorkload.Sum(x => x.workload.Sum(y => y.Value));
            shaper.intramuralGroups = teacherWorkload.Where(x => x.type == GroupType.intramural).ToList();
            shaper.absentiaGroups = teacherWorkload.Where(x => x.type == GroupType.absentia).ToList();
            shaper.intramuralSummary = shaper.intramuralGroups.Sum(x => x.workload.Sum(y => y.Value));
            shaper.absentiaSummary = shaper.absentiaGroups.Sum(x => x.workload.Sum(y => y.Value));
            shaper.monthlySummary = shaper.totalWorkload;
            shaper.yearlySummary = 0;
            shaper.methodicalWorks = new List<string>();
            shaper.scientificWorks = new List<string>();
            shaper.pedagogicalWorks = new List<string>();
            shaper.otherWorks = new List<string>();
            foreach (Group group in teacherWorkload)
            {
                foreach (KeyValuePair<string, int> workloadItem in group.workload)
                {
                    switch (workloadItem.Key)
                    {
                        case "Лекция":
                            shaper.methodicalWorks.Add(workloadItem.Value.ToString());
                            break;
                        case "Практическое занятие":
                            shaper.scientificWorks.Add(workloadItem.Value.ToString());
                            break;
                        case "Лабораторная работа":
                            shaper.pedagogicalWorks.Add(workloadItem.Value.ToString());
                            break;
                        case "Консультация":
                            shaper.otherWorks.Add(workloadItem.Value.ToString());
                            break;
                        case "Зачет":
                            shaper.otherWorks.Add(workloadItem.Value.ToString());
                            break;
                        case "Экзамен":
                            shaper.otherWorks.Add(workloadItem.Value.ToString());
                            break;
                        case "Курсовой проект":
                            shaper.otherWorks.Add(workloadItem.Value.ToString());
                            break;
                        case "КР":
                            shaper.otherWorks.Add(workloadItem.Value.ToString());
                            break;
                        case "Практика":
                            shaper.otherWorks.Add(workloadItem.Value.ToString());
                            break;
                        case "Выпускная квалификационная работа":
                            shaper.otherWorks.Add(workloadItem.Value.ToString());
                            break;
                        case "ГЭК":
                            shaper.otherWorks.Add(workloadItem.Value.ToString());
                            break;
                        case "Руководство магистрантами":
                            shaper.otherWorks.Add(workloadItem.Value.ToString());
                            break;
                        case "Руководство практикой":
                            shaper.otherWorks.Add(workloadItem.Value.ToString());
                            break;
                        default:
                            shaper.otherWorks.Add(workloadItem.Value.ToString());
                            break;
                    }
                }
            }
            return shaper;
        }


    }
}
