using System.Collections.Generic;

namespace Logic
{
    /// <summary>
    /// Класс отчёта. Хранит в себе всё то, что пойдёт в отчёт в отдельных полях.
    /// </summary>
    public class Report
    {
        public string month;
        public string teacher;
        public int totalWorkload;
        public List<Group> intramuralGroups;
        public List<Group> absentiaGroups;
        

        public Report()
        {

        }
    }
}
