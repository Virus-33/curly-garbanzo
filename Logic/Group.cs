using NPOI.HSSF.Model;
using System.Collections.Generic;

namespace Logic
{
    public enum GroupGrade
    {
        bachelor,
        magistracy,
        aspirant
    }

    public enum GroupType
    {
        absentia,
        intramural
    }

    /// <summary>
    /// Класс, отвечающий за хранение данных о группе.
    /// Какого она типа, и какие работы по сколько часов были проведены с этой группой.
    /// </summary>
    public class Group
    {
        public string code;
        public GroupGrade grade;
        public int course;
        public GroupType type;
        public Dictionary<string, int> workload;

        public Group()
        {

        }

        public Group(string code, GroupGrade grade, int course, GroupType type, Dictionary<string, int> load)
        {
            this.code = code;
            this.grade = grade;
            this.course = course;
            this.type = type;
            workload = load;
        }
    }
}
