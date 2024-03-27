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
        public string code;//поля должны быть приватными 
        public GroupGrade grade;
        public int course;
        public GroupType type;
        public Dictionary<string, int> workload;

        public Group(){}

        public Group(int course, Dictionary<string, int> load, string code, GroupGrade grade , GroupType type ) //Поменяете когда будете знать откуда брать эти данные
        {
            this.code = code;
            this.grade = grade;
            this.course = course;
            this.type = type;
            workload = load;
        }
    }
}
