using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Logic.Utility
{

    public class GroupPreprocessor
    {
        static Dictionary<string, string> specs = new()
        {
            { "09.03.03", "БПИ" },
            { "38.03.02", "МЛ" },
            { "09.03.02", "БИСТ" },
            { "09.04.02", "МИСТ" },
            { "38.04.02", "ММЛ" }
        };


        public static List<Group> MergeByAffiliation(List<Group> planned, List<Group> teacher)
        {
            List<Group> res = new();

            foreach (Group g1 in planned)
            {
                foreach (Group g2 in teacher)
                {
                    if (g1.code == g2.code)
                    {
                        if (g1.workload.Keys.Contains("Э"))
                        {
                            double temp = Convert.ToDouble(g1.workload["Э"]) / 2;
                            g1.workload.Add("З", (byte)Math.Floor(temp));
                            g1.workload["Э"] = (byte)Math.Ceiling(temp);
                        }

                        res.Add(new Group(code: g2.code, grade: g2.grade, course: g2.course, type: g2.type,load: g1.workload));
                        break;
                    }
                }
            }

            return res;
        }

        public static void GenerateGroupCode(Group group)
        {
            string res = "";
            Random rnd = new();
            
            if (specs.Keys.Contains(group.code)) {
                res = specs[group.code];
            }
            res += "-" + group.course + rnd.Next(11, 13);

            var t = group.code.Split('.');

            switch (t[1])
            {
                case "0.3":
                    group.grade = GroupGrade.bachelor;
                    break;
                case "0.4":
                    group.grade = GroupGrade.magistracy;
                    break;
                default:
                    group.grade = GroupGrade.aspirant;
                    break;
            }

            group.code = res;
        }

        public static void Summator(Group group)
        {
            ref var shortcut = ref group.workload;
            shortcut.Add("Пр", 0);

            if (shortcut.Keys.Contains("У"))
            {
                shortcut["П"] += shortcut["У"];
            }
            if (shortcut.Keys.Contains("Пд"))
            {
                shortcut["П"] += shortcut["Пд"];
            }
            if (shortcut.Keys.Contains("П"))
            {
                shortcut["П"] += shortcut["П"];
            }
        }

        public static void GenerateCode(Group group)
        {
            Random rnd = new();

            group.code = group.code + "-" + group.course + rnd.Next(11, 13);
        }

    }
}
