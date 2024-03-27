using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Logic.Utility
{
    public class GroupPreprocessor
    {
        public static List<Group> MergeByAffiliation(List<Group> planned, List<Group> teacher)
        {
            List<Group> res = new();

            foreach (Group g1 in planned)
            {
                foreach (Group g2 in teacher)
                {
                    if (g1.code == g2.code)
                    {
                        res.Add(new Group(code: g2.code, grade: g2.grade, course: g2.course, type: g2.type,load: g1.workload));
                    }
                }
            }

            return res;
        }
    }
}
