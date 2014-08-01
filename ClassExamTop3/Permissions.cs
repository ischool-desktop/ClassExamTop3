using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassExamTop3
{
    class Permissions
    {
        public static string ExamReporter { get { return "ExamReporter.AB87E6DA-2768-46B4-A3AB-320BA2474D5D"; } }

        public static bool ExamReporter權限
        {
            get { return FISCA.Permission.UserAcl.Current[ExamReporter].Executable; }
        }

        public static string SemsReporter { get { return "SemsReporter.AB87E6DA-2768-46B4-A3AB-320BA2474D5D"; } }

        public static bool SemsReporter權限
        {
            get { return FISCA.Permission.UserAcl.Current[SemsReporter].Executable; }
        }
    }
}
