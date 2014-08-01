using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassExamTop3
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void main()
        {
            //全班前三名名單(評量成績)
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["班級", "資料統計"];
            item1["報表"]["成績相關報表"]["全班前三名名單(評量成績)"].Enable = false;
            item1["報表"]["成績相關報表"]["全班前三名名單(評量成績)"].Click += delegate
            {
                new ExamReporter().ShowDialog();
            };

            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                item1["報表"]["成績相關報表"]["全班前三名名單(評量成績)"].Enable = K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0 && Permissions.ExamReporter權限;
            };

            //全班前三名名單(評量成績)權限設定
            Catalog permission = RoleAclSource.Instance["班級"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.ExamReporter, "全班前三名名單(評量成績)"));

            //全班前三名名單(學期成績)
            FISCA.Presentation.RibbonBarItem item2 = FISCA.Presentation.MotherForm.RibbonBarItems["班級", "資料統計"];
            item2["報表"]["成績相關報表"]["全班前三名名單(學期成績)"].Enable = false;
            item2["報表"]["成績相關報表"]["全班前三名名單(學期成績)"].Click += delegate
            {
                new SemsReporter().ShowDialog();
            };

            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                item2["報表"]["成績相關報表"]["全班前三名名單(學期成績)"].Enable = K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0 && Permissions.SemsReporter權限;
            };

            //全班前三名名單(學期成績)權限設定
            Catalog permission2 = RoleAclSource.Instance["班級"]["功能按鈕"];
            permission2.Add(new RibbonFeature(Permissions.SemsReporter, "全班前三名名單(學期成績)"));
        }
    }
}
