using Aspose.Cells;
using CourseGradeB;
using FISCA.Data;
using FISCA.Presentation.Controls;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace ClassExamTop3
{
    public partial class SemsReporter : BaseForm
    {
        private int _schoolYear, _semester;
        private string _ExamName;
        private BackgroundWorker _BW;
        private QueryHelper _Q;
        private Dictionary<string, StudentObj> _studentObjs;
        private Dictionary<string, List<string>> _classStudents;
        private List<ClassRecord> _classRecords;

        public SemsReporter()
        {
            InitializeComponent();
            _studentObjs = new Dictionary<string, StudentObj>();
            _classStudents = new Dictionary<string, List<string>>();
            _classRecords = K12.Data.Class.SelectByIDs(K12.Presentation.NLDPanels.Class.SelectedSource);

            _BW = new BackgroundWorker();
            _BW.DoWork += new DoWorkEventHandler(BW_DoWork);
            _BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BW_Completed);

            _Q = new QueryHelper();

            _schoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
            _semester = int.Parse(K12.Data.School.DefaultSemester);

            //Set cboSchoolYear Item
            for (int i = -2; i <= 2; i++)
                cboSchoolYear.Items.Add(_schoolYear + i);

            //Set cboSemester Item
            cboSemester.Items.Add(1);
            cboSemester.Items.Add(2);

            cboSchoolYear.Text = _schoolYear + "";
            cboSemester.Text = _semester + "";
        }

        private void BW_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            FormEnable(true);
            Workbook wb = e.Result as Workbook;
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "另存新檔";
            save.FileName = "全班前三名名單(學期成績).xls";
            save.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";

            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    wb.Save(save.FileName, Aspose.Cells.SaveFormat.Excel97To2003);
                    System.Diagnostics.Process.Start(save.FileName);
                }
                catch
                {
                    MessageBox.Show("檔案儲存失敗");
                }
            }
        }

        private void BW_DoWork(object sender, DoWorkEventArgs e)
        {
            _studentObjs.Clear();
            _classStudents.Clear();

            //建立資料對照
            Dictionary<string, ClassRecord> class_record_dic = new Dictionary<string, ClassRecord>();
            foreach (ClassRecord class_record in _classRecords)
            {
                if (!class_record_dic.ContainsKey(class_record.ID))
                    class_record_dic.Add(class_record.ID, class_record);

                //建立班級及學生ID對照
                if (!_classStudents.ContainsKey(class_record.ID))
                    _classStudents.Add(class_record.ID, new List<string>());

                //遞迴指定班級的學生
                foreach (StudentRecord student in class_record.Students)
                {
                    if (student.Status != StudentRecord.StudentStatus.一般)
                        continue;

                    if (!_classStudents[class_record.ID].Contains(student.ID))
                        _classStudents[class_record.ID].Add(student.ID);

                    //建立學生資料對照
                    if (!_studentObjs.ContainsKey(student.ID))
                        _studentObjs.Add(student.ID, new StudentObj(class_record, student));
                }
            }

            List<SemesterScoreRecord> sems_score_list = K12.Data.SemesterScore.SelectBySchoolYearAndSemester(_studentObjs.Keys, _schoolYear, _semester);
            foreach (SemesterScoreRecord ssr in sems_score_list)
            {
                StudentObj obj = _studentObjs[ssr.RefStudentID];
                obj.AvgScore = ssr.AvgScore.HasValue ? ssr.AvgScore.Value : 0;
                obj.AvgGPA = ssr.AvgGPA.HasValue ? ssr.AvgGPA.Value : 0;
            }

            //Rank
            foreach (string class_id in _classStudents.Keys)
                Rank(class_id);

            //Sort by Rank
            foreach (string cid in _classStudents.Keys)
                _classStudents[cid].Sort(delegate(string x, string y)
                {
                    string xx = (_studentObjs[x].Rank + "").PadLeft(3, '0');
                    xx += (_studentObjs[x].Student.SeatNo + "").PadLeft(3, '0');
                    string yy = (_studentObjs[y].Rank + "").PadLeft(3, '0');
                    yy += (_studentObjs[y].Student.SeatNo + "").PadLeft(3, '0');
                    return xx.CompareTo(yy);
                });

            //Print
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.Top3Template));

            Cells cs = wb.Worksheets[0].Cells;
            cs[0, 0].PutValue("雙語部 " + (_schoolYear + 1911) + "-" + (_schoolYear + 1912) + "年度 第" + _semester + "學期 全班前三名名單");
            cs[1, 0].PutValue("考試別:" + _ExamName);
            cs[1, 3].PutValue("列印日期:" + SelectTime());

            Range titleRange = cs.CreateRange(0, 0, 3, 7);
            Range eachRowRange = cs.CreateRange(3, 0, 1, 7);

            int row_index = 3;
            foreach (string cid in _classStudents.Keys)
            {
                foreach (string sid in _classStudents[cid])
                {
                    StudentObj obj = _studentObjs[sid];

                    if (obj.Rank > 3)
                        continue;

                    if (row_index % 50 == 0)
                    {
                        //wb.Worksheets[0].HorizontalPageBreaks.Add(row_index);
                        cs.CreateRange(row_index, 0, 3, 7).Copy(titleRange);
                        row_index += 3;
                    }

                    cs.CreateRange(row_index, 7, false).CopyStyle(eachRowRange);

                    cs[row_index, 0].PutValue(obj.Class.Name);
                    cs[row_index, 1].PutValue(obj.Student.SeatNo + "");
                    cs[row_index, 2].PutValue(obj.Student.StudentNumber);
                    cs[row_index, 3].PutValue(obj.Student.Name + " " + obj.Student.EnglishName);
                    cs[row_index, 4].PutValue(obj.AvgScore);
                    cs[row_index, 5].PutValue(obj.AvgGPA);
                    cs[row_index, 6].PutValue(obj.Rank);

                    row_index++;
                }
            }

            e.Result = wb;
        }

        private void Rank(string class_id)
        {
            List<StudentObj> score_list = new List<StudentObj>();

            foreach (string student_id in _classStudents[class_id])
                score_list.Add(_studentObjs[student_id]);

            score_list.Sort(delegate(StudentObj x, StudentObj y)
            {
                if (x.AvgGPA == y.AvgGPA)
                    return x.AvgScore.CompareTo(y.AvgScore);
                else
                    return x.AvgGPA.CompareTo(y.AvgGPA);
            });

            score_list.Reverse();

            int rank = 0;
            int count = 0;
            decimal temp_score = decimal.MinValue;
            decimal temp_gpa = decimal.MinValue;
            foreach (StudentObj obj in score_list)
            {
                count++;

                if (temp_gpa != obj.AvgGPA)
                {
                    rank = count;
                }
                else
                {
                    if (temp_score != obj.AvgScore)
                        rank = count;
                }

                obj.Rank = rank;
                temp_score = obj.AvgScore;
                temp_gpa = obj.AvgGPA;
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            _schoolYear = int.Parse(cboSchoolYear.Text);
            _semester = int.Parse(cboSemester.Text);
            _ExamName = "學期成績";

            FormEnable(false);

            if (_BW.IsBusy)
                MessageBox.Show("系統忙碌中,請稍後再試...");
            else
                _BW.RunWorkerAsync();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private string SelectTime() //取得Server的時間
        {
            QueryHelper Sql = new QueryHelper();
            DataTable dtable = Sql.Select("select now()"); //取得時間
            DateTime dt = DateTime.Now;
            DateTime.TryParse("" + dtable.Rows[0][0], out dt); //Parse資料
            string ComputerSendTime = dt.ToString("yyyy/MM/dd"); //最後時間

            return ComputerSendTime;
        }

        private void FormEnable(bool b)
        {
            cboSchoolYear.Enabled = b;
            cboSemester.Enabled = b;
            btnOk.Enabled = b;
        }

        private class StudentObj
        {
            public int Rank;
            public decimal AvgScore, AvgGPA;
            public StudentRecord Student;
            public ClassRecord Class;

            public StudentObj(ClassRecord classRecord, StudentRecord studentRecord)
            {
                Class = classRecord;
                Student = studentRecord;
            }
        }
    }
}
