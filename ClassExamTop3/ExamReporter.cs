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
    public partial class ExamReporter : BaseForm
    {
        private int _schoolYear, _semester, _examid;
        private string _ExamName;
        private BackgroundWorker _BW;
        private QueryHelper _Q;
        private Dictionary<string, StudentObj> _studentObjs;
        private Dictionary<string, List<string>> _classStudents;
        private List<ClassRecord> _classRecords;

        public ExamReporter()
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

            //Set cboExamType Item
            cboExamType.Items.Add("Midterm");
            cboExamType.Items.Add("Final");

            cboSchoolYear.Text = _schoolYear + "";
            cboSemester.Text = _semester + "";
            cboExamType.Text = "Midterm";
        }

        private void BW_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            FormEnable(true);
            Workbook wb = e.Result as Workbook;
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "另存新檔";
            save.FileName = "全班前三名名單(評量成績).xls";
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

            //取得成績資料
            string student_ID = string.Join(",", _studentObjs.Keys);
            string sql = "";
            if (!string.IsNullOrWhiteSpace(student_ID))
            {
                //組SQL語法抓取學生ID、科目、成績
                sql = "select sc_attend.ref_student_id,course.subject,course.credit,$ischool.subject.list.type,xpath_string(sce_take.extension,'//Extension/Score') as score from sc_attend";
                sql += " join sce_take on sce_take.ref_sc_attend_id=sc_attend.id";
                sql += " join course on course.id=sc_attend.ref_course_id";
                sql += " join $ischool.subject.list on $ischool.subject.list.name=course.subject";
                sql += " where sc_attend.ref_student_id in (" + student_ID + ") and sce_take.ref_exam_id=" + _examid + " and course.school_year=" + _schoolYear + " and course.semester=" + _semester;
            }

            if (!string.IsNullOrWhiteSpace(sql))
            {
                DataTable dt = _Q.Select(sql);
                foreach (DataRow row in dt.Rows)
                {
                    string student_id = row["ref_student_id"] + "";

                    _studentObjs[student_id].LoadData(row);
                }
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
            cs[0, 0].PutValue("雙語部 " + _schoolYear + "年度 第" + _semester + "學期 全班前三名名單");
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
                if (x.AvgScore == y.AvgScore)
                    return x.AvgGPA.CompareTo(y.AvgGPA);
                else
                    return x.AvgScore.CompareTo(y.AvgScore);
            });

            score_list.Reverse();

            int rank = 0;
            int count = 0;
            decimal temp_score = decimal.MinValue;
            decimal temp_gpa = decimal.MinValue;
            foreach (StudentObj obj in score_list)
            {
                count++;

                if (temp_score != obj.AvgScore)
                {
                    rank = count;
                }
                else
                {
                    if (temp_gpa != obj.AvgGPA)
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
            _ExamName = cboExamType.Text;

            if (cboExamType.Text == "Midterm")
                _examid = 1;
            else
                _examid = 2;

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
            cboExamType.Enabled = b;
            btnOk.Enabled = b;
        }

        private class StudentObj
        {
            public int Rank;
            public Dictionary<string, decimal> SubjectCredit;
            public Dictionary<string, decimal> SubjectScore;
            public Dictionary<string, decimal> SubjectGPA;
            public StudentRecord Student;
            public ClassRecord Class;

            public StudentObj(ClassRecord classRecord, StudentRecord studentRecord)
            {
                SubjectCredit = new Dictionary<string, decimal>();
                SubjectScore = new Dictionary<string, decimal>();
                SubjectGPA = new Dictionary<string, decimal>();

                Class = classRecord;
                Student = studentRecord;
            }

            public void SetSubjectScore(string subject, string type, string str_credit, string str_score)
            {
                decimal credit_d = 0;
                decimal credit = decimal.TryParse(str_credit, out credit_d) ? credit_d : 0;

                decimal score_d = 0;
                decimal score = decimal.TryParse(str_score, out score_d) ? score_d : 0;

                decimal gpa = 0;
                if (type == "Honor")
                    gpa = Tool.GPA.Eval(score).Honors;
                else
                    gpa = Tool.GPA.Eval(score).Regular;

                if (!SubjectCredit.ContainsKey(subject))
                    SubjectCredit.Add(subject, credit);

                if (!SubjectScore.ContainsKey(subject))
                    SubjectScore.Add(subject, score);

                if (!SubjectGPA.ContainsKey(subject))
                    SubjectGPA.Add(subject, gpa);
            }

            public void LoadData(DataRow row)
            {
                string subject = row["subject"] + "";
                string type = row["type"] + "";
                string credit = row["credit"] + "";
                string score = row["score"] + "";

                SetSubjectScore(subject, type, credit, score);
            }

            public decimal AvgScore
            {
                get
                {
                    decimal count = 0;
                    decimal total = 0;
                    foreach (string subj in SubjectScore.Keys)
                    {
                        decimal score = SubjectScore[subj];
                        decimal credit = SubjectCredit[subj];

                        count += credit;
                        total += score * credit;
                    }

                    //Subject_Credit.Values.Select(x => x * x).ToList().Sum();
                    if (count > 0)
                        return Math.Round(total / count, 2, MidpointRounding.AwayFromZero);
                    else
                        return 0;
                }
            }

            public decimal AvgGPA
            {
                get
                {
                    decimal count = 0;
                    decimal total = 0;
                    foreach (string subj in SubjectGPA.Keys)
                    {
                        decimal gpa = SubjectGPA[subj];
                        decimal credit = SubjectCredit[subj];

                        count += credit;
                        total += gpa * credit;
                    }

                    if (count > 0)
                        return Math.Round(total / count, 2, MidpointRounding.AwayFromZero);
                    else
                        return 0;
                }
            }
        }
    }
}
