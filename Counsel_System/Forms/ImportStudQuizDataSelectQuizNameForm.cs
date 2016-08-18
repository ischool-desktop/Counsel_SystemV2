using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using K12.Data;

namespace Counsel_System2.Forms
{
    public partial class ImportStudQuizDataSelectQuizNameForm : FISCA.Presentation.Controls.BaseForm
    {
        DAO.UDTTransfer _UDTTransfer;
        string _ConfigData_Name = "輔導系統_匯入測驗選擇方式";
        string _ConfigData_ItemSnum = "學號";
        string _ConfigData_ItemSeatNo = "班級座號";

        List<DAO.UDT_QuizDef> _QuizDataList = new List<DAO.UDT_QuizDef>();
        public ImportStudQuizDataSelectQuizNameForm()
        {
            InitializeComponent();
            this.MaximumSize = this.MinimumSize = this.Size;
            _UDTTransfer = new DAO.UDTTransfer();            
            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(cbxQuizName.Text))
            {
                FISCA.Presentation.Controls.MsgBox.Show("請選擇匯入測驗名稱.");
                return;
            }

            // 儲存畫面選項設定
            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration[_ConfigData_Name];
            cd[_ConfigData_ItemSnum] = rbStudentNumber.Checked.ToString();
            cd[_ConfigData_ItemSeatNo] = rbClassSeatNo.Checked.ToString();
            cd.Save();

            foreach(DAO.UDT_QuizDef qd in _QuizDataList.Where(x=>x.QuizName==cbxQuizName.Text).ToList ())
            {
                // 動態 XML
                XElement elmRoot=null;

                // 當選學號需要動態加入
                if (rbStudentNumber.Checked)
                    elmRoot = XElement.Parse(Properties.Resources.ImportStudQuizDataVal_SNum);
                

                // 當選班級座號需要動態加入
                if (rbClassSeatNo.Checked)
                    elmRoot = XElement.Parse(Properties.Resources.ImportStudQuizDataVa_CSeatNo);


                foreach (XElement elm1 in Utility.ConvertStringToXelm1(qd.QuizDataField).Elements("Field"))
                {
                    XElement elm = new XElement("Field");
                    elm.SetAttributeValue("Required", "True");
                    elm.SetAttributeValue("Name", elm1.Attribute("name").Value);
                    elmRoot.Element("FieldList").Add(elm);
                }
                Global._ImportStudQuizDataValElement = elmRoot;
                //// 取得該試別所有測驗
                //List<DAO.UDT_StudQuizDataDef> studQuizList = _UDTTransfer.GetStudQuizDataByQuizID(qd.UID);

                //List<string> sidlist = (from data in studQuizList select data.StudentID.ToString()).ToList ();


                //foreach (K12.Data.StudentRecord stud in K12.Data.Student.SelectByIDs(sidlist))
                //{
                //    if (!Global._HasStudQuizDataDictTemp.ContainsKey(stud.StudentNumber))
                //        Global._HasStudQuizDataDictTemp.Add(stud.StudentNumber, stud.ID);
                //}
                // 依測驗ID取得已有資料的學生學號
                Global._HasStudQuizDataDictTemp = Utility.GetHasCounselStudQuizDataStudNumberByQuizID(qd.UID);


                ImportExport.ImportStudQuizData isqd = new ImportExport.ImportStudQuizData();
                isqd.SetQuizID(qd.UID);
                isqd.SetQuizName(qd.QuizName);
                List<string> dataFieldList = (from data in Utility.ConvertStringToXelm1(qd.QuizDataField).Elements("Field") select data.Attribute("name").Value).ToList();
                isqd.SetDataFieldNameList(dataFieldList);
                isqd.Execute();
            }
            this.Close();
        }

        private void ImportStudQuizDataSelectQuizNameForm_Load(object sender, EventArgs e)
        {
            // 預設用學號
            rbClassSeatNo.Checked = false;
            rbStudentNumber.Checked = true;
            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration[_ConfigData_Name];
            
            bool b1,b2;
            if (bool.TryParse(cd[_ConfigData_ItemSnum],out b1))
                rbStudentNumber.Checked = b1;

            if (bool.TryParse(cd[_ConfigData_ItemSeatNo],out b2))
                rbClassSeatNo.Checked = b2;

            _QuizDataList = _UDTTransfer.GetAllQuizData();
            List<string> NameList = _QuizDataList.Select(x => x.QuizName).ToList();
            cbxQuizName.Items.AddRange(NameList.ToArray());
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
