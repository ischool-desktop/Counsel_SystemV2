using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using K12.Data ;
using FISCA.UDT;
using System.Xml.Linq;

namespace Counsel_System2.Forms
{
    public partial class StudCaseMeetingRecordForm : FISCA.Presentation.Controls.BaseForm
    {
        DAO.UDT_CounselCaseMeetingRecordDef _CounselCaseMeetingRecord;
        DAO.UDTTransfer _UDTTransfer;
        DAO.LogTransfer _LogTransfer;
        public enum accessType { Insert, Edit }
        private accessType _accessType;
        private StudentRecord _studRec;
        Dictionary<int, string> _TeacherIDNameDict;
        Dictionary<string, int> _TeacherNameIDDict;
        string strItem = "Item";
        string strName = "name";
        string strRemark = "remark";
        /// <summary>
        /// 學生個案會議
        /// </summary>
        public StudCaseMeetingRecordForm(DAO.UDT_CounselCaseMeetingRecordDef CounselCaseMeetingRecord,accessType accType)
        {
            InitializeComponent();            
            _UDTTransfer = new DAO.UDTTransfer();
            _CounselCaseMeetingRecord = CounselCaseMeetingRecord;
            _accessType = accType;
            _studRec = Student.SelectByID(CounselCaseMeetingRecord.StudentID.ToString());
            _TeacherIDNameDict = Utility.GetCounselTeacherIDNameDict(CounselCaseMeetingRecord.StudentID.ToString ());
            _TeacherNameIDDict = Utility.GetCounselTeacherNameIDDict(CounselCaseMeetingRecord.StudentID.ToString());
            //if (_accessType == accessType.Insert)
            //    _CounselCaseMeetingRecord.AuthorID = Utility.GetAuthorID();
            _LogTransfer = new DAO.LogTransfer();            
        }

        private void StudCaseMeetingRecordForm_Load(object sender, EventArgs e)
        {
            LoadDefaultData();
            LoadUDTDataToForm();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (SaveFormDataToUDTDef() == false)
            {
                Utility.ShowCannotSaveMessage();
                return;
            }

            // log
            LogData();

            string studStr = "學號:" + _studRec.StudentNumber + ",姓名:" + _studRec.Name + ",";

            if (_accessType == accessType.Insert)
            {
                // 檢查是否可以新增 會議日期+會議事由 不能重複
                List<DAO.UDT_CounselCaseMeetingRecordDef> dataList = _UDTTransfer.GetCaseMeetingRecordListByStudentID(_studRec.ID);
                bool pass=true;
                foreach (DAO.UDT_CounselCaseMeetingRecordDef data in dataList)
                {
                    if (data.MeetingDate.HasValue && _CounselCaseMeetingRecord.MeetingDate.HasValue)
                        if (data.MeetingDate.Value.ToShortDateString() == _CounselCaseMeetingRecord.MeetingDate.Value.ToShortDateString())
                            if (data.MeetingCause == _CounselCaseMeetingRecord.MeetingCause)
                                pass = false;
                }

                if (pass)
                {
                    // log
                    _LogTransfer.SaveInsertLog("學生.個案會議-新增", "新增", studStr, "", "student", _studRec.ID);
                    _UDTTransfer.InsertCaseMeetingRecord(_CounselCaseMeetingRecord);
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("已有相同的會議日與會議事由，無法新增。");
                    return;
                }
            }
            else
            {
                // log
                _LogTransfer.SaveChangeLog("學生.個案會議-修改", "修改", studStr, "", "student", _studRec.ID);

                _UDTTransfer.UpdateCaseMeetingRecord(_CounselCaseMeetingRecord);
            }
            Utility.ShowSavedMessage();
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        private void LoadDefaultData()
        {
            cbxInterviewer.Items.AddRange(_TeacherNameIDDict.Keys.ToArray());
            
        
        }

        private void LoadUDTDataToForm()
        {
            txtCaseNo.Text = _CounselCaseMeetingRecord.CaseNo;
            if (_studRec.Class != null)
            {
                if (_studRec.Class.GradeYear.HasValue)
                    lblGradeYear.Text = _studRec.Class.GradeYear.Value.ToString();

                lblClassName.Text = _studRec.Class.Name;
            }
            lblName.Text = _studRec.Name;
            if (_TeacherIDNameDict.ContainsKey(_CounselCaseMeetingRecord.CounselTeacherID))
                cbxInterviewer.Text = _TeacherIDNameDict[_CounselCaseMeetingRecord.CounselTeacherID];
           
            txtPlace.Text = _CounselCaseMeetingRecord.Place;
            if(_CounselCaseMeetingRecord.MeetingDate.HasValue )
                dtDate.Value  = _CounselCaseMeetingRecord.MeetingDate.Value;

            txtTime.Text  = _CounselCaseMeetingRecord.MeetigTime;
                        
            txtMeetingCause.Text = _CounselCaseMeetingRecord.MeetingCause;

            StringBuilder sb1 = new StringBuilder();
            sb1.Append("<root>"); sb1.Append(_CounselCaseMeetingRecord.Attendees); sb1.Append("</root>");
            // 解析 ContentXML
            XElement xmlAttendees = XElement.Parse(sb1.ToString());
            if (xmlAttendees != null)
            {
                // 參與人員 ---
                foreach (XElement elm in xmlAttendees.Elements(strItem))
                {
                    if (elm.Attribute(strName) == null)
                        continue;

                    switch (elm.Attribute(strName).Value.ToString())
                    {
                        // 學生
                        case "學生":
                            cb001.Checked = true;
                            break;
                        // 家長
                        case "家長":
                            cb002.Checked = true;
                            break;

                        // 專家
                        case "專家":
                            cb003.Checked = true;
                            break;

                        // 醫師
                        case "醫師":
                            cb004.Checked = true;
                            break;

                        // 社工人員
                        case "社工人員":
                            cb005.Checked = true;
                            break;

                        // 導師
                        case "導師":
                            cb006.Checked = true;
                            break;

                        // 教官
                        case "教官":
                            cb007.Checked = true;
                            break;

                        // 輔導老師
                        case "輔導老師":
                            cb008.Checked = true;
                            break;

                        // 任課老師
                        case "任課老師":
                            cb009.Checked = true;
                            break;

                        // 其它**
                        case "其它":
                            cb010.Checked = true;
                            if (elm.Attribute(strRemark) != null)
                                txt_cb010.Text = elm.Attribute(strRemark).Value.ToString();
                            break;

                    }

                }
            }

            StringBuilder sb2 = new StringBuilder();
            sb2.Append("<root>"); sb2.Append(_CounselCaseMeetingRecord.CounselType); sb2.Append("</root>");
            XElement xmlCounselType = XElement.Parse(sb2.ToString());
            // 輔導方式--
            if (xmlCounselType != null)
            {
                foreach (XElement elm in xmlCounselType.Elements(strItem))
                {
                    if (elm.Attribute(strName) == null)
                        continue;

                    switch (elm.Attribute(strName).Value.ToString())
                    {
                        // 暫時結案
                        case "暫時結案":
                            cb101.Checked = true;
                            break;

                        // 專案輔導
                        case "專案輔導":
                            cb102.Checked = true;
                            break;
                        // 導師輔導
                        case "導師輔導":
                            cb103.Checked = true;
                            break;
                        // 轉介**
                        case "轉介":
                            cb104.Checked = true;
                            if (elm.Attribute(strRemark) != null)
                                txt_cb104.Text = elm.Attribute(strRemark).Value.ToString();
                            break;
                        // 就醫**
                        case "就醫":
                            cb105.Checked = true;
                            if (elm.Attribute(strRemark) != null)
                                txt_cb105.Text = elm.Attribute(strRemark).Value.ToString();
                            break;
                        // 其它**
                        case "其它":
                            cb106.Checked = true;
                            if (elm.Attribute(strRemark) != null)
                                txt_cb106.Text = elm.Attribute(strRemark).Value.ToString();
                            break;
                    }
                }
            }

            StringBuilder sb3 = new StringBuilder();
            sb3.Append("<root>"); sb3.Append(_CounselCaseMeetingRecord.CounselTypeKind); sb3.Append("</root>");
            XElement xmlCounselTypeKind = XElement.Parse(sb3.ToString());

            // 輔導歸類--
            if (xmlCounselTypeKind != null)
            {
                foreach (XElement elm in xmlCounselTypeKind.Elements(strItem))
                {
                    if (elm.Attribute(strName) == null)
                        continue;

                    switch (elm.Attribute(strName).Value.ToString())
                    {
                        // 違規
                        case "違規":
                            cb201.Checked = true;
                            break;

                        // 遲曠
                        case "遲曠":
                            cb202.Checked = true;
                            break;

                        // 學習
                        case "學習":
                            cb203.Checked = true;
                            break;

                        // 生涯
                        case "生涯":
                            cb204.Checked = true;
                            break;

                        // 人
                        case "人":
                            cb205.Checked = true;
                            break;

                        // 休退轉
                        case "休退轉":
                            cb206.Checked = true;
                            break;

                        // 家庭
                        case "家庭":
                            cb207.Checked = true;
                            break;

                        // 師生
                        case "師生":
                            cb208.Checked = true;
                            break;

                        // 情感
                        case "情感":
                            cb209.Checked = true;
                            break;

                        // 精神
                        case "精神":
                            cb210.Checked = true;
                            break;

                        // 其它**
                        case "其它":
                            cb211.Checked = true;
                            if (elm.Attribute(strRemark) != null)
                                txt_cb211.Text = elm.Attribute(strRemark).Value.ToString();
                            break;

                    }
                }
            }

            txtContentDigest.Text = _CounselCaseMeetingRecord.ContentDigest;
            txtAuthor_id.Text = _CounselCaseMeetingRecord.AuthorID;
            txtAuthorName.Text = _CounselCaseMeetingRecord.AuthorName;

            //log

            _LogTransfer.Clear();
            LogData();
         
            if(string.IsNullOrEmpty(txtAuthor_id.Text))
                txtAuthor_id.Text= Utility.GetAuthorID();
        }

        private void LogData()
        {

            _LogTransfer.SetLogValue("個案編號", txtCaseNo.Text );
            _LogTransfer.SetLogValue("晤談老師", cbxInterviewer.Text);            
            _LogTransfer.SetLogValue("會議日期", dtDate.Text);
            _LogTransfer.SetLogValue("會議時間", txtTime.Text);
            _LogTransfer.SetLogValue("會議事由", txtMeetingCause.Text);
            _LogTransfer.SetLogValue("會議地點", txtPlace.Text);

            // 參與人員
            List<string> info1 = new List<string>();
            foreach (Control cr in groupPanel1.Controls)
            {
                CheckBox cb = cr as CheckBox;
                if (cb != null)
                {
                    if (cb.Checked)
                    {
                        if (cb.Text.IndexOf("其它") > -1)
                            info1.Add("其它:" + txt_cb010.Text);
                        else
                            info1.Add(cb.Text);
                    }                
                }            
            }
            _LogTransfer.SetLogValue("參與人員:", string.Join(",",info1.ToArray()));

            // 輔導方式
            List<string> info2 = new List<string>();
            foreach (Control cr in groupPanel2.Controls)
            {
                CheckBox cb = cr as CheckBox;
                if (cb != null)
                {
                    if (cb.Checked)
                    {
                        if (cb.Text.IndexOf("轉介") > -1)
                            info2.Add("轉介:" + txt_cb104.Text);
                        else if (cb.Text.IndexOf("就醫") > -1)
                            info2.Add("就醫:" + txt_cb105.Text);
                        else if(cb.Text.IndexOf("其它") > -1)
                            info2.Add("其它:" + txt_cb106.Text);
                        else
                            info2.Add(cb.Text);
                    }                
                }            
            }
            _LogTransfer.SetLogValue("輔導方式:", string.Join(",", info2.ToArray()));

            // 輔導歸類
            List<string> info3 = new List<string>();
            foreach (Control cr in groupPanel3.Controls)
            {
                CheckBox cb = cr as CheckBox;
                if (cb != null)
                {
                    if (cb.Checked)
                    {
                        if (cb.Text.IndexOf("其它") > -1)
                            info3.Add("其它:" + txt_cb211.Text);
                        else
                            info3.Add(cb.Text);
                    }
                }
            }
            _LogTransfer.SetLogValue("輔導歸類:", string.Join(",", info3.ToArray()));
            _LogTransfer.SetLogValue("內容要點", txtContentDigest.Text);
            _LogTransfer.SetLogValue("記錄者", _CounselCaseMeetingRecord.AuthorID);
            _LogTransfer.SetLogValue("記錄者姓名", _CounselCaseMeetingRecord.AuthorName);
        }

        private bool SaveFormDataToUDTDef()
        {
            bool checkPass = true;
            _CounselCaseMeetingRecord.CaseNo = txtCaseNo.Text;
            if (_TeacherNameIDDict.ContainsKey(cbxInterviewer.Text))
                _CounselCaseMeetingRecord.CounselTeacherID = _TeacherNameIDDict[cbxInterviewer.Text];
           

            if (string.IsNullOrEmpty(dtDate.Text))
                checkPass = false;

            // 日期            
            DateTime? newDateTime = null;
            newDateTime = dtDate.Value;
            _CounselCaseMeetingRecord.MeetigTime = txtTime.Text;
            _CounselCaseMeetingRecord.MeetingDate = dtDate.Value;
            
            _CounselCaseMeetingRecord.Place = txtPlace.Text;
            _CounselCaseMeetingRecord.MeetingCause = txtMeetingCause.Text;

            // 參與人員 --
            StringBuilder sb1 = new StringBuilder();
            // 學生
            if (cb001.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "學生");
                sb1.Append(se1.ToString());
            }

            // 家長
            if (cb002.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "家長");
                sb1.Append(se1.ToString());
            }
            // 專家
            if (cb003.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "專家");
                sb1.Append(se1.ToString());
            }
            // 醫師
            if (cb004.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "醫師");
                sb1.Append(se1.ToString());
            }
            // 社工人員
            if (cb005.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "社工人員");
                sb1.Append(se1.ToString());
            }
            // 導師
            if (cb006.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "導師");
                sb1.Append(se1.ToString());
            }
            // 教官
            if (cb007.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "教官");
                sb1.Append(se1.ToString());
            }
            // 輔導老師
            if (cb008.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "輔導老師");
                sb1.Append(se1.ToString());
            }
            // 任課老師
            if (cb009.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "任課老師");
                sb1.Append(se1.ToString());
            }
            // 其它**
            if (cb010.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "其它");
                se1.SetAttributeValue(strRemark, txt_cb010.Text);
                sb1.Append(se1.ToString());
            }

            StringBuilder sb2 = new StringBuilder();
            if (cb101.Checked)
            {
                XElement se2 = new XElement(strItem);
                se2.SetAttributeValue(strName, "暫時結案");
                sb2.Append(se2.ToString());
            }
            // 專案輔導
            if (cb102.Checked)
            {
                XElement se2 = new XElement(strItem);
                se2.SetAttributeValue(strName, "專案輔導");
                sb2.Append(se2.ToString());
            }

            // 導師輔導
            if (cb103.Checked)
            {
                XElement se2 = new XElement(strItem);
                se2.SetAttributeValue(strName, "導師輔導");
                sb2.Append(se2.ToString());
            }

            // 轉介**
            if (cb104.Checked)
            {
                XElement se2 = new XElement(strItem);
                se2.SetAttributeValue(strName, "轉介");
                se2.SetAttributeValue(strRemark, txt_cb104.Text);
                sb2.Append(se2.ToString());
            }

            // 就醫**
            if (cb105.Checked)
            {
                XElement se2 = new XElement(strItem);
                se2.SetAttributeValue(strName, "就醫");
                se2.SetAttributeValue(strRemark, txt_cb105.Text);
                sb2.Append(se2.ToString());
            }

            // 其它**
            if (cb106.Checked)
            {
                XElement se2 = new XElement(strItem);
                se2.SetAttributeValue(strName, "其它");
                se2.SetAttributeValue(strRemark, txt_cb106.Text);
                sb2.Append(se2.ToString());
            }


            // 輔導歸類 --
            StringBuilder sb3 = new StringBuilder();

            // 違規
            if (cb201.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "違規");
                sb3.Append(se3.ToString());
            }
            // 遲曠
            if (cb202.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "遲曠");
                sb3.Append(se3.ToString());
            }
            // 學習
            if (cb203.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "學習");
                sb3.Append(se3.ToString());
            }
            // 生涯
            if (cb204.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "生涯");
                sb3.Append(se3.ToString());
            }
            // 人
            if (cb205.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "人");
                sb3.Append(se3.ToString());
            }
            // 休退轉
            if (cb206.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "休退轉");
                sb3.Append(se3.ToString());
            }
            // 家庭
            if (cb207.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "家庭");
                sb3.Append(se3.ToString());
            }
            // 師生
            if (cb208.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "師生");
                sb3.Append(se3.ToString());
            }
            // 情感
            if (cb209.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "情感");
                sb3.Append(se3.ToString());
            }
            // 精神
            if (cb210.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "精神");
                sb3.Append(se3.ToString());
            }
            // 其它**
            if (cb211.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "其它");
                se3.SetAttributeValue(strRemark, txt_cb211.Text);
                sb3.Append(se3.ToString());
            }

            if (sb1.ToString().Trim() != "")
                _CounselCaseMeetingRecord.Attendees = sb1.ToString();
            else
                checkPass = false;


            if (sb2.ToString().Trim() != "")
                _CounselCaseMeetingRecord.CounselType = sb2.ToString();
            else
                checkPass = false;

            if (sb3.ToString().Trim() != "")
                _CounselCaseMeetingRecord.CounselTypeKind = sb3.ToString();
            else
                checkPass = false;

            _CounselCaseMeetingRecord.ContentDigest = txtContentDigest.Text;
            _CounselCaseMeetingRecord.AuthorID = txtAuthor_id.Text;
            _CounselCaseMeetingRecord.AuthorName = txtAuthorName.Text;

            return checkPass;
        }

    }
}
