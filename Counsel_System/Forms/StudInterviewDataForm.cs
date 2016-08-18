using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using K12.Data;
using System.Xml.Linq;
using DevComponents.DotNetBar;

namespace Counsel_System2.Forms
{
    public partial class StudInterviewDataForm : FISCA.Presentation.Controls.BaseForm 
    {
        /// <summary>
        /// 晤談紀錄
        /// </summary>
        DAO.UDT_CounselStudentInterviewRecordDef _StudInterviewRecord;
        DAO.UDTTransfer _UDTTransfer;
        Dictionary<int, string> _TeacherIDNameDict;
        Dictionary<string, int> _TeacherNameIDDict;
        public enum AccessType {Insert,Update }
        private StudentRecord _studRec;
        List<TeacherRecord> _TeacherRecList;
        // XML 內用的固定字串        
        string strItem = "Item";
        string strName = "name";
        string strRemark = "remark";
        DAO.LogTransfer _LogTransfer;

        private AccessType _accessType;

        public StudInterviewDataForm(DAO.UDT_CounselStudentInterviewRecordDef StudInterviewRecord,AccessType accType)
        {
            InitializeComponent();
            _StudInterviewRecord = StudInterviewRecord;
            _UDTTransfer = new DAO.UDTTransfer();
            _TeacherIDNameDict = new Dictionary<int, string>();
            _TeacherNameIDDict = new Dictionary<string, int>();
            _accessType = accType;
            _LogTransfer = new DAO.LogTransfer();
            //if (accType == AccessType.Insert)
            //{
            //    _StudInterviewRecord.AuthorID = FISCA.Authentication.DSAServices.UserAccount;
            //    _StudInterviewRecord.isPublic = true;
            //}

            _TeacherRecList = new List<TeacherRecord>();
            _studRec = Student.SelectByID(StudInterviewRecord.StudentID.ToString ());
            _TeacherIDNameDict = Utility.GetCounselTeacherIDNameDict(StudInterviewRecord.StudentID.ToString ());
            _TeacherNameIDDict = Utility.GetCounselTeacherNameIDDict(StudInterviewRecord.StudentID.ToString());
            LoadDefaultData();
            LoadDataToForm();
        }

        private void LoadDefaultData()
        {           
            // 載入預設選單資料
            List<string> TeacherNameList = _TeacherNameIDDict.Keys.ToList();
            TeacherNameList.Sort();
        
            // 加入對照資料
            // 晤談方式
            cbxInterviewType.Items.Clear();
            string strInterviewTypeItems = "面談,電話,家訪,電子信箱,聯絡簿,其它";
            cbxInterviewType.Items.AddRange(strInterviewTypeItems.Split(',').ToArray());
            
            // 晤談對象
            cbxInterveweeType.Items.Clear();
            string strInterveweeTypeItems = "學生,家長,其它";
            cbxInterveweeType.Items.AddRange(strInterveweeTypeItems.Split(',').ToArray());
            
            // 晤談老師
            cbxInterviewer.Items.Clear();
            cbxInterviewer.Items.AddRange(TeacherNameList.ToArray());

        }


        private void LoadDataToForm()
        {
            // 晤談編號
            txtInterviewNo.Text = _StudInterviewRecord.InterviewNo;

            if (_studRec != null)
            {
                if (_studRec.Class != null)
                {
                    // 年級
                    if (_studRec.Class.GradeYear.HasValue)
                        lblGradeYear.Text = _studRec.Class.GradeYear.Value.ToString();

                    // 班級
                    lblClassName.Text = _studRec.Class.Name;
                }
                // 姓名
                lblName.Text = _studRec.Name;

            }
            // 晤談老師
            if(_TeacherIDNameDict.ContainsKey(_StudInterviewRecord.TeacherID))
                cbxInterviewer.Text = _TeacherIDNameDict[_StudInterviewRecord.TeacherID];
                     

            // 晤談方式
            cbxInterviewType.Text = _StudInterviewRecord.InterviewType;

            // 晤談對象
            cbxInterveweeType.Text = _StudInterviewRecord.IntervieweeType;

            if (_StudInterviewRecord.InterviewDate.HasValue)
                dtDate.Value = _StudInterviewRecord.InterviewDate.Value;

            txtTime.Text = _StudInterviewRecord.InterviewTime;
            // 地點
            txtPlace.Text = _StudInterviewRecord.Place;
            // 晤談事由
            txtCause.Text = _StudInterviewRecord.Cause;

            StringBuilder sb1 = new StringBuilder();
            sb1.Append("<root>"); sb1.Append(_StudInterviewRecord.Attendees); sb1.Append("</root>");
            // 解析 ContentXML
            XElement xmlAttendees = XElement.Parse(sb1.ToString());
            if (xmlAttendees != null)
            {
                // 參與人員 ---
                    foreach (XElement elm in xmlAttendees.Elements(strItem))
                    {
                        if (elm.Attribute(strName) == null)
                            continue;
                                                
                        switch (elm.Attribute(strName).Value.ToString ())
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
            
            StringBuilder sb2 = new StringBuilder ();
            sb2.Append("<root>");sb2.Append(_StudInterviewRecord.CounselType);sb2.Append("</root>");
            XElement xmlCounselType=XElement.Parse(sb2.ToString());
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
                
            StringBuilder sb3 = new StringBuilder ();
            sb3.Append("<root>");sb3.Append(_StudInterviewRecord.CounselTypeKind);sb3.Append("</root>");
            XElement xmlCounselTypeKind=XElement.Parse(sb3.ToString ());

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

                            // 人際
                            case "人際":
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

                            // 家暴
                            case "家暴":
                                cb211.Checked = true;
                                break;

                            // 霸凌
                            case "霸凌":
                                cb212.Checked = true;
                                break;

                            // 中輟
                            case "中輟":
                                cb213.Checked = true;
                                break;

                            // 性議題
                            case "性議題":
                                cb214.Checked = true;
                                break;

                            // 戒毒
                            case "戒毒":
                                cb215.Checked = true;
                                break;

                            // 網路成癮
                            case "網路成癮":
                                cb216.Checked = true;
                                break;

                            // 情緒障礙
                            case "情緒障礙":
                                cb217.Checked = true;
                                break;

                            // 其它**
                            case "其它":
                                cb218.Checked = true;
                                if (elm.Attribute(strRemark) != null)
                                    txt_cb218.Text = elm.Attribute(strRemark).Value.ToString();
                                break;

                        }
                    }
                }
            
            // 內容要點
            txtContentDigest.Text = _StudInterviewRecord.ContentDigest;
            //chkIsPublic.Checked = _StudInterviewRecord.isPublic;
            txtAuthorID.Text = _StudInterviewRecord.AuthorID;
            // 紀錄者姓名˙
            txtAuthorName.Text = _StudInterviewRecord.AuthorName;
    
            // Log
            _LogTransfer.Clear();
            LogData();

            // 當記錄者是空白加入系統使用者
            if(string.IsNullOrEmpty(txtAuthorID.Text))
                txtAuthorID.Text = FISCA.Authentication.DSAServices.UserAccount;
        }

        private void LogData()
        {
            _LogTransfer.SetLogValue("晤談編號", txtInterviewNo.Text);
            _LogTransfer.SetLogValue("晤談老師", cbxInterviewer.Text);
            _LogTransfer.SetLogValue("晤談對象", cbxInterveweeType.Text);
            _LogTransfer.SetLogValue("晤談方式", cbxInterviewType.Text);
            _LogTransfer.SetLogValue("晤談日期", dtDate.Text);
            _LogTransfer.SetLogValue("晤談時間", txtTime.Text);
            _LogTransfer.SetLogValue("晤談地點", txtPlace.Text);
            _LogTransfer.SetLogValue("晤談事由", txtCause.Text);

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
            _LogTransfer.SetLogValue("參與人員:", string.Join(",", info1.ToArray()));

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
                        else if (cb.Text.IndexOf("其它") > -1)
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
                            info3.Add("其它:" + txt_cb218.Text);
                        else
                            info3.Add(cb.Text);
                    }
                }
            }
            _LogTransfer.SetLogValue("輔導歸類:", string.Join(",", info3.ToArray()));
            _LogTransfer.SetLogValue("內容要點", txtContentDigest.Text);
            _LogTransfer.SetLogValue("記錄者", txtAuthorID.Text);
            _LogTransfer.SetLogValue("記錄者姓名", txtAuthorName.Text);
        }

        /// <summary>
        /// 寫入資料
        /// </summary>
        private bool SaveData()
        {
            bool checkPass = true;

            // 晤談編號
            _StudInterviewRecord.InterviewNo = txtInterviewNo.Text;
            // 晤談老師
            if(_TeacherNameIDDict.ContainsKey(cbxInterviewer.Text ))
                _StudInterviewRecord.TeacherID = _TeacherNameIDDict[cbxInterviewer.Text];
           

            // 晤談方式
            _StudInterviewRecord.InterviewType = cbxInterviewType.Text;
            // 晤談對象
            _StudInterviewRecord.IntervieweeType = cbxInterveweeType.Text;

            if (string.IsNullOrEmpty(cbxInterviewer.Text) || string.IsNullOrEmpty(cbxInterveweeType.Text) || string.IsNullOrEmpty(txtCause.Text) || string.IsNullOrEmpty(cbxInterviewType.Text))
                checkPass = false;

            if (string.IsNullOrEmpty(dtDate.Text))
                checkPass = false;

            _StudInterviewRecord.InterviewDate = dtDate.Value;
            _StudInterviewRecord.InterviewTime = txtTime.Text;

            // 地點
            _StudInterviewRecord.Place = txtPlace.Text;
            // 晤談事由
            _StudInterviewRecord.Cause = txtCause.Text;
                      

            // 參與人員 --
            StringBuilder sb1 = new StringBuilder();
            // 學生
            if (cb001.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "學生");
                sb1.Append(se1.ToString ());
            }
            
            // 家長
            if (cb002.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "家長");
                 sb1.Append(se1.ToString ());
            }
            // 專家
            if (cb003.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "專家");
                 sb1.Append(se1.ToString ());
            }
            // 醫師
            if (cb004.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "醫師");
                 sb1.Append(se1.ToString ());
            }
            // 社工人員
            if (cb005.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "社工人員");
                 sb1.Append(se1.ToString ());
            }
            // 導師
            if (cb006.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "導師");
                 sb1.Append(se1.ToString ());
            }
            // 教官
            if (cb007.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "教官");
                 sb1.Append(se1.ToString ());
            }
            // 輔導老師
            if (cb008.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "輔導老師");
                 sb1.Append(se1.ToString ());
            }
            // 任課老師
            if (cb009.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "任課老師");
                 sb1.Append(se1.ToString ());
            }
            // 其它**
            if (cb010.Checked)
            {
                XElement se1 = new XElement(strItem);
                se1.SetAttributeValue(strName, "其它");
                se1.SetAttributeValue(strRemark, txt_cb010.Text);
                 sb1.Append(se1.ToString ());
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
            // 人際
            if (cb205.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "人際");
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
            // 家暴
            if (cb211.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "家暴");
                sb3.Append(se3.ToString());
            }
            // 霸凌
            if (cb212.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "霸凌");
                sb3.Append(se3.ToString());
            }
            // 中輟
            if (cb213.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "中輟");
                sb3.Append(se3.ToString());
            }
            // 性議題
            if (cb214.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "性議題");
                sb3.Append(se3.ToString());
            }
            // 戒毒
            if (cb215.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "戒毒");
                sb3.Append(se3.ToString());
            }
            // 網路成癮
            if (cb216.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "網路成癮");
                sb3.Append(se3.ToString());
            }
            // 情緒障礙
            if (cb217.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "情緒障礙");
                sb3.Append(se3.ToString());
            }

            // 其它**
            if (cb218.Checked)
            {
                XElement se3 = new XElement(strItem);
                se3.SetAttributeValue(strName, "其它");
                se3.SetAttributeValue(strRemark, txt_cb218.Text);
               sb3.Append(se3.ToString());
            }

            if (sb1.ToString().Trim() != "")
                _StudInterviewRecord.Attendees = sb1.ToString();
            else
                checkPass = false;


            if (sb2.ToString().Trim() != "")
                _StudInterviewRecord.CounselType = sb2.ToString();
            else
                checkPass = false;

            if (sb3.ToString().Trim() != "")
                _StudInterviewRecord.CounselTypeKind = sb3.ToString();
            else
                checkPass = false;
            
            // 內容要點
            _StudInterviewRecord.ContentDigest = txtContentDigest.Text;
            //_StudInterviewRecord.isPublic = chkIsPublic.Checked;
            _StudInterviewRecord.AuthorID = txtAuthorID.Text;

            // 記錄者姓名
            _StudInterviewRecord.AuthorName = txtAuthorName.Text;

            return checkPass;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // 儲存畫面資料
            if (SaveData() == false)
            {
                FISCA.Presentation.Controls.MsgBox.Show("必填欄位沒有填寫，無法儲存.");
                return;            
            }

            // Log
            LogData();

            string studStr="學號:"+_studRec.StudentNumber+",姓名:"+_studRec.Name+",";
            
            if (_accessType == AccessType.Update)
            {
                // log
                _LogTransfer.SaveChangeLog("學生.晤談紀錄-修改", "修改", studStr, "", "student", _studRec.ID);            

                _UDTTransfer.UpdateCounselStudentInterviewRecord(_StudInterviewRecord);
            }
            else
            {
                // 檢查是否可已新增 日期+晤談事由 不能重複
                List<DAO.UDT_CounselStudentInterviewRecordDef> dataList = _UDTTransfer.GetCounselStudentInterviewRecordByStudentID(_studRec.ID);
                bool pass = true;
                foreach (DAO.UDT_CounselStudentInterviewRecordDef data in dataList)
                {
                    if (data.InterviewDate.HasValue && _StudInterviewRecord.InterviewDate.HasValue)
                        if (data.InterviewDate.Value.ToShortDateString() == _StudInterviewRecord.InterviewDate.Value.ToShortDateString())
                            if (data.Cause == _StudInterviewRecord.Cause)
                                pass = false;
                }

                if (pass)
                {
                    // log
                    _LogTransfer.SaveInsertLog("學生.晤談紀錄-新增", "新增", studStr, "", "student", _studRec.ID);
                    _UDTTransfer.InstallCounselStudentInterviewRecord(_StudInterviewRecord);
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("已有相同的日期與晤談事由，無法新增");
                    return;
                }
            }

            FISCA.Presentation.Controls.MsgBox.Show("儲存完成");
            this.Close();
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SelectRoleChanged_Click(object sender, EventArgs e)
        {
            List<LabelX> toBeChanged = new List<LabelX>();
            toBeChanged.Add(this.labelX4);
            toBeChanged.Add(this.labelX9);
            toBeChanged.Add(this.labelX8);
            toBeChanged.Add(this.labelX13);
            toBeChanged.Add(this.labelX1);

            bool isParent = (this.cbxInterveweeType.SelectedIndex == 1);
            changeCaption(toBeChanged, isParent);
        }

        private void changeCaption(List<LabelX> labels, bool isParent)
        {
            String prefix = (isParent ? "聯繫" : "晤談");
            foreach(LabelX lbl in labels)
            {
                lbl.Text = String.Format("{0}{1}", prefix, lbl.Text.Substring(2));
            }
        }
    }
}
