using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Counsel_System2.DAO;
using Campus.Windows;

namespace Counsel_System2.Contents
{
    [FISCA.Permission.FeatureCode(PermissionCode.綜合表現紀錄表_資料項目, "綜合表現紀錄表-本人概況")]
    public partial class StudABCard01Content : FISCA.Presentation.DetailContent
    {
        /// <summary>
        /// 題目key
        /// </summary>
        private enum enumKey { 本人概況_血型, 本人概況_宗教, 本人概況_身高, 本人概況_體重, 本人概況_生理缺陷, 本人概況_曾患特殊疾病, 本人概況_原住民血統, 本人概況_手機號碼 }

        List<string> _StudenIDList;
        int _intStudentID = 0;
        private ChangeListener _ChangeManager = new ChangeListener();

        // 動態計算位置使用
        int y = 0;
        bool _reloadQuestion = true;
        string _Q_Flp01 = "血型";
        string _Q_Flp02 = "宗教";
        string _Q_Flp03 = "生理缺陷";
        string _Q_Flp04 = "曾患特殊疾病";

        /// <summary>
        /// 本人概況_血型
        /// </summary>
        UDTSingleRecordDef _udtSrFlp01;

        /// <summary>
        /// 本人概況_宗教
        /// </summary>
        UDTSingleRecordDef _udtSrFlp02;

        /// <summary>
        /// 本人概況_身高
        /// </summary>
        UDTSemesterDataDef _udtSdDg01;

        /// <summary>
        /// 本人概況_體重
        /// </summary>
        UDTSemesterDataDef _udtSdDg02;

        /// <summary>
        /// 本人概況_生理缺陷
        /// </summary>
        Dictionary<string, UDTMultipleRecordDef> _udtMrFlp01Dict;

        /// <summary>
        /// 本人概況_曾患特殊疾病
        /// </summary>
        Dictionary<string, UDTMultipleRecordDef> _udtMrFlp02Dict;

        /// <summary>
        /// 本人概況_原住民血統
        /// </summary>
        UDTSingleRecordDef _udtSrFlp05;

        
        //2016/9/20，穎驊新增，增加學生手機號碼欄位
        
        /// <summary>
        /// 本人概況_手機號碼
        /// </summary>
        UDTSingleRecordDef _udtSrFlp06;




        Dictionary<string, QuestionData> _QuestionDict;

        /// <summary>
        /// 身高
        /// </summary>
        int _Dg01RowIdx;
        /// <summary>
        /// 體重
        /// </summary>
        int _Dg02RowIdx;

        private BackgroundWorker _bgWorker;
        bool _isBusy = false;
        ABCardQuestionDataManager _QDMang;

        private string GroupName = "本人概況";
        public StudABCard01Content()
        {
            InitializeComponent();

            _StudenIDList = new List<string>();
            _udtMrFlp01Dict = new Dictionary<string, UDTMultipleRecordDef>();
            _udtMrFlp02Dict = new Dictionary<string, UDTMultipleRecordDef>();

            _QDMang = new ABCardQuestionDataManager();
            _QuestionDict = _QDMang.GetQuestionDataByGroupName(GroupName);
            this.Group = "綜合表現紀錄表-本人概況";
            _bgWorker = new BackgroundWorker();

            List<string> grYear = Utility.GetClassGradeYear();

            if (grYear.Count > 3)
                SetDgColumn4_6Visable(true);
            else
                SetDgColumn4_6Visable(false);

            _ChangeManager = new ChangeListener();
            _ChangeManager.StatusChanged += new EventHandler<ChangeEventArgs>(_ChangeManager_StatusChanged);
            LoadQuestionToUI();
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);
            EventHub.CounselChanged += new EventHandler(EventHub_CounselChanged);
        }

        void EventHub_CounselChanged(object sender, EventArgs e)
        {
            _reloadQuestion = true;
        }

        void _ChangeManager_StatusChanged(object sender, ChangeEventArgs e)
        {
            this.CancelButtonVisible = (e.Status == ValueStatus.Dirty);
            this.SaveButtonVisible = (e.Status == ValueStatus.Dirty);
        }
        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_isBusy)
            {
                _isBusy = false;
                _bgWorker.RunWorkerAsync();
                return;
            }
            ReLoadData();
        }

        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            _StudenIDList.Clear();
            _StudenIDList.Add(PrimaryKey);
            _intStudentID = int.Parse(PrimaryKey);
            _BGRun();
        }

        /// <summary>
        /// 設定控制項初始值
        /// </summary>
        private void SetControlsDefault()
        {
            foreach (Control c in flp01.Controls)
            {
                if (c is RadioButton)
                {
                    RadioButton rb = c as RadioButton;
                    rb.Checked = false;
                }

                if (c is TextBox)
                    c.Text = "";
            }

            foreach (Control c in flp02.Controls)
            {
                if (c is RadioButton)
                {
                    RadioButton rb = c as RadioButton;
                    rb.Checked = false;
                }

                if (c is TextBox)
                    c.Text = "";
            }

            SetDg01RowsDefault();

            foreach (Control c in flp03.Controls)
            {
                if (c is CheckBox)
                {
                    CheckBox cb = c as CheckBox;
                    cb.Checked = false;
                }

                if (c is TextBox)
                    c.Text = "";
            }

            foreach (Control c in flp04.Controls)
            {
                if (c is CheckBox)
                {
                    CheckBox cb = c as CheckBox;
                    cb.Checked = false;
                }

                if (c is TextBox)
                    c.Text = "";
            }
        }

        private void ReLoadData()
        {
            _ChangeManager.SuspendListen();

            if (_reloadQuestion)
                LoadQuestionToUI();

            SetControlsDefault();
            BindAnswerDataToUI();
            _ChangeManager.Reset();
            _ChangeManager.ResumeListen();
            this.Loading = false;
            _reloadQuestion = false;
        }

        protected override void OnSaveButtonClick(EventArgs e)
        {
            dg01.EndEdit();
            // 檢查資料
            bool hasError = false;
            foreach (DataGridViewRow drv in dg01.Rows)
            {
                foreach (DataGridViewCell drc in drv.Cells)
                    if (drc.ErrorText != "")
                        hasError = true;
            }

            if (hasError)
            {
                FISCA.Presentation.Controls.MsgBox.Show("身高與體重資料有錯誤，無法儲存.");
                return;
            }

            SaveAnswerToUDT();

            this.CancelButtonVisible = false;
            this.SaveButtonVisible = false;
            _BGRun();
        }

        private void SaveAnswerToUDT()
        {
            List<UDTSingleRecordDef> iDataList = new List<UDTSingleRecordDef>();
            List<UDTSingleRecordDef> uDataList = new List<UDTSingleRecordDef>();
            List<UDTSemesterDataDef> isDataList = new List<UDTSemesterDataDef>();
            List<UDTSemesterDataDef> usDataList = new List<UDTSemesterDataDef>();

            //原住民血統
            if (_udtSrFlp05 == null)
            {
                _udtSrFlp05 = new UDTSingleRecordDef();
                _udtSrFlp05.StudentID = _intStudentID;
            }
            _udtSrFlp05.Key = enumKey.本人概況_原住民血統.ToString();

            if (radioButton1.Checked)
            {
                _udtSrFlp05.Data = radioButton1.Text;
                _udtSrFlp05.Remark = txtParents.Text + "_" + txtGroupName.Text;
            }
            else
            {
                _udtSrFlp05.Data = radioButton2.Text;
                _udtSrFlp05.Remark = "";
            }
            

            if (string.IsNullOrEmpty(_udtSrFlp05.UID))
                iDataList.Add(_udtSrFlp05);
            else
                uDataList.Add(_udtSrFlp05);

            // 血型
            if (_udtSrFlp01 == null)
            {
                _udtSrFlp01 = new UDTSingleRecordDef();
                _udtSrFlp01.StudentID = _intStudentID;
            }
            _udtSrFlp01.Key = enumKey.本人概況_血型.ToString();
            foreach (Control c in flp01.Controls)
            {
                if (c is RadioButton)
                {
                    _udtSrFlp01.Remark = "";
                    RadioButton rb = c as RadioButton;
                    if (rb.Checked)
                    {
                        _udtSrFlp01.Data = rb.Name;

                        foreach (Control cr in flp01.Controls)
                            if (cr is TextBox)
                                if (cr.Name == rb.Name)
                                    _udtSrFlp01.Remark = cr.Text;
                    }
                }
            }
            if (string.IsNullOrEmpty(_udtSrFlp01.UID))
                iDataList.Add(_udtSrFlp01);
            else
                uDataList.Add(_udtSrFlp01);

            // 宗教
            if (_udtSrFlp02 == null)
            {
                _udtSrFlp02 = new UDTSingleRecordDef();
                _udtSrFlp02.StudentID = _intStudentID;
            }

            _udtSrFlp02.Key = enumKey.本人概況_宗教.ToString();
            foreach (Control c in flp02.Controls)
            {
                if (c is RadioButton)
                {
                    _udtSrFlp02.Remark = "";
                    RadioButton rb = c as RadioButton;
                    if (rb.Checked)
                    {
                        _udtSrFlp02.Data = rb.Name;

                        foreach (Control cr in flp02.Controls)
                            if (cr is TextBox)
                                if (cr.Name == rb.Name)
                                    _udtSrFlp02.Remark = cr.Text;

                    }
                }

            }
            if (string.IsNullOrEmpty(_udtSrFlp02.UID))
                iDataList.Add(_udtSrFlp02);
            else
                uDataList.Add(_udtSrFlp02);

            if (iDataList.Count > 0)
                UDTTransfer.ABUDTSingleRecordInsert(iDataList);

            if (uDataList.Count > 0)
                UDTTransfer.ABUDTSingleRecordUpdate(uDataList);


            // 寫入 身高
            if (_udtSdDg01 == null)
            {
                _udtSdDg01 = new UDTSemesterDataDef();
                _udtSdDg01.StudentID = _intStudentID;
                _udtSdDg01.Key = enumKey.本人概況_身高.ToString();
            }
            List<string> tmpData1 = new List<string>();
            for (int i = 1; i <= 12; i++)
            {
                if (dg01.Rows[_Dg01RowIdx].Cells[i].Value != null)
                    tmpData1.Add(dg01.Rows[_Dg01RowIdx].Cells[i].Value.ToString());
                else
                    tmpData1.Add("");
            }
            _udtSdDg01.S1a = tmpData1[0];
            _udtSdDg01.S1b = tmpData1[1];
            _udtSdDg01.S2a = tmpData1[2];
            _udtSdDg01.S2b = tmpData1[3];
            _udtSdDg01.S3a = tmpData1[4];
            _udtSdDg01.S3b = tmpData1[5];
            _udtSdDg01.S4a = tmpData1[6];
            _udtSdDg01.S4b = tmpData1[7];
            _udtSdDg01.S5a = tmpData1[8];
            _udtSdDg01.S5b = tmpData1[9];
            _udtSdDg01.S6a = tmpData1[10];
            _udtSdDg01.S6b = tmpData1[11];

            if (string.IsNullOrEmpty(_udtSdDg01.UID))
                isDataList.Add(_udtSdDg01);
            else
                usDataList.Add(_udtSdDg01);

            // 體重
            if (_udtSdDg02 == null)
            {
                _udtSdDg02 = new UDTSemesterDataDef();
                _udtSdDg02.StudentID = _intStudentID;
                _udtSdDg02.Key = enumKey.本人概況_體重.ToString();
            }
            List<string> tmpData2 = new List<string>();
            for (int i = 1; i <= 12; i++)
            {
                if (dg01.Rows[_Dg02RowIdx].Cells[i].Value != null)
                    tmpData2.Add(dg01.Rows[_Dg02RowIdx].Cells[i].Value.ToString());
                else
                    tmpData2.Add("");
            }
            _udtSdDg02.S1a = tmpData2[0];
            _udtSdDg02.S1b = tmpData2[1];
            _udtSdDg02.S2a = tmpData2[2];
            _udtSdDg02.S2b = tmpData2[3];
            _udtSdDg02.S3a = tmpData2[4];
            _udtSdDg02.S3b = tmpData2[5];
            _udtSdDg02.S4a = tmpData2[6];
            _udtSdDg02.S4b = tmpData2[7];
            _udtSdDg02.S5a = tmpData2[8];
            _udtSdDg02.S5b = tmpData2[9];
            _udtSdDg02.S6a = tmpData2[10];
            _udtSdDg02.S6b = tmpData2[11];

            if (string.IsNullOrEmpty(_udtSdDg02.UID))
                isDataList.Add(_udtSdDg02);
            else
                usDataList.Add(_udtSdDg02);

            if (isDataList.Count > 0)
                UDTTransfer.ABUDTSemesterDataInsert(isDataList);

            if (usDataList.Count > 0)
                UDTTransfer.ABUDTSemesterDataUpdate(usDataList);


            List<UDTMultipleRecordDef> delData = new List<UDTMultipleRecordDef>();
            List<UDTMultipleRecordDef> InsertData = new List<UDTMultipleRecordDef>();

            delData = _udtMrFlp01Dict.Values.ToList();

            if (delData.Count > 0)
                UDTTransfer.ABUDTMultipleRecordDelete(delData);

            foreach (Control c in flp03.Controls)
            {
                if (c is CheckBox)
                {
                    CheckBox cb = c as CheckBox;
                    if (cb.Checked)
                    {
                        UDTMultipleRecordDef data = new UDTMultipleRecordDef();
                        data.StudentID = _intStudentID;
                        data.Key = enumKey.本人概況_生理缺陷.ToString();
                        data.Data = cb.Name;
                        data.Remark = "";
                        foreach (Control cr in flp03.Controls)
                            if (cr is TextBox)
                                if (cr.Name == cb.Name)
                                    data.Remark = cr.Text;


                        InsertData.Add(data);
                    }
                }
            }
            if (InsertData.Count > 0)
                UDTTransfer.ABUDTMultipleRecordInsert(InsertData);


            delData.Clear();
            InsertData.Clear();

            delData = _udtMrFlp02Dict.Values.ToList();

            if (delData.Count > 0)
                UDTTransfer.ABUDTMultipleRecordDelete(delData);

            foreach (Control c in flp04.Controls)
            {
                if (c is CheckBox)
                {
                    CheckBox cb = c as CheckBox;
                    if (cb.Checked)
                    {
                        UDTMultipleRecordDef data = new UDTMultipleRecordDef();
                        data.StudentID = _intStudentID;
                        data.Key = enumKey.本人概況_曾患特殊疾病.ToString();
                        data.Data = cb.Name;
                        data.Remark = "";

                        foreach (Control cr in flp04.Controls)
                            if (cr is TextBox)
                                if (cr.Name == cb.Name)
                                    data.Remark = cr.Text;

                        InsertData.Add(data);
                    }
                }
            }
            if (InsertData.Count > 0)
                UDTTransfer.ABUDTMultipleRecordInsert(InsertData);

        }


        protected override void OnCancelButtonClick(EventArgs e)
        {
            this.CancelButtonVisible = false;
            this.SaveButtonVisible = false;
            ReLoadData();
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if (_reloadQuestion)
                _QuestionDict = _QDMang.GetQuestionDataByGroupName(GroupName);

            LoadAnswerData();
        }

        private void _BGRun()
        {
            if (_bgWorker.IsBusy)
                _isBusy = true;
            else
            {
                this.Loading = true;
                _bgWorker.RunWorkerAsync();
            }
        }

        private void SetDg01RowsDefault()
        {
            dg01.Rows.Clear();
            _Dg01RowIdx = dg01.Rows.Add();
            dg01.Rows[_Dg01RowIdx].Cells[colName.Index].Value = "身高";
            dg01.Rows[_Dg01RowIdx].Cells[colName.Index].ReadOnly = true;
            _Dg02RowIdx = dg01.Rows.Add();
            dg01.Rows[_Dg02RowIdx].Cells[colName.Index].Value = "體重";
            dg01.Rows[_Dg02RowIdx].Cells[colName.Index].ReadOnly = true;
        }

        /// <summary>
        /// 將答案填入畫面
        /// </summary>
        private void BindAnswerDataToUI()
        {
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            //原住民血統
            if (_udtSrFlp05 != null)
            {
                if (_udtSrFlp05.Data == radioButton1.Text)
                {
                    radioButton1.Checked = true;
                    if (!string.IsNullOrWhiteSpace(_udtSrFlp05.Remark) && _udtSrFlp05.Remark.Contains("_"))
                    {
                        string[] str = _udtSrFlp05.Remark.Split('_');
                        txtParents.Text = str[0];
                        txtGroupName.Text = str[1];
                    }
                    else
                    {
                        txtParents.Text = "";
                        txtGroupName.Text = "";
                    }
                }
                else
                {
                    radioButton2.Checked = true;
                }
            }

  

            //血型
            if (_udtSrFlp01 != null)
            {
                foreach (Control c in flp01.Controls)
                {
                    if (c is RadioButton)
                    {
                        RadioButton rb = c as RadioButton;
                        if (_udtSrFlp01 != null)
                            if (rb.Name == _udtSrFlp01.Data)
                            {
                                rb.Checked = true;

                                if (!string.IsNullOrEmpty(_udtSrFlp01.Remark))
                                {
                                    foreach (Control cr in flp01.Controls)
                                        if (cr is TextBox)
                                            if (cr.Name == _udtSrFlp01.Data)
                                                cr.Text = _udtSrFlp01.Remark;
                                }
                            }
                    }
                }
            }

            if (_udtSrFlp02 != null)
            {
                foreach (Control c in flp02.Controls)
                {
                    if (c is RadioButton)
                    {
                        RadioButton rb = c as RadioButton;
                        if (_udtSrFlp02 != null)
                            if (rb.Name == _udtSrFlp02.Data)
                            {
                                rb.Checked = true;

                                if (!string.IsNullOrEmpty(_udtSrFlp02.Remark))
                                {
                                    foreach (Control cr in flp02.Controls)
                                        if (cr is TextBox)
                                            if (cr.Name == _udtSrFlp02.Data)
                                                cr.Text = _udtSrFlp02.Remark;
                                }
                            }
                    }
                }
            }

            // 身高
            if (_udtSdDg01 != null)
            {
                dg01.Rows[_Dg01RowIdx].Cells[1].Value = _udtSdDg01.S1a;
                dg01.Rows[_Dg01RowIdx].Cells[2].Value = _udtSdDg01.S1b;
                dg01.Rows[_Dg01RowIdx].Cells[3].Value = _udtSdDg01.S2a;
                dg01.Rows[_Dg01RowIdx].Cells[4].Value = _udtSdDg01.S2b;
                dg01.Rows[_Dg01RowIdx].Cells[5].Value = _udtSdDg01.S3a;
                dg01.Rows[_Dg01RowIdx].Cells[6].Value = _udtSdDg01.S3b;
                dg01.Rows[_Dg01RowIdx].Cells[7].Value = _udtSdDg01.S4a;
                dg01.Rows[_Dg01RowIdx].Cells[8].Value = _udtSdDg01.S4b;
                dg01.Rows[_Dg01RowIdx].Cells[9].Value = _udtSdDg01.S5a;
                dg01.Rows[_Dg01RowIdx].Cells[10].Value = _udtSdDg01.S5b;
                dg01.Rows[_Dg01RowIdx].Cells[11].Value = _udtSdDg01.S6a;
                dg01.Rows[_Dg01RowIdx].Cells[12].Value = _udtSdDg01.S6b;
            }

            // 體重
            if (_udtSdDg02 != null)
            {
                dg01.Rows[_Dg02RowIdx].Cells[1].Value = _udtSdDg02.S1a;
                dg01.Rows[_Dg02RowIdx].Cells[2].Value = _udtSdDg02.S1b;
                dg01.Rows[_Dg02RowIdx].Cells[3].Value = _udtSdDg02.S2a;
                dg01.Rows[_Dg02RowIdx].Cells[4].Value = _udtSdDg02.S2b;
                dg01.Rows[_Dg02RowIdx].Cells[5].Value = _udtSdDg02.S3a;
                dg01.Rows[_Dg02RowIdx].Cells[6].Value = _udtSdDg02.S3b;
                dg01.Rows[_Dg02RowIdx].Cells[7].Value = _udtSdDg02.S4a;
                dg01.Rows[_Dg02RowIdx].Cells[8].Value = _udtSdDg02.S4b;
                dg01.Rows[_Dg02RowIdx].Cells[9].Value = _udtSdDg02.S5a;
                dg01.Rows[_Dg02RowIdx].Cells[10].Value = _udtSdDg02.S5b;
                dg01.Rows[_Dg02RowIdx].Cells[11].Value = _udtSdDg02.S6a;
                dg01.Rows[_Dg02RowIdx].Cells[12].Value = _udtSdDg02.S6b;
            }

            foreach (Control c in flp03.Controls)
            {
                if (c is CheckBox)
                {
                    CheckBox cb = c as CheckBox;
                    if (_udtMrFlp01Dict.ContainsKey(cb.Name))
                    {
                        cb.Checked = true;
                        if (!string.IsNullOrEmpty(_udtMrFlp01Dict[cb.Name].Remark))
                        {
                            foreach (Control cr in flp03.Controls)
                                if (cr is TextBox)
                                    if (cr.Name == _udtMrFlp01Dict[cb.Name].Data)
                                        cr.Text = _udtMrFlp01Dict[cb.Name].Remark;
                        }
                    }
                }
            }

            foreach (Control c in flp04.Controls)
            {
                if (c is CheckBox)
                {
                    CheckBox cb = c as CheckBox;
                    if (_udtMrFlp02Dict.ContainsKey(cb.Name))
                    {
                        cb.Checked = true;

                        if (!string.IsNullOrEmpty(_udtMrFlp02Dict[cb.Name].Remark))
                        {
                            foreach (Control cr in flp04.Controls)
                                if (cr is TextBox)
                                    if (cr.Name == _udtMrFlp02Dict[cb.Name].Data)
                                        cr.Text = _udtMrFlp02Dict[cb.Name].Remark;
                        }
                    }

                }
            }

        }


        /// <summary>
        /// 載入答案
        /// </summary>
        private void LoadAnswerData()
        {
            _udtMrFlp01Dict.Clear();
            _udtMrFlp02Dict.Clear();
            _udtSdDg01 = _udtSdDg02 = null;
            _udtSrFlp01 = _udtSrFlp02 = null;
            _udtSrFlp05 = null;

            List<UDTSingleRecordDef> SingleRecordList = UDTTransfer.ABUDTSingleRecordSelectByStudentIDList(_StudenIDList);
            foreach (UDTSingleRecordDef data in SingleRecordList)
            {
                if (data.Key == enumKey.本人概況_血型.ToString())
                    _udtSrFlp01 = data;

                if (data.Key == enumKey.本人概況_宗教.ToString())
                    _udtSrFlp02 = data;

                if (data.Key == enumKey.本人概況_原住民血統.ToString())
                    _udtSrFlp05 = data;


                //2016/9/20 穎驊新增，手機號碼
                if (data.Key == enumKey.本人概況_手機號碼.ToString())
                    _udtSrFlp06 = data;

            }
            List<UDTSemesterDataDef> SemesterDataList = UDTTransfer.ABUDTSemesterDataSelectByStudentIDList(_StudenIDList);

            foreach (UDTSemesterDataDef data in SemesterDataList)
            {
                if (data.Key == enumKey.本人概況_身高.ToString())
                    _udtSdDg01 = data;

                if (data.Key == enumKey.本人概況_體重.ToString())
                    _udtSdDg02 = data;
            }

            List<UDTMultipleRecordDef> MultipleRecordList = UDTTransfer.ABUDTMultipleRecordSelectByStudentIDList(_StudenIDList);
            foreach (UDTMultipleRecordDef data in MultipleRecordList)
            {
                if (data.Key == enumKey.本人概況_生理缺陷.ToString())
                    if (!_udtMrFlp01Dict.ContainsKey(data.Data))
                        _udtMrFlp01Dict.Add(data.Data, data);

                if (data.Key == enumKey.本人概況_曾患特殊疾病.ToString())
                    if (!_udtMrFlp02Dict.ContainsKey(data.Data))
                        _udtMrFlp02Dict.Add(data.Data, data);
            }

        }

        /// <summary>
        /// 將題目載入
        /// </summary>
        private void LoadQuestionToUI()
        {
            List<RadioButton> cm01 = new List<RadioButton>();
            List<RadioButton> cm02 = new List<RadioButton>();
            List<CheckBox> cm03 = new List<CheckBox>();
            List<CheckBox> cm04 = new List<CheckBox>();

            y = flp01.Location.Y;
            labelX1.Location = new Point(labelX1.Location.X, y + 4);
            flp01.Controls.Clear();
            flp02.Controls.Clear();
            flp03.Controls.Clear();
            flp04.Controls.Clear();

            if (_QuestionDict.ContainsKey(_Q_Flp01))
            {
                flp01.AutoScroll = true;
                flp01.FlowDirection = FlowDirection.LeftToRight;

                foreach (QuestionItem qi in _QuestionDict[_Q_Flp01].itemList)
                {
                    RadioButton rb = new RadioButton();
                    rb.Name = qi.Key;
                    if (qi.hasRemark)
                        rb.Text = qi.Key + "：";
                    else
                        rb.Text = qi.Key;

                    rb.AutoSize = true;
                    cm01.Add(rb);

                    flp01.Controls.Add(rb);
                    if (qi.hasRemark)
                    {
                        TextBox tb = new TextBox();
                        tb.Name = qi.Key;
                        tb.Width = 70;
                        tb.Text = "";
                        flp01.Controls.Add(tb);

                    }
                }
            }

            y = y + flp01.Size.Height + 12;
            flp02.Location = new Point(flp02.Location.X, y);
            labelX2.Location = new Point(labelX2.Location.X, y + 4);

            if (_QuestionDict.ContainsKey(_Q_Flp02))
            {
                flp02.AutoScroll = true;
                flp02.FlowDirection = FlowDirection.LeftToRight;

                foreach (QuestionItem qi in _QuestionDict[_Q_Flp02].itemList)
                {
                    RadioButton rb = new RadioButton();
                    rb.Name = qi.Key;
                    if (qi.hasRemark)
                        rb.Text = qi.Key + "：";

                    else
                        rb.Text = qi.Key;
                    rb.AutoSize = true;
                    cm02.Add(rb);
                    flp02.Controls.Add(rb);
                    if (qi.hasRemark)
                    {
                        TextBox tb = new TextBox();
                        tb.Name = qi.Key;
                        tb.Width = 70;
                        tb.Text = "";
                        flp02.Controls.Add(tb);
                    }
                }
            }

            SetDg01RowsDefault();


            y = y + flp02.Size.Height + 12;
            flp03.Location = new Point(flp03.Location.X, y);
            labelX3.Location = new Point(labelX3.Location.X, y + 4);

            if (_QuestionDict.ContainsKey(_Q_Flp03))
            {
                flp03.AutoScroll = true;
                flp03.FlowDirection = FlowDirection.LeftToRight;

                foreach (QuestionItem qi in _QuestionDict[_Q_Flp03].itemList)
                {
                    CheckBox cb = new CheckBox();
                    cb.Name = qi.Key;
                    if (qi.hasRemark)
                        cb.Text = qi.Key + "：";

                    else
                        cb.Text = qi.Key;
                    cb.AutoSize = true;
                    cm03.Add(cb);

                    flp03.Controls.Add(cb);
                    if (qi.hasRemark)
                    {
                        TextBox tb = new TextBox();
                        tb.Name = qi.Key;
                        tb.Width = 70;
                        tb.Text = "";
                        flp03.Controls.Add(tb);
                    }
                }
            }

            y = y + flp03.Size.Height + 12;
            flp04.Location = new Point(flp04.Location.X, y);
            labelX4.Location = new Point(labelX4.Location.X, y + 4);

            if (_QuestionDict.ContainsKey(_Q_Flp04))
            {
                flp04.AutoScroll = true;
                flp04.FlowDirection = FlowDirection.LeftToRight;

                foreach (QuestionItem qi in _QuestionDict[_Q_Flp04].itemList)
                {
                    CheckBox cb = new CheckBox();
                    cb.Name = qi.Key;
                    if (qi.hasRemark)
                        cb.Text = qi.Key + "：";

                    else
                        cb.Text = qi.Key;

                    cb.AutoSize = true;
                    cm04.Add(cb);

                    flp04.Controls.Add(cb);
                    if (qi.hasRemark)
                    {
                        TextBox tb = new TextBox();
                        tb.Name = qi.Key;
                        tb.Width = 70;
                        tb.Text = "";
                        flp04.Controls.Add(tb);
                    }

                }
            }

            //原住民
            y = y + flp04.Size.Height + 12;
            flp05.Location = new Point(flp05.Location.X, y);

            // 調整 datagridview
            y = y + flp05.Size.Height + 12;
            labelX5.Location = new Point(labelX5.Location.X, y);

            y = y + 25;

            dg01.Location = new Point(dg01.Location.X, y);

            y = y + dg01.Size.Height + 12;
            this.Size = new System.Drawing.Size(this.Width, y);

            // 加入資料變更管理
            _ChangeManager.Add(new DataGridViewSource(dg01));
            _ChangeManager.Add(new RadioButtonSource(cm01.ToArray()));
            _ChangeManager.Add(new RadioButtonSource(cm02.ToArray()));
            _ChangeManager.Add(new CheckBoxSource(cm03.ToArray()));
            _ChangeManager.Add(new CheckBoxSource(cm04.ToArray()));

            _ChangeManager.Add(new RadioButtonSource(radioButton1));
            _ChangeManager.Add(new RadioButtonSource(radioButton2));
            _ChangeManager.Add(new TextBoxSource(txtParents));
            _ChangeManager.Add(new TextBoxSource(txtGroupName));

            foreach (Control c in flp01.Controls)
                if (c is TextBox)
                {
                    TextBox tb = c as TextBox;
                    _ChangeManager.Add(new TextBoxSource(tb));
                }

            foreach (Control c in flp02.Controls)
                if (c is TextBox)
                {
                    TextBox tb = c as TextBox;
                    _ChangeManager.Add(new TextBoxSource(tb));
                }

            foreach (Control c in flp03.Controls)
                if (c is TextBox)
                {
                    TextBox tb = c as TextBox;
                    _ChangeManager.Add(new TextBoxSource(tb));
                }

            foreach (Control c in flp04.Controls)
                if (c is TextBox)
                {
                    TextBox tb = c as TextBox;
                    _ChangeManager.Add(new TextBoxSource(tb));
                }


        }

        /// <summary>
        /// 設定DG四~六年級是否顯示
        /// </summary>
        /// <param name="bol"></param>
        private void SetDgColumn4_6Visable(bool bol)
        {
            colS4a.Visible = colS4b.Visible = colS5a.Visible = colS5b.Visible = colS6a.Visible = colS6b.Visible = bol;
        }

        private void dg01_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dg01.EndEdit();

            this.SaveButtonVisible = true;
            this.CancelButtonVisible = true;
            if (dg01.CurrentCell.ColumnIndex != colName.Index)
            {
                double num;

                if (dg01.CurrentCell.Value != null)
                {
                    if (double.TryParse(dg01.CurrentCell.Value.ToString(), out num) == false)
                        dg01.CurrentCell.ErrorText = "必須為數字";
                    else
                        dg01.CurrentCell.ErrorText = "";

                }

            }
            dg01.BeginEdit(false);

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                txtParents.Enabled = true;
                txtGroupName.Enabled = true;
            }
            else
            {
                txtParents.Enabled = false;
                txtGroupName.Enabled = false;
                txtParents.Text = "";
                txtGroupName.Text = "";
            }
        }

    }
}
