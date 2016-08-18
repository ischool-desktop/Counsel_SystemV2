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
    [FISCA.Permission.FeatureCode(PermissionCode.綜合表現紀錄表_資料項目, "綜合表現紀錄表-家庭狀況")]
    public partial class StudABCard02Content : FISCA.Presentation.DetailContent
    {
        string GroupName = "家庭狀況";
        
        BackgroundWorker _bgWorker;
        ChangeListener _ChangeListener = new ChangeListener();
        List<string> _StudentIDList;
        int _intStudentID = 0;
        Dictionary<string, QuestionData> _QuestionDataDict;
        ABCardQuestionDataManager _QDMang;

        bool _isBusy = false;
        bool _reloadQuestion = true;
        List<UDTSiblingDef> _UDTSiblingList;
        List<UDTRelativeDef> _UDTRelativeList;
        List<string> _UDTYearlyDataKeyList;
        List<string> _UDTSiblingKeyList;
        List<string> _UDTRelativeKeyList;
        Dictionary<string,UDTYearlyDataDef> _UDTYearlyDataDict;

        string KeySiblingNoStr = "家庭狀況_兄弟姊妹_排行";

        // 監護人使用
        Dictionary<string, UDTSingleRecordDef> _UDTSingleRecordDict;
        /// <summary>
        /// 年級索引(班級)
        /// </summary>
        Dictionary<string, int> _ClassGradeYearDict;

        /// <summary>
        /// 資料存取對照用 Index
        /// </summary>
        Dictionary<string, int> _RowIndexDict;
        /// <summary>
        /// 資料存取對照用 Index
        /// </summary>
        List<string> _RowNameList;

        /// <summary>
        /// 資料存取對照用 Index
        /// </summary>
        Dictionary<string, int> _ColumIndexDict;

        public StudABCard02Content()
        {
            InitializeComponent();
            this.Group = "綜合表現紀錄表-家庭狀況";
            _QDMang = new ABCardQuestionDataManager();
            _StudentIDList = new List<string>();
           
            _UDTSingleRecordDict = new Dictionary<string, UDTSingleRecordDef>();
            _UDTSiblingList = new List<UDTSiblingDef>();
            _UDTRelativeList = new List<UDTRelativeDef>();
            _UDTYearlyDataKeyList = new List<string> ();
            _UDTRelativeKeyList = new List<string>();
            _UDTSiblingKeyList = new List<string>();
            _UDTYearlyDataDict = new Dictionary<string, UDTYearlyDataDef>();

            LoadQuestion();
            _ClassGradeYearDict = Utility.GetClassGradeYearDict();

            // 會動態改變
            _RowIndexDict = new Dictionary<string, int>();
            _RowNameList = _ClassGradeYearDict.Keys.ToList();
            _RowNameList.Sort();

            int row = 0;
            foreach (string str in _RowNameList)
            {
                _RowIndexDict.Add(str, row);
                row++;
            }

            // 固定不變
            _ColumIndexDict = new Dictionary<string, int>();
            _ColumIndexDict.Add(GroupName + "_父母關係", 1);
            _ColumIndexDict.Add(GroupName + "_家庭氣氛", 2);
            _ColumIndexDict.Add(GroupName + "_父親管教方式", 3);
            _ColumIndexDict.Add(GroupName + "_母親管教方式", 4);
            _ColumIndexDict.Add(GroupName + "_居住環境", 5);
            _ColumIndexDict.Add(GroupName + "_本人住宿", 6);
            _ColumIndexDict.Add(GroupName + "_經濟狀況", 7);
            _ColumIndexDict.Add(GroupName + "_每星期零用錢", 8);
            _ColumIndexDict.Add(GroupName + "_我覺得是否足夠", 9);

            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);            
            _ChangeListener.StatusChanged += new EventHandler<ChangeEventArgs>(_ChangeListener_StatusChanged);
            EventHub.CounselChanged += new EventHandler(EventHub_CounselChanged);
            _ChangeListener.Add(new TextBoxSource(txtGender));
            _ChangeListener.Add(new TextBoxSource(txtMailingAddress));
            _ChangeListener.Add(new TextBoxSource(txtName));
            _ChangeListener.Add(new TextBoxSource(txtPhone));
            _ChangeListener.Add(new TextBoxSource(txtRelationship));
            _ChangeListener.Add(new TextBoxSource(txtSiblingNo));
            _ChangeListener.Add(new DataGridViewSource(dgRelative));
            _ChangeListener.Add(new DataGridViewSource(dgSibling));
            _ChangeListener.Add(new DataGridViewSource(dgYearly));
        }

        void EventHub_CounselChanged(object sender, EventArgs e)
        {
            _reloadQuestion = true;
        }

        private void LoadQuestion()
        {
            if (_reloadQuestion)
            {
                _QuestionDataDict = _QDMang.GetQuestionDataByGroupName(GroupName);

                foreach (QuestionData qd in _QuestionDataDict.Values)
                {
                    if (qd.QQuestionType == EnumQuestionType.YEARLY)
                        _UDTYearlyDataKeyList.Add(qd.Name);

                    if (qd.Name == "直系血親_稱謂")
                        _UDTRelativeKeyList = (from da in qd.itemList select da.Key).ToList();

                    if (qd.Name == "兄弟姊妹_姓名")
                        _UDTSiblingKeyList = (from da in qd.itemList select da.Key).ToList();
                }
            }           
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_isBusy)
            {
                _isBusy = false;
                _bgWorker.RunWorkerAsync();
                return;
            }
            _ChangeListener.SuspendListen();
            BindDataToDataGrid();
            _ChangeListener.Reset();
            _ChangeListener.ResumeListen();
            this.Loading = false;
            _reloadQuestion = false;
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            LoadQuestion();

            _UDTSingleRecordDict.Clear();
            string key = GroupName + "_監護人_";
            // 讀取資料
            foreach (UDTSingleRecordDef data in UDTTransfer.ABUDTSingleRecordSelectByStudentIDList(_StudentIDList))
            {
                // 監護人
                if (data.Key.Contains(key))
                {
                    string newkey = data.Key.Replace(key, "");

                    if (!_UDTSingleRecordDict.ContainsKey(newkey))
                        _UDTSingleRecordDict.Add(newkey, data);
                }

                // 兄弟姊妹排行
                if(data.Key.Contains(KeySiblingNoStr))
                {
                    if (!_UDTSingleRecordDict.ContainsKey(KeySiblingNoStr))
                    _UDTSingleRecordDict.Add(KeySiblingNoStr, data);                
                }
            }


            _UDTYearlyDataDict.Clear();
            string key1 = GroupName + "_";

            foreach (UDTYearlyDataDef data in UDTTransfer.ABUDTYearlyDataSelectByStudentIDList(_StudentIDList))
                if (data.Key.Contains(key1))
                    if (!_UDTYearlyDataDict.ContainsKey(data.Key))
                        _UDTYearlyDataDict.Add(data.Key, data);

            _UDTSiblingList = UDTTransfer.ABUDTSiblingSelectByStudentIDList(_StudentIDList);

            _UDTRelativeList = UDTTransfer.ABUDTRelativeSelectByStudentIDList(_StudentIDList);
        }
        
        void _ChangeListener_StatusChanged(object sender, ChangeEventArgs e)
        {
            this.CancelButtonVisible = (e.Status == ValueStatus.Dirty);
            this.SaveButtonVisible = (e.Status == ValueStatus.Dirty);
        }


        private void BindDataToDataGrid()
        {
            LoadGuardian();
            ReloadYearlyData();
            ReloadRelativeData();
            ReloadSiblingData();            
        }

        /// <summary>
        /// 載入監護人資料至畫面
        /// </summary>
        private void LoadGuardian()
        {
         
            txtName.Text = txtGender.Text = txtMailingAddress.Text = txtPhone.Text = txtRelationship.Text = "";            

            if(_UDTSingleRecordDict.ContainsKey("姓名"))
                txtName.Text=_UDTSingleRecordDict["姓名"].Data;

            if (_UDTSingleRecordDict.ContainsKey("性別"))
                txtGender.Text = _UDTSingleRecordDict["性別"].Data;

            if (_UDTSingleRecordDict.ContainsKey("關係"))
                txtRelationship.Text = _UDTSingleRecordDict["關係"].Data;


            if (_UDTSingleRecordDict.ContainsKey("電話"))
                txtPhone.Text = _UDTSingleRecordDict["電話"].Data;


            if (_UDTSingleRecordDict.ContainsKey("通訊地址"))
                txtMailingAddress.Text = _UDTSingleRecordDict["通訊地址"].Data;

            // 家庭狀況_兄弟姊妹_排行
            txtSiblingNo.Text = "";
            if (_UDTSingleRecordDict.ContainsKey(KeySiblingNoStr))
                txtSiblingNo.Text = _UDTSingleRecordDict[KeySiblingNoStr].Data;
        }


        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            _StudentIDList.Clear();
            _StudentIDList.Add(PrimaryKey);
            _intStudentID = int.Parse(PrimaryKey);
            _BGRun();
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

        protected override void OnCancelButtonClick(EventArgs e)
        {
            this.CancelButtonVisible = this.SaveButtonVisible = false;
            _BGRun();
        }

        protected override void OnSaveButtonClick(EventArgs e)
        {
            dgRelative.EndEdit();
            dgSibling.EndEdit();
            dgYearly.EndEdit();
            
            SaveGuardianData();
            SaveRelativeData();
            SaveSiblingData();
            SaveYearlyData();
            this.CancelButtonVisible = this.SaveButtonVisible = false;
            _BGRun();
        }

        /// <summary>
        /// 儲存監護人
        /// </summary>
        private void SaveGuardianData()
        {
            List<UDTSingleRecordDef> i_dataList = new List<UDTSingleRecordDef>();
            List<UDTSingleRecordDef> u_dataList = new List<UDTSingleRecordDef>();
            string saveKey1 = GroupName + "_監護人_";

            // 儲存監護人資料
            if (_UDTSingleRecordDict.ContainsKey("姓名"))
            {
                _UDTSingleRecordDict["姓名"].Data = txtName.Text;
                u_dataList.Add(_UDTSingleRecordDict["姓名"]);
            }
            else
            {
                UDTSingleRecordDef newData = new UDTSingleRecordDef();
                newData.Key = saveKey1 + "姓名";
                newData.StudentID = _intStudentID;
                newData.Data = txtName.Text;
                i_dataList.Add(newData);
            }

            if (_UDTSingleRecordDict.ContainsKey("性別"))
            {
                _UDTSingleRecordDict["性別"].Data = txtGender.Text;
                u_dataList.Add(_UDTSingleRecordDict["性別"]);
            }
            else
            {
                UDTSingleRecordDef newData = new UDTSingleRecordDef();
                newData.Key = saveKey1 + "性別";
                newData.StudentID = _intStudentID;
                newData.Data = txtGender.Text;
                i_dataList.Add(newData);
            }


            if (_UDTSingleRecordDict.ContainsKey("關係"))
            {
                _UDTSingleRecordDict["關係"].Data = txtRelationship.Text;
                u_dataList.Add(_UDTSingleRecordDict["關係"]);
            }
            else
            {
                UDTSingleRecordDef newData = new UDTSingleRecordDef();
                newData.Key = saveKey1 + "關係";
                newData.StudentID = _intStudentID;
                newData.Data = txtRelationship.Text;
                i_dataList.Add(newData);
            }

            if (_UDTSingleRecordDict.ContainsKey("電話"))
            {
                _UDTSingleRecordDict["電話"].Data = txtPhone.Text;
                u_dataList.Add(_UDTSingleRecordDict["電話"]);
            }
            else
            {
                UDTSingleRecordDef newData = new UDTSingleRecordDef();
                newData.Key = saveKey1 + "電話";
                newData.StudentID = _intStudentID;
                newData.Data = txtPhone.Text;
                i_dataList.Add(newData);
            }


            if (_UDTSingleRecordDict.ContainsKey("通訊地址"))
            {
                _UDTSingleRecordDict["通訊地址"].Data = txtMailingAddress.Text;
                u_dataList.Add(_UDTSingleRecordDict["通訊地址"]);
            }
            else
            {
                UDTSingleRecordDef newData = new UDTSingleRecordDef();
                newData.Key = saveKey1 + "通訊地址";
                newData.StudentID = _intStudentID;
                newData.Data = txtMailingAddress.Text;
                i_dataList.Add(newData);
            }

            // 家庭狀況_兄弟姊妹_排行
            if (_UDTSingleRecordDict.ContainsKey(KeySiblingNoStr))
            {
                _UDTSingleRecordDict[KeySiblingNoStr].Data = txtSiblingNo.Text;
                u_dataList.Add(_UDTSingleRecordDict[KeySiblingNoStr]);
            }
            else
            {
                UDTSingleRecordDef newData = new UDTSingleRecordDef();
                newData.Key = KeySiblingNoStr;
                newData.StudentID = _intStudentID;
                newData.Data = txtSiblingNo.Text;
                i_dataList.Add(newData);
            }


            if (i_dataList.Count > 0)
                UDTTransfer.ABUDTSingleRecordInsert(i_dataList);

            if (u_dataList.Count > 0)
                UDTTransfer.ABUDTSingleRecordUpdate(u_dataList);
        }

        /// <summary>
        /// 儲存直系血親
        /// </summary>
        private void SaveRelativeData()
        {
            List<UDTRelativeDef> insertDataList = new List<UDTRelativeDef>();
            List<UDTRelativeDef> updateDataList = new List<UDTRelativeDef>();
            foreach (DataGridViewRow dr in dgRelative.Rows)
            {
                if (dr.IsNewRow)
                    continue;
                UDTRelativeDef data = dr.Tag as UDTRelativeDef;

                if (data == null)
                {
                    data = new UDTRelativeDef();
                    data.StudentID = _intStudentID;
                    data.IsAlive = null;
                }

                // 稱謂
                if (dr.Cells[0].Value == null)
                    data.Title = "";
                else
                    data.Title = dr.Cells[0].Value.ToString();

                // 姓名
                if (dr.Cells[1].Value == null)
                    data.Name = "";
                else
                    data.Name = dr.Cells[1].Value.ToString();

                // 存歿
                if (dr.Cells[2].Value == null)
                    data.IsAlive = null;
                else
                {
                    if (dr.Cells[2].Value.ToString() == "存")
                        data.IsAlive = true;
                    else if (dr.Cells[2].Value.ToString() == "歿")
                        data.IsAlive = false;
                    else
                        data.IsAlive = null;
                }

                // 出生年
                if (dr.Cells[3].Value == null)
                    data.BirthYear = null;
                else
                {
                    int sy;
                    if (int.TryParse(dr.Cells[3].Value.ToString(), out sy))
                        data.BirthYear = sy;
                }

                // 職業
                if (dr.Cells[4].Value == null)
                    data.Job = "";
                else
                    data.Job = dr.Cells[4].Value.ToString();

                // 工作機構
                if (dr.Cells[5].Value == null)
                    data.Institute = "";
                else
                    data.Institute = dr.Cells[5].Value.ToString();

                // 職稱
                if (dr.Cells[6].Value == null)
                    data.JobTitle = "";
                else
                    data.JobTitle = dr.Cells[6].Value.ToString();
                
                // 教育程度
                if (dr.Cells[7].Value == null)
                    data.EduDegree = "";
                else
                    data.EduDegree = dr.Cells[7].Value.ToString();

                // 電話
                if (dr.Cells[8].Value == null)
                    data.Phone = "";
                else
                    data.Phone = dr.Cells[8].Value.ToString();

                //原國籍
                if (dr.Cells[9].Value == null)
                    data.National = "";
                else
                    data.National = dr.Cells[9].Value.ToString();

                if (string.IsNullOrEmpty(data.UID))
                    insertDataList.Add(data);
                else
                    updateDataList.Add(data);

                //行動電話，Added by kevin 2016/08/17
                if (dr.Cells[10].Value == null)
                    data.CellPhone = "";
                else
                    data.CellPhone = dr.Cells[10].Value.ToString();
            }

            // UDT
            if (insertDataList.Count > 0)
                UDTTransfer.ABUDTRelativeInsert(insertDataList);

            if (updateDataList.Count > 0)
                UDTTransfer.ABUDTRelativeUpdate(updateDataList);

        }

        /// <summary>
        /// 儲存兄弟姊妹
        /// </summary>
        private void SaveSiblingData()
        {
            List<UDTSiblingDef> insertData = new List<UDTSiblingDef>();
            List<UDTSiblingDef> updateData = new List<UDTSiblingDef>();

            foreach (DataGridViewRow dr in dgSibling.Rows)
            {
                if (dr.IsNewRow)
                    continue;
                UDTSiblingDef data = dr.Tag as UDTSiblingDef;
                if (data == null)
                {
                    data = new UDTSiblingDef();
                    data.StudentID = _intStudentID;
                }

                // 稱謂
                if (dr.Cells[0].Value == null)
                    data.Title = "";
                else
                    data.Title = dr.Cells[0].Value.ToString();

                // 姓名
                if (dr.Cells[1].Value == null)
                    data.Name = "";
                else
                    data.Name = dr.Cells[1].Value.ToString();


                // 畢業學校
                if (dr.Cells[2].Value == null)
                    data.SchoolName = "";
                else
                    data.SchoolName = dr.Cells[2].Value.ToString();

                // 出生年
                if (dr.Cells[3].Value == null)
                    data.BirthYear = null;
                else
                {
                    int sy;
                    if (int.TryParse(dr.Cells[3].Value.ToString(), out sy))
                        data.BirthYear = sy;
                    else
                        data.BirthYear = null;
                }

                // 備註
                if (dr.Cells[4].Value == null)
                    data.Remark = "";
                else
                    data.Remark= dr.Cells[4].Value.ToString();

                if (string.IsNullOrEmpty(data.UID))
                    insertData.Add(data);
                else
                    updateData.Add(data);
            }

            if (insertData.Count > 0)
                UDTTransfer.ABUDTSiblingInsert(insertData);

            if (updateData.Count > 0)
                UDTTransfer.ABUDTSiblingUpdate(updateData);
        }

        /// <summary>
        /// 儲存其它資料
        /// </summary>
        private void SaveYearlyData()
        {
            List<UDTYearlyDataDef> insertDataList = new List<UDTYearlyDataDef>();
            List<UDTYearlyDataDef> updateDataList = new List<UDTYearlyDataDef>();

            foreach (DataGridViewColumn drc in dgYearly.Columns)
            {
                string name = drc.HeaderText.Replace("(元)","");
                string key = GroupName + "_" + name;
                if (_ColumIndexDict.ContainsKey(key))
                {
                    UDTYearlyDataDef data = drc.Tag as UDTYearlyDataDef;
                    if (data == null)
                    {
                        data = new UDTYearlyDataDef();
                        data.StudentID = _intStudentID;
                    }
                    data.Key = key;

                    foreach (KeyValuePair<string, int> rowDa in _RowIndexDict)
                    {
                        if (rowDa.Key == "一年級")
                        {
                            if (dgYearly.Rows[rowDa.Value].Cells[drc.Index].Value != null)
                                data.G1 = dgYearly.Rows[rowDa.Value].Cells[drc.Index].Value.ToString();
                            else
                                data.G1 = "";
                        }

                        if (rowDa.Key == "二年級")
                        {
                            if (dgYearly.Rows[rowDa.Value].Cells[drc.Index].Value != null)
                                data.G2 = dgYearly.Rows[rowDa.Value].Cells[drc.Index].Value.ToString();
                            else
                                data.G2 = "";
                        }

                        if (rowDa.Key == "三年級")
                        {
                            if (dgYearly.Rows[rowDa.Value].Cells[drc.Index].Value != null)
                                data.G3 = dgYearly.Rows[rowDa.Value].Cells[drc.Index].Value.ToString();
                            else
                                data.G3 = "";
                        }

                        if (rowDa.Key == "四年級")
                        {
                            if (dgYearly.Rows[rowDa.Value].Cells[drc.Index].Value != null)
                                data.G4 = dgYearly.Rows[rowDa.Value].Cells[drc.Index].Value.ToString();
                            else
                                data.G4 = "";
                        }

                        if (rowDa.Key == "五年級")
                        {
                            if (dgYearly.Rows[rowDa.Value].Cells[drc.Index].Value != null)
                                data.G5 = dgYearly.Rows[rowDa.Value].Cells[drc.Index].Value.ToString();
                            else
                                data.G5 = "";
                        }

                        if (rowDa.Key == "六年級")
                        {
                            if (dgYearly.Rows[rowDa.Value].Cells[drc.Index].Value != null)
                                data.G6 = dgYearly.Rows[rowDa.Value].Cells[drc.Index].Value.ToString();
                            else
                                data.G6 = "";
                        }
                    }
                    if (string.IsNullOrEmpty(data.UID))
                        insertDataList.Add(data);
                    else
                        updateDataList.Add(data);
                }
            }        

            if (insertDataList.Count > 0)
                UDTTransfer.ABUDTYearlyDataInsert(insertDataList);

            if (updateDataList.Count > 0)
                UDTTransfer.ABUDTYearlyDataUpdate(updateDataList);
        }

        private void ReloadRelativeData()
        {
            dgRelative.Rows.Clear();
            int rowIdx = 0;
            dgRelative.AllowUserToAddRows = true;

            dgRelativeCol02.Items.Clear();
            dgRelativeCol02.Items.Add("存");
            dgRelativeCol02.Items.Add("歿");
            // 職業
            dgRelativeCol04.Items.Clear();

            // 職稱
            dgRelativeCol06.Items.Clear();

            // 教育程度
            dgRelativeCol07.Items.Clear();

            if (_QuestionDataDict.ContainsKey("直系血親_職業"))
            {
                QuestionData da = _QuestionDataDict["直系血親_職業"];
                if (da.itemList.Count > 0)
                    dgRelativeCol04.Items.AddRange((from dd in da.itemList select dd.Key).ToArray()); 
            }

            if (_QuestionDataDict.ContainsKey("直系血親_職稱"))
            {
                QuestionData da = _QuestionDataDict["直系血親_職稱"];
                if (da.itemList.Count > 0)
                    dgRelativeCol06.Items.AddRange((from dd in da.itemList select dd.Key).ToArray());
            }

            if (_QuestionDataDict.ContainsKey("直系血親_教育程度"))
            {
                QuestionData da = _QuestionDataDict["直系血親_教育程度"];
                if (da.itemList.Count > 0)
                    dgRelativeCol07.Items.AddRange((from dd in da.itemList select dd.Key).ToArray());
            }
            
            foreach (string key in _UDTRelativeKeyList)
            {
                rowIdx = dgRelative.Rows.Add();
                dgRelative.Rows[rowIdx].Cells[0].Value = key;

                foreach (UDTRelativeDef data in _UDTRelativeList)
                {
                    // 當稱謂相同
                    if (key == data.Title)
                    {
                        dgRelative.Rows[rowIdx].Tag = data;
                        dgRelative.Rows[rowIdx].Cells[1].Value = data.Name;
                        dgRelative.Rows[rowIdx].Cells[2].Value = "";
                        if (data.IsAlive.HasValue)
                            if (data.IsAlive.Value)
                                dgRelative.Rows[rowIdx].Cells[2].Value = "存";
                            else
                                dgRelative.Rows[rowIdx].Cells[2].Value = "歿";

                        dgRelative.Rows[rowIdx].Cells[3].Value = data.BirthYear;
                        dgRelative.Rows[rowIdx].Cells[4].Value = data.Job;
                        dgRelative.Rows[rowIdx].Cells[5].Value = data.Institute;
                        dgRelative.Rows[rowIdx].Cells[6].Value = data.JobTitle;
                        dgRelative.Rows[rowIdx].Cells[7].Value = data.EduDegree;
                        dgRelative.Rows[rowIdx].Cells[8].Value = data.Phone;
                        dgRelative.Rows[rowIdx].Cells[9].Value = data.National;
                        dgRelative.Rows[rowIdx].Cells[10].Value = data.CellPhone;
                    }
                }
            }
            dgRelative.AllowUserToAddRows = false;
        }

        private void ReloadSiblingData()
        {
            dgSibling.Rows.Clear();
            int rowIdx = 0;
            dgSibling.AllowUserToAddRows = true;
            //foreach (string key in _UDTSiblingKeyList)
            //{

                //dgSibling.Rows[rowIdx].Cells[1].Value = key;
                foreach (UDTSiblingDef data in _UDTSiblingList)
                {
                    rowIdx = dgSibling.Rows.Add();
                    dgSibling.Rows[rowIdx].Tag = data;
                    dgSibling.Rows[rowIdx].Cells[0].Value = data.Title;
                    dgSibling.Rows[rowIdx].Cells[1].Value = data.Name;
                    dgSibling.Rows[rowIdx].Cells[2].Value = data.SchoolName;
                    dgSibling.Rows[rowIdx].Cells[3].Value = data.BirthYear;
                    dgSibling.Rows[rowIdx].Cells[4].Value = data.Remark;
                }
            //}
            dgSibling.AllowUserToAddRows = false;
        }

        private void ReloadYearlyData()
        {
            dgYearly.Rows.Clear();                     
            dgYearly.AllowUserToAddRows = true;
            foreach (DataGridViewColumn col in dgYearly.Columns)
                col.Tag = null;


            // 選項
            dgYearlyCol01.Items.Clear();
            dgYearlyCol02.Items.Clear();
            dgYearlyCol03.Items.Clear();
            dgYearlyCol04.Items.Clear();
            dgYearlyCol05.Items.Clear();
            dgYearlyCol06.Items.Clear();
            dgYearlyCol07.Items.Clear();
            dgYearlyCol09.Items.Clear();
            
            //父母關係
            if (_QuestionDataDict.ContainsKey("父母關係"))
                dgYearlyCol01.Items.AddRange((from data in _QuestionDataDict["父母關係"].itemList select data.Key).ToArray());
            
            //家庭氣氛
            if (_QuestionDataDict.ContainsKey("家庭氣氛"))
                dgYearlyCol02.Items.AddRange((from data in _QuestionDataDict["家庭氣氛"].itemList select data.Key).ToArray());

            //父親管教方式
            if (_QuestionDataDict.ContainsKey("父親管教方式"))
                dgYearlyCol03.Items.AddRange((from data in _QuestionDataDict["父親管教方式"].itemList select data.Key).ToArray());

            //母親管教方式
            if (_QuestionDataDict.ContainsKey("母親管教方式"))
                dgYearlyCol04.Items.AddRange((from data in _QuestionDataDict["母親管教方式"].itemList select data.Key).ToArray());

            //居住環境
            if (_QuestionDataDict.ContainsKey("居住環境"))
                dgYearlyCol05.Items.AddRange((from data in _QuestionDataDict["居住環境"].itemList select data.Key).ToArray());

            //本人住宿
            if (_QuestionDataDict.ContainsKey("本人住宿"))
                dgYearlyCol06.Items.AddRange((from data in _QuestionDataDict["本人住宿"].itemList select data.Key).ToArray());

            //經濟狀況
            if (_QuestionDataDict.ContainsKey("經濟狀況"))
                dgYearlyCol07.Items.AddRange((from data in _QuestionDataDict["經濟狀況"].itemList select data.Key).ToArray());

            // 我覺得是否足夠
            if(_QuestionDataDict.ContainsKey("我覺得是否足夠"))
                dgYearlyCol09.Items.AddRange((from data in _QuestionDataDict["我覺得是否足夠"].itemList select data.Key).ToArray());

            int Row = 0;
            // 放入年級
            foreach (string str in _RowNameList)
            {
                Row = dgYearly.Rows.Add();
                dgYearly.Rows[Row].Cells[0].Value = str;
            }

            foreach (KeyValuePair<string, UDTYearlyDataDef> data in _UDTYearlyDataDict)
            {
                int ColIdx = 0;
                if (_ColumIndexDict.ContainsKey(data.Key))
                    ColIdx = _ColumIndexDict[data.Key];

                // 內容值放入 Tag
                dgYearly.Columns[ColIdx].Tag = data.Value;

                foreach (KeyValuePair<string, int> rowDa in _RowIndexDict)
                {
                    if (rowDa.Key == "一年級")
                        dgYearly.Rows[rowDa.Value].Cells[ColIdx].Value = data.Value.G1;

                    if (rowDa.Key == "二年級")
                        dgYearly.Rows[rowDa.Value].Cells[ColIdx].Value = data.Value.G2;

                    if (rowDa.Key == "三年級")
                        dgYearly.Rows[rowDa.Value].Cells[ColIdx].Value = data.Value.G3;

                    if (rowDa.Key == "四年級")
                        dgYearly.Rows[rowDa.Value].Cells[ColIdx].Value = data.Value.G4;

                    if (rowDa.Key == "五年級")
                        dgYearly.Rows[rowDa.Value].Cells[ColIdx].Value = data.Value.G5;

                    if (rowDa.Key == "六年級")
                        dgYearly.Rows[rowDa.Value].Cells[ColIdx].Value = data.Value.G6;
                }
            }            
            dgYearly.AllowUserToAddRows = false;
        }
    }
}
