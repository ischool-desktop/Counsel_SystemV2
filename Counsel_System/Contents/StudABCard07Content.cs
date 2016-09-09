using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Campus.Windows;
using DevComponents.DotNetBar.Controls;
using Counsel_System2.DAO;

namespace Counsel_System2.Contents
{
    [FISCA.Permission.FeatureCode(PermissionCode.綜合表現紀錄表_資料項目, "綜合表現紀錄表-畢業後計畫")]
    public partial class StudABCard07Content : FISCA.Presentation.DetailContent
    {
        private BackgroundWorker _bgWorkerM;        
        bool _isBusyM = false;

        bool _reloadQuestion = true;
        Dictionary<string, UDTMultipleRecordDef> _dataMDict;
        Dictionary<string, UDTPriorityDataDef> _dataPDict;

        UDTPriorityDataDef _PriorityData1;
        UDTPriorityDataDef _PriorityData2;

        int _RowIdx1 = 0;
        int _RowIdx2 = 0;

        // 動態計算位置使用
        int y = 0;

        Dictionary<string, QuestionData> _QuestionDict;
        string GroupName = "畢業後計畫";
        string _keyName1 = "升學意願";
        string _keyName2 = "就業意願";
        string _keyName3 = "參加職業訓練";
        string _keyName4 = "受訓地區";
        string _keyName5 = "將來職業";
        string _keyName6 = "就業地區";

        List<string> _KeyNameList;
        List<FlowLayoutPanel> _flpList;

        List<string> _StudentIDList;
        int _intStudentID = 0;
        ChangeListener _ChangeListener = new ChangeListener();
        Dictionary<string, UDTMultipleRecordDef> _InsertMultipleRecordDict;
        List<UDTMultipleRecordDef> _deleteMultipleRecordList;
        List<UDTPriorityDataDef> _insertPriorityDataList;
        List<UDTPriorityDataDef> _updatePriorityDataList;
        
        ABCardQuestionDataManager _QDMang;
        public StudABCard07Content()
        {
            InitializeComponent();
            this.Group = "綜合表現紀錄表-畢業後計畫";
            _StudentIDList = new List<string>();
            _KeyNameList = new List<string>();
            _KeyNameList.Add(_keyName1);
            _KeyNameList.Add(_keyName2);
            _KeyNameList.Add(_keyName3);
            _KeyNameList.Add(_keyName4);
            _KeyNameList.Add(_keyName5);
            _KeyNameList.Add(_keyName6);
            _flpList = new List<FlowLayoutPanel>();
            _flpList.Add(flp1);
            _flpList.Add(flp2);
            _flpList.Add(flp3);
            _flpList.Add(flp4);

            _QDMang = new ABCardQuestionDataManager();          
            _PriorityData1 = new UDTPriorityDataDef();
            _PriorityData2 = new UDTPriorityDataDef();
            _QuestionDict = _QDMang.GetQuestionDataByGroupName(GroupName);
            _dataMDict = new Dictionary<string, UDTMultipleRecordDef>();
            _dataPDict = new Dictionary<string, UDTPriorityDataDef>();
            _insertPriorityDataList = new List<UDTPriorityDataDef>();
            _InsertMultipleRecordDict = new Dictionary<string, UDTMultipleRecordDef>();
            _updatePriorityDataList = new List<UDTPriorityDataDef>();
            _deleteMultipleRecordList = new List<UDTMultipleRecordDef>();
            _ChangeListener.StatusChanged += new EventHandler<ChangeEventArgs>(_ChangeListener_StatusChanged);
            _ChangeListener.Add(new DataGridViewSource(dgPriority));
            LoadQuestionToUI();
            EventHub.CounselChanged += new EventHandler(EventHub_CounselChanged);
            _bgWorkerM = new BackgroundWorker();            
            _bgWorkerM.DoWork += new DoWorkEventHandler(_bgWorkerM_DoWork);
            _bgWorkerM.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorkerM_RunWorkerCompleted);
        }

        void EventHub_CounselChanged(object sender, EventArgs e)
        {
            _reloadQuestion = true;
        }

        void _bgWorkerM_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_isBusyM)
            {
                _isBusyM = false;
                _bgWorkerM.RunWorkerAsync();
                return;
            }
            LoadData();
        }

        void _bgWorkerM_DoWork(object sender, DoWorkEventArgs e)
        {
            if (_reloadQuestion)            
                _QuestionDict = _QDMang.GetQuestionDataByGroupName(GroupName);
            
                _dataMDict.Clear();
                foreach (UDTMultipleRecordDef data in UDTTransfer.ABUDTMultipleRecordSelectByStudentIDList(_StudentIDList))
                {
                    if (data.Key.Contains(GroupName))
                    {
                        string key = data.Key + "_" + data.Data;
                        if (!_dataMDict.ContainsKey(key))
                            _dataMDict.Add(key, data);
                    }
                }

                _PriorityData1 = null;
                _PriorityData2 = null;
                string key1 = GroupName + "_" + _keyName5;
                string key2 = GroupName + "_" + _keyName6;
                foreach (UDTPriorityDataDef data in UDTTransfer.ABUDTPriorityDataSelectByStudentIDList(_StudentIDList))
                {
                    if (data.Key.Contains(GroupName))
                    {
                        if (key1 == data.Key)
                            _PriorityData1 = data;

                        if (key2 == data.Key)
                            _PriorityData2 = data;
                    }
                }
                       
        }

        /// <summary>
        /// 載入題目至畫面
        /// </summary>
        private void LoadQuestionToUI()
        {            
            for (int idx = 0; idx < 4; idx++)
            {
                List<CheckBox> cbx = new List<CheckBox>();
                FlowLayoutPanel flp = _flpList[idx];
                string kName = _KeyNameList[idx];

                flp.Controls.Clear();
                if (_QuestionDict.ContainsKey(kName))
                {                    
                    flp.FlowDirection = FlowDirection.LeftToRight;
                    foreach (QuestionItem qi in _QuestionDict[kName].itemList)
                    {
                        CheckBox cb = new CheckBox();
                        cb.Name = qi.Key;
                        if (qi.hasRemark)
                            cb.Text = qi.Key + "：";                            
                        else
                            cb.Text = qi.Key;
                        cb.AutoSize = true;
                        cbx.Add(cb);
                        flp.Controls.Add(cb);
                        if (qi.hasRemark)
                        {
                            TextBox tb = new TextBox();
                            tb.Name = qi.Key;
                            tb.Width = 70;
                            tb.Text = "";
                            flp.Controls.Add(tb);
                        }
                    }
                    _ChangeListener.Add(new CheckBoxSource(cbx.ToArray()));
                }               
            }           
            
            // 動態調整位置
            y = flp1.Location.Y;
            lbl01.Location = new Point(lbl01.Location.X, y + 4);
            y += flp1.Size.Height + 12;
            flp2.Location = new Point(flp2.Location.X, y);
            lbl02.Location = new Point(lbl02.Location.X, y + 4);

            y += flp2.Size.Height + 12;
            flp3.Location = new Point(flp3.Location.X, y);
            lbl03.Location = new Point(lbl03.Location.X, y + 4);

            y += flp3.Size.Height + 12;
            flp4.Location = new Point(flp4.Location.X, y);
            lbl04.Location = new Point(lbl04.Location.X, y + 4);

            y += flp4.Size.Height + 12;

            lbl05.Location = new Point(lbl05.Location.X, y);
            
            dgPriority.Location = new Point(dgPriority.Location.X, y);
            y += dgPriority.Size.Height;
            
            this.Size = new Size(this.Width, y + 10);

        }

        /// <summary>
        /// 當資料狀態有改變
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void _ChangeListener_StatusChanged(object sender, ChangeEventArgs e)
        {
            this.CancelButtonVisible = (e.Status == ValueStatus.Dirty);
            this.SaveButtonVisible = (e.Status == ValueStatus.Dirty);
            
        }

        /// <summary>
        /// 載入資料至控制項
        /// </summary>
        private void LoadData()
        {
            _ChangeListener.SuspendListen();
            if(_reloadQuestion)
                LoadQuestionToUI();

            SetControlsDefault();
            BindAnswerDataToUI();
            _ChangeListener.Reset();
            _ChangeListener.ResumeListen();
            this.Loading = false;
            _reloadQuestion = false;
        }

        /// <summary>
        /// 將答案填入畫面
        /// </summary>
        private void BindAnswerDataToUI()
        {
            for (int idx = 0; idx < 4; idx++)
            {   // 讀取每個 flp
                FlowLayoutPanel flp = _flpList[idx];
                string KeyName = _KeyNameList[idx];
                foreach (Control cr in flp.Controls)
                {   // 檢查是否是 CheckBox
                    if (cr is CheckBox)
                    {
                        CheckBox cb = cr as CheckBox;
                        // 建立 key 值 
                        string key =GroupName+"_"+ KeyName + "_" + cb.Name;
                        if (_dataMDict.ContainsKey(key))
                        {
                            cb.Checked = true;
                            // 當備註不是空白
                            if (!string.IsNullOrEmpty(_dataMDict[key].Remark))
                            {
                                foreach (Control t in flp.Controls)
                                    if (t is TextBox)
                                        if (t.Name == _dataMDict[key].Data)
                                            t.Text = _dataMDict[key].Remark;
                            }
                        }
                    }

                }
            }
                       
            // 放入欄意願1
            dgPriority.Rows.Clear();
            for (int i = 1; i <= 3; i++)
            {
                string key = "意願" + i;
                int row = dgPriority.Rows.Add();
                dgPriority.Rows[row].Cells[0].Value = key;
            }

            // 放入選項
            dgColP1.Items.Clear();
            dgColP2.Items.Clear();
            // 將來就業意願
            if (_QuestionDict.ContainsKey(_keyName5))
                dgColP1.Items.AddRange((from data in _QuestionDict[_keyName5].itemList select data.Key).ToArray());
            // 就業地區
            if (_QuestionDict.ContainsKey(_keyName6))
                dgColP2.Items.AddRange((from data in _QuestionDict[_keyName6].itemList select data.Key).ToArray());

            // 將來就業意願
            if (_PriorityData1 != null)
            {
                dgPriority.Columns[1].Tag = _PriorityData1;                
                dgPriority.Rows[0].Cells[1].Value = _PriorityData1.P1;
                dgPriority.Rows[1].Cells[1].Value = _PriorityData1.P2;
                dgPriority.Rows[2].Cells[1].Value = _PriorityData1.P3;                              
            }

            // 就業地區
            if (_PriorityData2 != null)
            {
                dgPriority.Columns[2].Tag = _PriorityData2;                
                dgPriority.Rows[0].Cells[2].Value = _PriorityData2.P1;
                dgPriority.Rows[1].Cells[2].Value = _PriorityData2.P2;
                dgPriority.Rows[2].Cells[2].Value = _PriorityData2.P3;
            }
        }

        /// <summary>
        /// 控制項初始化
        /// </summary>
        private void SetControlsDefault()
        {
            foreach (Control cr in this.Controls)
            {
                if (cr is FlowLayoutPanel)
                {                    
                    foreach (Control c in cr.Controls)
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
            }
        }

        /// <summary>
        /// 設定儲存資料
        /// </summary>
        private void SetData()
        {
            // 刪除舊資料
            _deleteMultipleRecordList = _dataMDict.Values.ToList();
            UDTTransfer.ABUDTMultipleRecordDelete(_deleteMultipleRecordList);
            _InsertMultipleRecordDict.Clear();

            // 儲存資料
            for (int idx = 0; idx < 4; idx++)
            {
                FlowLayoutPanel flp = _flpList[idx];
                string KeyName = _KeyNameList[idx];

                foreach (Control cr in flp.Controls)
                {
                    if (cr is CheckBox)
                    {
                        CheckBox cb = cr as CheckBox;
                        if (cb.Checked)
                        {
                            UDTMultipleRecordDef data1 = new UDTMultipleRecordDef();
                            data1.StudentID = _intStudentID;
                            data1.Key = GroupName + "_" + KeyName;
                            data1.Data = cb.Name;
                            data1.Remark = "";

                            // 處理備註資料
                            foreach (Control t in flp.Controls)
                                if (t is TextBox)
                                    if (t.Name == cb.Name)
                                        data1.Remark = t.Text;

                            string insertKey = data1.Key + "_" + cb.Name;
                            if (!_InsertMultipleRecordDict.ContainsKey(insertKey))
                                _InsertMultipleRecordDict.Add(insertKey, data1);
                        }
                    }
                }
            }

            // 新增至 UDT
            if (_InsertMultipleRecordDict.Count > 0)
                UDTTransfer.ABUDTMultipleRecordInsert(_InsertMultipleRecordDict.Values.ToList());
        
            // 讀存資料
            List<UDTPriorityDataDef> insertData = new List<UDTPriorityDataDef>();
            List<UDTPriorityDataDef> updateData = new List<UDTPriorityDataDef>();

            // 將來就業
            UDTPriorityDataDef p1 = dgPriority.Columns[1].Tag as UDTPriorityDataDef;
            if (p1 == null)
            {
                p1 = new UDTPriorityDataDef();
                p1.Key = GroupName + "_" + _keyName5;
                p1.StudentID = _intStudentID;
            }
            if (dgPriority.Rows[0].Cells[1].Value != null)
                p1.P1 = dgPriority.Rows[0].Cells[1].Value.ToString();
            else
                p1.P1 = "";
            if (dgPriority.Rows[1].Cells[1].Value != null)
                p1.P2 = dgPriority.Rows[1].Cells[1].Value.ToString();
            else
                p1.P2 = "";
            if (dgPriority.Rows[2].Cells[1].Value != null)
                p1.P3 = dgPriority.Rows[2].Cells[1].Value.ToString();
            else
                p1.P3 = "";

            if (string.IsNullOrEmpty(p1.UID))
                insertData.Add(p1);
            else
                updateData.Add(p1);

            // 就業地區
            UDTPriorityDataDef p2 = dgPriority.Columns[2].Tag as UDTPriorityDataDef;
            if (p2 == null)
            {
                p2 = new UDTPriorityDataDef();
                p2.Key = GroupName+"_"+_keyName6;
                p2.StudentID = _intStudentID;
            }
            if (dgPriority.Rows[0].Cells[2].Value != null)
                p2.P1 = dgPriority.Rows[0].Cells[2].Value.ToString();
            else
                p2.P1 = "";
            if (dgPriority.Rows[1].Cells[2].Value != null)
                p2.P2 = dgPriority.Rows[1].Cells[2].Value.ToString();
            else
                p2.P2 = "";
            if (dgPriority.Rows[2].Cells[2].Value != null)
                p2.P3 = dgPriority.Rows[2].Cells[2].Value.ToString();
            else
                p2.P3 = "";

            if (string.IsNullOrEmpty(p2.UID))
                insertData.Add(p2);
            else
                updateData.Add(p2);

            if (insertData.Count > 0)
                UDTTransfer.ABUDTPriorityDataInsert(insertData);

            if (updateData.Count > 0)
                UDTTransfer.ABUDTPriorityDataUpdate(updateData);

        }

        /// <summary>
        /// 切換學生
        /// </summary>
        /// <param name="e"></param>
        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            _StudentIDList.Clear();
            _StudentIDList.Add(PrimaryKey);
            _intStudentID = int.Parse(PrimaryKey);
            _BGRunM();            
        }

        /// <summary>
        /// 按儲存按鈕
        /// </summary>
        /// <param name="e"></param>
        protected override void OnSaveButtonClick(EventArgs e)
        {
            dgPriority.EndEdit();
            SetData();
            this.CancelButtonVisible = this.SaveButtonVisible = false;
            _BGRunM();            
        }

        /// <summary>
        /// 按取消按鈕
        /// </summary>
        /// <param name="e"></param>
        protected override void OnCancelButtonClick(EventArgs e)
        {
            this.CancelButtonVisible = this.SaveButtonVisible = false;
            _BGRunM();            
        }

        /// <summary>
        /// 重新背景讀取資料
        /// </summary>
        private void _BGRunM()
        {
            if (_bgWorkerM.IsBusy)
                _isBusyM = true;
            else
            {
                this.Loading = true;
                _bgWorkerM.RunWorkerAsync();
            }
        }       
    }
}
