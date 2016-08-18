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
    [FISCA.Permission.FeatureCode(PermissionCode.綜合表現紀錄表_資料項目, "綜合表現紀錄表-適應情形")]
    public partial class StudABCard09Content : FISCA.Presentation.DetailContent
    {
        private BackgroundWorker _bgWorker;
        bool _isBusy = false;
        bool _reloadQuestion = true;
        List<string> _StudentIDList;
        int _intStudentID = 0;
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


        ChangeListener _ChangeListener = new ChangeListener();
        
        List<UDTYearlyDataDef> _insertDataList;
        List<UDTYearlyDataDef> _updateDataList;
        string GroupName = "適應情形";
        Dictionary<string, UDTYearlyDataDef> _UDTYearlyDataDict;
        Dictionary<string, QuestionData> _QuestionDataDict;
        ABCardQuestionDataManager _QDMang;
        public StudABCard09Content()
        {
            InitializeComponent();
            this.Group = "綜合表現紀錄表-適應情形";
            _StudentIDList = new List<string>();
            _insertDataList = new List<UDTYearlyDataDef>();
            _updateDataList = new List<UDTYearlyDataDef>();
            
            _QDMang = new ABCardQuestionDataManager();
            // 讀取題目
            _QuestionDataDict = _QDMang.GetQuestionDataByGroupName(GroupName);
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
            _ColumIndexDict.Add(GroupName+"_生活習慣", 1);
            _ColumIndexDict.Add(GroupName + "_人際關係", 2);
            _ColumIndexDict.Add(GroupName + "_外向行為", 3);
            _ColumIndexDict.Add(GroupName + "_內向行為", 4);
            _ColumIndexDict.Add(GroupName + "_學習動機", 5);
            _ColumIndexDict.Add(GroupName + "_服務熱忱", 6);
            _ColumIndexDict.Add(GroupName + "_人生態度", 7);

            _bgWorker = new BackgroundWorker();
            _UDTYearlyDataDict = new Dictionary<string, UDTYearlyDataDef>();
            _ChangeListener.Add(new DataGridViewSource(dgData));
            _ChangeListener.StatusChanged += new EventHandler<ChangeEventArgs>(_ChangeListener_StatusChanged);
            EventHub.CounselChanged += new EventHandler(EventHub_CounselChanged);
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);
        }

        void EventHub_CounselChanged(object sender, EventArgs e)
        {
            _reloadQuestion = true;
        }

        void _ChangeListener_StatusChanged(object sender, ChangeEventArgs e)
        {
            CancelButtonVisible = (e.Status == ValueStatus.Dirty);
            SaveButtonVisible = (e.Status == ValueStatus.Dirty);
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_isBusy)
            {
                _isBusy = false;
                _bgWorker.RunWorkerAsync();
                return;
            }
            LoadData();
        }

        private void LoadData()
        {
            _ChangeListener.SuspendListen();
            dgData.Rows.Clear();
            foreach (DataGridViewColumn col in dgData.Columns)
                col.Tag = null;
            dgData.AllowUserToAddRows = true;

            int Row = 0;
            // 放入年級
            foreach (string str in _RowNameList)
            {
                Row = dgData.Rows.Add();
                dgData.Rows[Row].Cells[0].Value = str;
            }

            // 放入選項
            colData1.Items.Clear();
            colData2.Items.Clear();
            colData3.Items.Clear();
            colData4.Items.Clear();
            colData5.Items.Clear();
            colData6.Items.Clear();
            colData7.Items.Clear();

            //生活習慣
            if (_QuestionDataDict.ContainsKey("生活習慣"))
                colData1.Items.AddRange((from data in _QuestionDataDict["生活習慣"].itemList select data.Key).ToArray());

            //人際關係
            if (_QuestionDataDict.ContainsKey("人際關係"))
                colData2.Items.AddRange((from data in _QuestionDataDict["人際關係"].itemList select data.Key).ToArray());

            //外向行為
            if (_QuestionDataDict.ContainsKey("外向行為"))
                colData3.Items.AddRange((from data in _QuestionDataDict["外向行為"].itemList select data.Key).ToArray());

            //內向行為
            if (_QuestionDataDict.ContainsKey("內向行為"))
                colData4.Items.AddRange((from data in _QuestionDataDict["內向行為"].itemList select data.Key).ToArray());

            //學習動機
            if (_QuestionDataDict.ContainsKey("學習動機"))
                colData5.Items.AddRange((from data in _QuestionDataDict["學習動機"].itemList select data.Key).ToArray());

            //服務熱忱
            if (_QuestionDataDict.ContainsKey("服務熱忱"))
                colData6.Items.AddRange((from data in _QuestionDataDict["服務熱忱"].itemList select data.Key).ToArray());

            //人生態度
            if (_QuestionDataDict.ContainsKey("人生態度"))
                colData7.Items.AddRange((from data in _QuestionDataDict["人生態度"].itemList select data.Key).ToArray());


            // 填入答案
            foreach (KeyValuePair<string, UDTYearlyDataDef> data in _UDTYearlyDataDict)
            {
                int ColIdx = 0;
                if (_ColumIndexDict.ContainsKey(data.Key))
                    ColIdx = _ColumIndexDict[data.Key];

                // 內容值放入 Tag
                dgData.Columns[ColIdx].Tag = data.Value;

                foreach (KeyValuePair<string, int> rowDa in _RowIndexDict)
                {
                    if (rowDa.Key == "一年級")
                        dgData.Rows[rowDa.Value].Cells[ColIdx].Value = data.Value.G1;

                    if (rowDa.Key == "二年級")
                        dgData.Rows[rowDa.Value].Cells[ColIdx].Value = data.Value.G2;

                    if (rowDa.Key == "三年級")
                        dgData.Rows[rowDa.Value].Cells[ColIdx].Value = data.Value.G3;

                    if (rowDa.Key == "四年級")
                        dgData.Rows[rowDa.Value].Cells[ColIdx].Value = data.Value.G4;

                    if (rowDa.Key == "五年級")
                        dgData.Rows[rowDa.Value].Cells[ColIdx].Value = data.Value.G5;

                    if (rowDa.Key == "六年級")
                        dgData.Rows[rowDa.Value].Cells[ColIdx].Value = data.Value.G6;
                }
            }

            dgData.AllowUserToAddRows = false;
            _ChangeListener.Reset();
            _ChangeListener.ResumeListen();
            this.Loading = false;
            _reloadQuestion = false;
        
        }

        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            _StudentIDList.Clear();
            _StudentIDList.Add(PrimaryKey);
            _intStudentID = int.Parse(PrimaryKey);
            _BGRun();
        }

        /// <summary>
        /// 呼叫重新讀取資料
        /// </summary>
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

        private void SetData()
        {
            _insertDataList.Clear();
            _updateDataList.Clear();
            foreach (DataGridViewColumn drc in dgData.Columns)
            {
                string key = GroupName + "_" + drc.HeaderText;
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
                            if (dgData.Rows[rowDa.Value].Cells[drc.Index].Value != null)
                                data.G1 = dgData.Rows[rowDa.Value].Cells[drc.Index].Value.ToString();
                            else
                                data.G1 = "";
                        }

                        if (rowDa.Key == "二年級")
                        {
                            if (dgData.Rows[rowDa.Value].Cells[drc.Index].Value != null)
                                data.G2 = dgData.Rows[rowDa.Value].Cells[drc.Index].Value.ToString();
                            else
                                data.G2 = "";
                        }

                        if (rowDa.Key == "三年級")
                        {
                            if (dgData.Rows[rowDa.Value].Cells[drc.Index].Value != null)
                                data.G3 = dgData.Rows[rowDa.Value].Cells[drc.Index].Value.ToString();
                            else
                                data.G3 = "";
                        }

                        if (rowDa.Key == "四年級")
                        {
                            if (dgData.Rows[rowDa.Value].Cells[drc.Index].Value != null)
                                data.G4 = dgData.Rows[rowDa.Value].Cells[drc.Index].Value.ToString();
                            else
                                data.G4 = "";
                        }

                        if (rowDa.Key == "五年級")
                        {
                            if (dgData.Rows[rowDa.Value].Cells[drc.Index].Value != null)
                                data.G5 = dgData.Rows[rowDa.Value].Cells[drc.Index].Value.ToString();
                            else
                                data.G5 = "";
                        }

                        if (rowDa.Key == "六年級")
                        {
                            if (dgData.Rows[rowDa.Value].Cells[drc.Index].Value != null)
                                data.G6 = dgData.Rows[rowDa.Value].Cells[drc.Index].Value.ToString();
                            else
                                data.G6 = "";
                        }
                    }
                    if (string.IsNullOrEmpty(data.UID))
                        _insertDataList.Add(data);
                    else
                        _updateDataList.Add(data);
                }
            }        
        }

        protected override void OnSaveButtonClick(EventArgs e)
        {
            dgData.EndEdit();
            SetData();
            if (_insertDataList.Count > 0)
                UDTTransfer.ABUDTYearlyDataInsert(_insertDataList);

            if (_updateDataList.Count > 0)
                UDTTransfer.ABUDTYearlyDataUpdate(_updateDataList);

            this.CancelButtonVisible = this.SaveButtonVisible = false;
            _BGRun();

        }

        protected override void OnCancelButtonClick(EventArgs e)
        {
            this.CancelButtonVisible = this.SaveButtonVisible = false;
            _BGRun();
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if (_reloadQuestion)
                _QuestionDataDict = _QDMang.GetQuestionDataByGroupName(GroupName);

                _UDTYearlyDataDict.Clear();
                foreach (UDTYearlyDataDef data in UDTTransfer.ABUDTYearlyDataSelectByStudentIDList(_StudentIDList))
                {
                    if (data.Key.Contains(GroupName))
                        if (!_UDTYearlyDataDict.ContainsKey(data.Key))
                            _UDTYearlyDataDict.Add(data.Key, data);
                }
                        
        }
    }
}
