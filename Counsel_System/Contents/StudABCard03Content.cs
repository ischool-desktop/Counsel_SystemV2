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
    [FISCA.Permission.FeatureCode(PermissionCode.綜合表現紀錄表_資料項目, "綜合表現紀錄表-學習狀況")]
    public partial class StudABCard03Content : FISCA.Presentation.DetailContent
    {
        string GroupName = "學習狀況";
        
        List<string> _StudentIDList;
        int _intStudentID = 0;
        Dictionary<string, QuestionData> _QuestionDataDict;
        List<UDTYearlyDataDef> _insertDataList;
        List<UDTYearlyDataDef> _updateDataList;
        List<UDTSemesterDataDef> _insertSemesterDataList;
        List<UDTSemesterDataDef> _updateSemesterDataList;
        ABCardQuestionDataManager _QDMang;
        BackgroundWorker _bgWorker;
        List<string> _YearlyDataKeyList;
        List<string> _SemesterDataKeyList;
        Dictionary<string, UDTYearlyDataDef> _YearlyDataDict;
        Dictionary<string, UDTSemesterDataDef> _SemesterDataDict;
        string Key1 = "學習狀況_社團幹部";
        string Key2 = "學習狀況_班級幹部";
        private ChangeListener _ChangeListener = new ChangeListener();
        bool _isBusy = false;
        bool _reloadQuestion = true;
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
        
        public StudABCard03Content()
        {
            InitializeComponent();

            this.Group = "綜合表現紀錄表-學習狀況";
            _QDMang = new ABCardQuestionDataManager();
            _StudentIDList = new List<string>();
            _YearlyDataKeyList = new List<string> ();            
            _YearlyDataDict = new Dictionary<string, UDTYearlyDataDef>();
            _bgWorker = new BackgroundWorker();
            _insertDataList = new List<UDTYearlyDataDef>();
            _updateDataList = new List<UDTYearlyDataDef>();
            _insertSemesterDataList = new List<UDTSemesterDataDef>();
            _updateSemesterDataList = new List<UDTSemesterDataDef>();
            _QuestionDataDict = new Dictionary<string, QuestionData>();
            _SemesterDataDict = new Dictionary<string, UDTSemesterDataDef>();
            _SemesterDataKeyList = new List<string>();
            _SemesterDataKeyList.Add(Key1);
            _SemesterDataKeyList.Add(Key2);
            _ClassGradeYearDict = Utility.GetClassGradeYearDict();

            List<string> grYear = Utility.GetClassGradeYear();

            if (grYear.Count > 3)
                SetDgColumn4_6Visable(true);
            else
                SetDgColumn4_6Visable(false);

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
            _ColumIndexDict.Add(GroupName + "_特殊專長", 1);
            _ColumIndexDict.Add(GroupName + "_休閒興趣", 2);
            _ColumIndexDict.Add(GroupName + "_最喜歡的學科", 3);
            _ColumIndexDict.Add(GroupName + "_最感困難的學科", 4);
            _ColumIndexDict.Add(GroupName + "_特殊專長_樂器演奏", 5);
            _ColumIndexDict.Add(GroupName + "_特殊專長_外語能力", 6);
            //_ColumIndexDict.Add(GroupName + "_社團幹部", 5);
            //_ColumIndexDict.Add(GroupName + "_班級幹部", 6);

            lvData.FullRowSelect = true;
            lvData.MultiSelect = false;

            _ChangeListener.StatusChanged += new EventHandler<ChangeEventArgs>(_ChangeListener_StatusChanged);
            _ChangeListener.Add(new DataGridViewSource(dgData));
            _QuestionDataDict.Clear();
            _QuestionDataDict = _QDMang.GetQuestionDataByGroupName(GroupName);
            _YearlyDataKeyList = _QuestionDataDict.Keys.ToList();
            EventHub.CounselChanged += new EventHandler(EventHub_CounselChanged);
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);
        }

        void EventHub_CounselChanged(object sender, EventArgs e)
        {
            _reloadQuestion = true;
        }

        /// 設定DG四~六年級是否顯示
        /// </summary>
        /// <param name="bol"></param>
        private void SetDgColumn4_6Visable(bool bol)
        {
            col07.Visible = col08.Visible = col09.Visible = col10.Visible = col11.Visible = col12.Visible = bol;
        }

        void _ChangeListener_StatusChanged(object sender, ChangeEventArgs e)
        {
            this.CancelButtonVisible = (e.Status == ValueStatus.Dirty);
            this.SaveButtonVisible = (e.Status == ValueStatus.Dirty);
        }

        protected override void OnSaveButtonClick(EventArgs e)
        {
            dgData.EndEdit();
            SaveDgDataToUDT();
            this.CancelButtonVisible = false;
            this.SaveButtonVisible = false;
            _BGRun();
        }


        protected override void OnCancelButtonClick(EventArgs e)
        {
            this.CancelButtonVisible = false;
            this.SaveButtonVisible = false;
            LoadData();
        }

        private void SaveDgDataToUDT()
        {
            _insertSemesterDataList.Clear();
            _updateSemesterDataList.Clear();
            foreach (DataGridViewRow dr in dgData.Rows)
            {
                UDTSemesterDataDef data = dr.Tag as UDTSemesterDataDef;
                data.S1a = GetCellValue(dr.Cells[col01.Index]);
                data.S1b = GetCellValue(dr.Cells[col02.Index]);
                data.S2a = GetCellValue(dr.Cells[col03.Index]);
                data.S2b = GetCellValue(dr.Cells[col04.Index]);
                data.S3a = GetCellValue(dr.Cells[col05.Index]);
                data.S3b = GetCellValue(dr.Cells[col06.Index]);
                data.S4a = GetCellValue(dr.Cells[col07.Index]);
                data.S4b = GetCellValue(dr.Cells[col08.Index]);
                data.S5a = GetCellValue(dr.Cells[col09.Index]);
                data.S5b = GetCellValue(dr.Cells[col10.Index]);
                data.S6a = GetCellValue(dr.Cells[col11.Index]);
                data.S6b = GetCellValue(dr.Cells[col12.Index]);
                if (string.IsNullOrEmpty(data.UID))
                    _insertSemesterDataList.Add(data);
                else
                    _updateSemesterDataList.Add(data);
            }

            if (_insertSemesterDataList.Count > 0)
                UDTTransfer.ABUDTSemesterDataInsert(_insertSemesterDataList);

            if (_updateSemesterDataList.Count > 0)
                UDTTransfer.ABUDTSemesterDataUpdate(_updateSemesterDataList);
        }

        private string GetCellValue(DataGridViewCell cell)
        {
            string retVal = "";
            if (cell.Value != null)
                retVal = cell.Value.ToString();

            return retVal;
        }

        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            _StudentIDList.Clear();
            _StudentIDList.Add(PrimaryKey);
            _intStudentID = int.Parse(PrimaryKey);
            _BGRun();
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

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if (_reloadQuestion)            
                _QuestionDataDict = _QDMang.GetQuestionDataByGroupName(GroupName);

                _YearlyDataDict.Clear();
                foreach (UDTYearlyDataDef data in UDTTransfer.ABUDTYearlyDataSelectByStudentIDList(_StudentIDList))
                {
                    if (_ColumIndexDict.ContainsKey(data.Key))
                        if (!_YearlyDataDict.ContainsKey(data.Key))
                            _YearlyDataDict.Add(data.Key, data);
                }

                // 檢查資料是否完整,沒有資料補資料
                foreach (string str in _ColumIndexDict.Keys)
                {
                    if (!_YearlyDataDict.ContainsKey(str))
                    {
                        UDTYearlyDataDef data = new UDTYearlyDataDef();
                        data.Key = str;
                        data.StudentID = _intStudentID;
                        data.G1 = data.G2 = data.G3 = data.G4 = data.G5 = data.G6 = "";
                        _YearlyDataDict.Add(str, data);
                    }
                }

                // 社團與班級幹部
                _SemesterDataDict.Clear();
                foreach (UDTSemesterDataDef data in UDTTransfer.ABUDTSemesterDataSelectByStudentIDList(_StudentIDList))
                {
                    if (_SemesterDataKeyList.Contains(data.Key))
                        if (!_SemesterDataDict.ContainsKey(data.Key))
                            _SemesterDataDict.Add(data.Key, data);
                }

                // 初始值檢查,沒有建立初始值
                foreach (string str in _SemesterDataKeyList)
                {
                    if (!_SemesterDataDict.ContainsKey(str))
                    {
                        UDTSemesterDataDef data1 = new UDTSemesterDataDef();
                        data1.Key = str;
                        data1.StudentID = _intStudentID;
                        data1.S1a = data1.S1b = data1.S2a = data1.S2b = data1.S3a = data1.S3b = data1.S4a = data1.S4b = data1.S5a = data1.S5b = data1.S6a = data1.S6b = "";
                        _SemesterDataDict.Add(str, data1);
                    }
                }
                   
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

        private void LoadData()
        {
            lvData.Items.Clear();
            dgData.Rows.Clear();
            int row=0;
            foreach (string str in _RowNameList)
            {                
                lvData.Items.Add(str);
                foreach(string str1 in _ColumIndexDict.Keys)
                    lvData.Items[row].SubItems.Add("");
                row++;                    
            }
            foreach (KeyValuePair<string,UDTYearlyDataDef> data in _YearlyDataDict)
            {
                if (_ColumIndexDict.ContainsKey(data.Key))
                {
                    int colIdx = 0;
                    colIdx = _ColumIndexDict[data.Key];
                    lvData.Columns[colIdx].Tag = data.Value;

                    foreach (KeyValuePair<string, int> rowIdx in _RowIndexDict)
                    {                        
                        if (rowIdx.Key == "一年級")
                            lvData.Items[rowIdx.Value].SubItems[colIdx].Text = data.Value.G1;

                        if (rowIdx.Key == "二年級")
                            lvData.Items[rowIdx.Value].SubItems[colIdx].Text = data.Value.G2;

                        if (rowIdx.Key == "三年級")
                            lvData.Items[rowIdx.Value].SubItems[colIdx].Text = data.Value.G3;

                        if (rowIdx.Key == "四年級")
                            lvData.Items[rowIdx.Value].SubItems[colIdx].Text = data.Value.G4;

                        if (rowIdx.Key == "五年級")
                            lvData.Items[rowIdx.Value].SubItems[colIdx].Text = data.Value.G5;

                        if (rowIdx.Key == "六年級")
                            lvData.Items[rowIdx.Value].SubItems[colIdx].Text = data.Value.G6;
                    }                
                }            
            }

            _ChangeListener.SuspendListen();
            string rpStr = "學習狀況_";
            // 社團班級幹部
            foreach (KeyValuePair<string, UDTSemesterDataDef> data1 in _SemesterDataDict)
            {
                int rowIdx = dgData.Rows.Add();
                dgData.Rows[rowIdx].Tag = data1.Value;
                dgData.Rows[rowIdx].Cells[colItem.Index].Value = data1.Value.Key.Replace(rpStr, "");
                dgData.Rows[rowIdx].Cells[col01.Index].Value = data1.Value.S1a;
                dgData.Rows[rowIdx].Cells[col02.Index].Value = data1.Value.S1b;
                dgData.Rows[rowIdx].Cells[col03.Index].Value = data1.Value.S2a;
                dgData.Rows[rowIdx].Cells[col04.Index].Value = data1.Value.S2b;
                dgData.Rows[rowIdx].Cells[col05.Index].Value = data1.Value.S3a;
                dgData.Rows[rowIdx].Cells[col06.Index].Value = data1.Value.S3b;
                dgData.Rows[rowIdx].Cells[col07.Index].Value = data1.Value.S4a;
                dgData.Rows[rowIdx].Cells[col08.Index].Value = data1.Value.S4b;
                dgData.Rows[rowIdx].Cells[col09.Index].Value = data1.Value.S5a;
                dgData.Rows[rowIdx].Cells[col10.Index].Value = data1.Value.S5b;
                dgData.Rows[rowIdx].Cells[col11.Index].Value = data1.Value.S6a;
                dgData.Rows[rowIdx].Cells[col12.Index].Value = data1.Value.S6b;
            }

            _ChangeListener.Reset();
            _ChangeListener.ResumeListen();
            this.Loading = false;
            _reloadQuestion = false;
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (lvData.SelectedItems.Count == 1)
            {               
                // 傳入資料
                Forms.StudABCard_YearlyForm_03 form = new Forms.StudABCard_YearlyForm_03(_YearlyDataDict, lvData.SelectedItems[0].Text);

                if (_QuestionDataDict.ContainsKey("特殊專長"))                
                    form.SetFlp01Items((from data in _QuestionDataDict["特殊專長"].itemList select data.Key).ToList());

                if (_QuestionDataDict.ContainsKey("休閒興趣"))
                    form.SetFlp02Items((from data in _QuestionDataDict["休閒興趣"].itemList select data.Key).ToList());

                if (form.ShowDialog() == DialogResult.OK)
                {
                    List<UDTYearlyDataDef> insertData = new List<UDTYearlyDataDef>();
                    List<UDTYearlyDataDef> updateData = new List<UDTYearlyDataDef>();
                   
                    foreach (UDTYearlyDataDef data in form.GetData().Values)
                    {
                        if (string.IsNullOrEmpty(data.UID))
                            insertData.Add(data);
                        else
                            updateData.Add(data);                    
                    }

                    if (insertData.Count > 0)
                        UDTTransfer.ABUDTYearlyDataInsert(insertData);

                    if (updateData.Count > 0)
                        UDTTransfer.ABUDTYearlyDataUpdate(updateData);

                    _BGRun();
                }
            }
        }

    }
}
