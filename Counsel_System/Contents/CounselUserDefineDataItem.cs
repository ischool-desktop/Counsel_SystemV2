using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation;
using Framework;
using FCode = Framework.Security.FeatureCodeAttribute;
using K12.Data;

namespace Counsel_System2.Contents
{
    // 自訂資料欄位
    [FCode(PermissionCode.輔導自訂欄位_資料項目, "輔導自訂欄位")]
    public partial class CounselUserDefineDataItem : DetailContent
    {
        
        DAO.LogTransfer _LogTransfer;
        StudentRecord studRec;
        private BackgroundWorker _BGWorker;
        private ChangeListener ChangeManager = new ChangeListener();
        private bool _isBusy = false;
        // 放待刪除資料
        private List<DAO.UDT_CounselUserDefDataDef> _DeleteDataList;
        // 放待新新增資料
        private List<DAO.UDT_CounselUserDefDataDef> _InsertDataList;

        List<DAO.UDT_CounselUserDefDataDef> _UpdateDataList;

        private Dictionary<string, string> _UseDefineDataType;
        List<DAO.UDT_CounselUserDefDataDef> _UserDefineData;
        private List<string> _CheckSameList;
        List<string> _StudentIDList;
        List<string> _FieldNameList;

        // 給刪除用
        List<string> _HasUIDList;
        DAO.UDTTransfer _UDTTransfer;
      
        public CounselUserDefineDataItem()
        {
            InitializeComponent();
            _DeleteDataList = new List<DAO.UDT_CounselUserDefDataDef>();
            _InsertDataList = new List<DAO.UDT_CounselUserDefDataDef>();
            _UserDefineData = new List<DAO.UDT_CounselUserDefDataDef>();
            _UpdateDataList = new List<DAO.UDT_CounselUserDefDataDef>();
            _CheckSameList = new List<string>();
            _StudentIDList = new List<string>();
            _HasUIDList = new List<string>();
            _UDTTransfer = new DAO.UDTTransfer();
            _FieldNameList = new List<string>();
            _LogTransfer = new DAO.LogTransfer();
            Group = "輔導自訂欄位";
            _BGWorker = new BackgroundWorker();
            _BGWorker.DoWork += new DoWorkEventHandler(_BGWorker_DoWork);
            _BGWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWorker_RunWorkerCompleted);
            ChangeManager.Add(new DataGridViewSource(dgv));
            ChangeManager.StatusChanged += delegate(object sender, ChangeEventArgs e)
            {
                this.CancelButtonVisible = (e.Status == ValueStatus.Dirty);
                this.SaveButtonVisible = (e.Status == ValueStatus.Dirty);
            };
        }

        void _BGWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_isBusy)
            {
                _isBusy = false;
                _BGWorker.RunWorkerAsync();
                return;
            }
            DataBindToDataGridView();
        }

        void _BGWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _StudentIDList.Clear();
            _StudentIDList.Add(PrimaryKey);
            // 取得使用者設定欄位型態
            _UseDefineDataType = Global.GetUserConfigData();
            _UserDefineData = _UDTTransfer.GetCounselUserDefineDataByStudentIDList(_StudentIDList);           
        }

        /// <summary>
        /// 將讀取資料填入DataGridView
        /// </summary>
        private void DataBindToDataGridView()
        {
            try
            {
                //_LogTransfer.Clear();
                this.Loading = true;
                ChangeManager.SuspendListen();
                dgv.Rows.Clear();
                _FieldNameList.Clear();                
                int rowIdx=0;
                foreach(DAO.UDT_CounselUserDefDataDef udd in _UserDefineData)                
                {
                    rowIdx = dgv.Rows.Add();
                    dgv.Rows[rowIdx].Tag = udd;
                    dgv.Rows[rowIdx].Cells[FieldName.Index].Value = udd.FieldName;
                    dgv.Rows[rowIdx].Cells[Value.Index].Value = udd.Value;      
                    _FieldNameList.Add(udd.FieldName);
                    //_LogTransfer.SetLogValue(udd.FieldName, udd.Value);
                }

                // 放入樣板有內容沒有的
                foreach(string str in _UseDefineDataType.Keys )
                    if (!_FieldNameList.Contains(str))
                    {
                        DAO.UDT_CounselUserDefDataDef cudd = new DAO.UDT_CounselUserDefDataDef();
                        cudd.FieldName = str;
                        cudd.Value = "";
                        cudd.StudentID = int.Parse(PrimaryKey);
                        rowIdx = dgv.Rows.Add();
                        dgv.Rows[rowIdx].Tag = cudd;
                        dgv.Rows[rowIdx].Cells[FieldName.Index].Value = cudd.FieldName;
                        dgv.Rows[rowIdx].Cells[Value.Index].Value = cudd.Value;
                    }
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show(ex.Message);
            }

            this.CancelButtonVisible = false;
            this.SaveButtonVisible = false;
            this.ContentValidated = true;

            ChangeManager.Reset();
            ChangeManager.ResumeListen();
            this.Loading = false;
        }

        /// <summary>
        /// 按下儲存
        /// </summary>
        /// <param name="e"></param>
        protected override void OnSaveButtonClick(EventArgs e)
        {
            bool canSave = true;
            dgv.EndEdit();
            // 檢查資料
            foreach (DataGridViewRow drv in dgv.Rows)
            {
                foreach (DataGridViewCell cell in drv.Cells)
                    if (cell.ErrorText != "")
                        canSave = false;
            }


            if (canSave)
            {
                try
                {
                    _InsertDataList.Clear();
                    _UpdateDataList.Clear();
                    _DeleteDataList.Clear();
                    _HasUIDList.Clear();

                    // 儲存資料到 UDT
                    foreach (DataGridViewRow row in dgv.Rows)
                    {
                        if (row.IsNewRow)
                            continue;

                        DAO.UDT_CounselUserDefDataDef udd = new DAO.UDT_CounselUserDefDataDef();

                        // 資料轉型
                        if (row.Tag != null)
                            udd = (DAO.UDT_CounselUserDefDataDef)row.Tag;

                        udd.StudentID = int.Parse(PrimaryKey);
                        if (row.Cells[FieldName.Index].Value != null)
                            udd.FieldName = row.Cells[FieldName.Index].Value.ToString();

                        if (row.Cells[Value.Index].Value != null)
                            udd.Value = row.Cells[Value.Index].Value.ToString();

                        // 新增或更新
                        if (string.IsNullOrEmpty(udd.UID))
                            _InsertDataList.Add(udd);
                        else
                        {
                            _UpdateDataList.Add(udd);
                            _HasUIDList.Add(udd.UID);
                        }
                        //_LogTransfer.SetLogValue(udd.FieldName, udd.Value);
                    }




                    // 新增或更新至 UDT
                    if (_InsertDataList.Count > 0)
                        _UDTTransfer.InsertCounselUsereDefinfDataList(_InsertDataList);

                    if (_UpdateDataList.Count > 0)
                        _UDTTransfer.UpdateCounselUserDefineDataList(_UpdateDataList);



                    // 刪除舊資料 UDT
                    foreach (DAO.UDT_CounselUserDefDataDef udd in _UserDefineData)
                        if (!_HasUIDList.Contains(udd.UID))
                            _DeleteDataList.Add(udd);

                    if (_DeleteDataList.Count > 0)
                        _UDTTransfer.DeleteCounselUserDefineDataList(_DeleteDataList);

                    this.CancelButtonVisible = false;
                    this.SaveButtonVisible = false;
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("儲存失敗!");
                }
            }
            else
            {
                this.CancelButtonVisible = true;
                this.SaveButtonVisible = false;
            
            }
        }

        /// <summary>
        /// 更換學生
        /// </summary>
        /// <param name="e"></param>
        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            this.Loading = true;
            this.CancelButtonVisible = false;
            this.SaveButtonVisible = false;

            if (_BGWorker.IsBusy)
            {
                _isBusy = true;
                return;
            }
            else
                _BGWorker.RunWorkerAsync();    
        }

        /// <summary>
        /// 按下取消
        /// </summary>
        /// <param name="e"></param>
        protected override void OnCancelButtonClick(EventArgs e)
        {
            DataBindToDataGridView();
        }

        private void dgv_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dgv.EndEdit();
            bool validated = true;
            _CheckSameList.Clear();

            this.SaveButtonVisible = true;
            this.CancelButtonVisible = true;

            // 檢查資料
            foreach (DataGridViewRow row in dgv.Rows)
            {
                // 清空錯誤訊息
                row.Cells[FieldName.Index].ErrorText = "";
                row.Cells[Value.Index].ErrorText = "";
                SaveButtonVisible = true;
                CancelButtonVisible = true;


                string FName = string.Empty;
                if (row.IsNewRow)
                    continue;
                decimal dd;
                DateTime dt;

                if (row.Cells[FieldName.Index].Value != null)
                    FName = row.Cells[FieldName.Index].Value.ToString();

                if (FName != string.Empty)
                {
                    //if (_UseDefineDataType.ContainsKey(FName))
                    //{
                    //    if (row.Cells[Value.Index].Value == null)
                    //        continue;

                    //    if (row.Cells[Value.Index].Value.ToString() == string.Empty)
                    //        continue;

                    //    string str = row.Cells[Value.Index].Value.ToString();
                    //    if (_UseDefineDataType[FName] == "Number")
                    //    {
                    //        if (!decimal.TryParse(str, out dd))
                    //        {
                    //            row.Cells[Value.Index].ErrorText = "非數字型態";
                    //            validated = false;
                    //        }
                    //    }

                    //    if (_UseDefineDataType[FName] == "Date")
                    //    {
                    //        if (!DateTime.TryParse(str, out dt))
                    //        {
                    //            row.Cells[Value.Index].ErrorText = "非日期型態";
                    //            validated = false;
                    //        }
                    //    }
                    //}
                }
                else
                {
                    row.Cells[FieldName.Index].ErrorText = "不允許空白";
                    validated = false;
                }
            }

            dgv.BeginEdit(false);
            this.ContentValidated = validated;

        }

        private void dgv_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            bool validated = true;
            _CheckSameList.Clear();
            this.SaveButtonVisible = true;
            this.CancelButtonVisible = true;
            // 檢查資料
            foreach (DataGridViewRow row in dgv.Rows)
            {
                string FName = string.Empty;
                if (row.IsNewRow)
                    continue;

                if (row.Cells[FieldName.Index].Value != null)
                    FName = row.Cells[FieldName.Index].Value.ToString();

                if (_CheckSameList.Contains(FName))
                {
                    row.Cells[FieldName.Index].ErrorText = "欄位名稱重複";
                    validated = false;
                }

                _CheckSameList.Add(FName);
            }

            foreach (DataGridViewRow row in dgv.Rows)
                foreach (DataGridViewCell cell in row.Cells )
                    if (cell.ErrorText != string.Empty)
                    {
                        validated = false;
                        break;
                    }
            this.ContentValidated = validated;

        }

        private void dgv_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
                SaveButtonVisible = true;
                CancelButtonVisible = true;
                this.ContentValidated = true;
        }

    }
}
