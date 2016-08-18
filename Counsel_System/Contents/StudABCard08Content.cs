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
    [FISCA.Permission.FeatureCode(PermissionCode.綜合表現紀錄表_資料項目, "綜合表現紀錄表-備註")]
    public partial class StudABCard08Content : FISCA.Presentation.DetailContent
    {
        private BackgroundWorker _bgWorker;
        bool _isBusy = false;
        ChangeListener _ChangeListener = new ChangeListener();
        List<string> _StudentIDList;
        int _intStudentID = 0;
        UDTSingleRecordDef _data = null;
        List<UDTSingleRecordDef> _insertDataList;
        List<UDTSingleRecordDef> _updateDataList;
        string GroupName = "備註";
        string NewKey=""; 

        public StudABCard08Content()
        {
            InitializeComponent();
            NewKey = GroupName + "_備註";
            this.Group = "綜合表現紀錄表-備註";
            _StudentIDList = new List<string>();
            _insertDataList = new List<UDTSingleRecordDef>();
            _updateDataList = new List<UDTSingleRecordDef>();
            _ChangeListener.Add(new TextBoxSource(txtRemark));
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);
            _ChangeListener.StatusChanged += new EventHandler<ChangeEventArgs>(_ChangeListener_StatusChanged);
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

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _data = null;
            foreach (UDTSingleRecordDef data in UDTTransfer.ABUDTSingleRecordSelectByStudentIDList(_StudentIDList))
            {             
                if (data.Key.Contains(GroupName)||data.Key.Contains((NewKey)))                   
                    _data = data;
            }
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

        private void LoadData()
        {
            _ChangeListener.SuspendListen();
            txtRemark.Text = "";
            txtRemark.Tag = _data;

            if (_data != null)
                txtRemark.Text = _data.Data;

            _ChangeListener.Reset();
            _ChangeListener.ResumeListen();
            this.Loading = false;
        }


        private void SetData()
        {
            _insertDataList.Clear();
            _updateDataList.Clear();

            UDTSingleRecordDef uDT = txtRemark.Tag as UDTSingleRecordDef;
            if (uDT == null)
            {
                uDT = new UDTSingleRecordDef();
                uDT.StudentID = _intStudentID;
                uDT.Key = NewKey;
            }

            uDT.Data = txtRemark.Text;

            if (string.IsNullOrEmpty(uDT.UID))
                _insertDataList.Add(uDT);
            else
                _updateDataList.Add(uDT);
        }

        protected override void OnSaveButtonClick(EventArgs e)
        {
            SetData();
            if (_insertDataList.Count > 0)
                UDTTransfer.ABUDTSingleRecordInsert(_insertDataList);

            if (_updateDataList.Count > 0)
                UDTTransfer.ABUDTSingleRecordUpdate(_updateDataList);

            this.CancelButtonVisible = this.SaveButtonVisible = false;
            _BGRun();
        }

        protected override void OnCancelButtonClick(EventArgs e)
        {
            this.CancelButtonVisible = this.SaveButtonVisible = false;            
            _BGRun();
        }

    }
}
