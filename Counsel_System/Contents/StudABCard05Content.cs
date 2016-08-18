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
using DevComponents.DotNetBar.Controls;

namespace Counsel_System2.Contents
{
    [FISCA.Permission.FeatureCode(PermissionCode.綜合表現紀錄表_資料項目, "綜合表現紀錄表-自我認識")]
    public partial class StudABCard05Content :FISCA.Presentation.DetailContent
    {
        Dictionary<string, UDTSingleRecordDef> _dataDict;
        private BackgroundWorker _bgWorker;
        bool _isBusy = false;
        List<string> _StudentIDList;
        int _intStudentID = 0;
        ChangeListener _ChangeListener = new ChangeListener();
        List<UDTSingleRecordDef> _insertDataList;
        List<UDTSingleRecordDef> _updateDataList;
        TextBox _txtdateV;
        string GroupName = "自我認識";
        string _dtName = "自我認識_填寫日期";
        string _txt01Name = "自我認識_個性";
        string _txt02Name = "自我認識_優點";
        string _txt03Name = "自我認識_需要改進的地方";


        public StudABCard05Content()
        {
            InitializeComponent();            
            this.Group = "綜合表現紀錄表-自我認識";
            _insertDataList = new List<UDTSingleRecordDef>();
            _updateDataList = new List<UDTSingleRecordDef>();
            _dataDict = new Dictionary<string, UDTSingleRecordDef>();
            _StudentIDList = new List<string>();
            _txtdateV = new TextBox();
            _txtdateV.Visible = false;
            cbxGradeYear.Items.AddRange(Utility.GetClassGradeYear().ToArray());

            // 預設未選年級前無法使用
            SetControlEnable(false);
            _ChangeListener.StatusChanged += new EventHandler<ChangeEventArgs>(_ChangeListener_StatusChanged);
            _ChangeListener.Add(new TextBoxSource(txt01));
            _ChangeListener.Add(new TextBoxSource(txt02));
            _ChangeListener.Add(new TextBoxSource(txt03));
            _ChangeListener.Add(new TextBoxSource(_txtdateV));
            
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);
        }

        void _ChangeListener_StatusChanged(object sender, ChangeEventArgs e)
        {
            CancelButtonVisible=(e.Status == ValueStatus.Dirty);
            SaveButtonVisible = (e.Status == ValueStatus.Dirty);
            if (CancelButtonVisible)
                cbxGradeYear.Enabled = false;
            else
                cbxGradeYear.Enabled = true;

        }

        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            _StudentIDList.Clear();
            _StudentIDList.Add(PrimaryKey);
            _intStudentID = int.Parse(PrimaryKey);
            cbxGradeYear.Text = "";
            SetDefaultControlName();
            SetControlEnable(false);
            _BGRun();
        }

        protected override void OnSaveButtonClick(EventArgs e)
        {
            SetData();
            if (_insertDataList.Count > 0)
                UDTTransfer.ABUDTSingleRecordInsert(_insertDataList);

            if (_updateDataList.Count > 0)
                UDTTransfer.ABUDTSingleRecordUpdate(_updateDataList);

            this.CancelButtonVisible = this.SaveButtonVisible = false;
            cbxGradeYear.Enabled = true;
            _BGRun();
        }

        protected override void OnCancelButtonClick(EventArgs e)
        {
            this.CancelButtonVisible = this.SaveButtonVisible = false;
            cbxGradeYear.Enabled = true;
            _BGRun();
        }

        /// <summary>
        /// 設定控制項初始名稱
        /// </summary>
        private void SetDefaultControlName()
        {
            txt01.Name = "txt01";
            txt02.Name = "txt02";
            txt03.Name = "txt03";
            dtFillDate.Name = "dtFill";
        }


        private void LoadData()
        {
            _ChangeListener.SuspendListen();
            ClearTxt();
            foreach (Control cr in this.Controls)
            {
                if (cr is TextBox)
                {
                    if (_dataDict.ContainsKey(cr.Name))
                    {
                        cr.Text = _dataDict[cr.Name].Data;
                        cr.Tag = _dataDict[cr.Name];
                    }
                }
            }

            if (_dataDict.ContainsKey(dtFillDate.Name))
            {
                DateTime dt;
                dtFillDate.Tag = _dataDict[dtFillDate.Name];
                if (DateTime.TryParse(_dataDict[dtFillDate.Name].Data, out dt))
                {
                    dtFillDate.Value = dt;
                }
                else
                    dtFillDate.Text = "";
            }
            _ChangeListener.Reset();
            _ChangeListener.ResumeListen();
            this.Loading = false;
        }

        private void ClearTxt()
        {            
            _txtdateV.Text= txt03.Text=txt02.Text=txt01.Text = "";
            dtFillDate.Text = "";
        }

        private void SetData()
        {
            _insertDataList.Clear();
            _updateDataList.Clear();
            foreach (Control cr in this.Controls)
            {
                if (cr is TextBox && cr.Enabled == true)
                {
                    UDTSingleRecordDef data = cr.Tag as UDTSingleRecordDef;
                    if (data == null)
                    {
                        data = new UDTSingleRecordDef();
                        data.Key = cr.Name;
                        data.StudentID = _intStudentID;
                    }
                    data.Data = cr.Text;
                    if (string.IsNullOrEmpty(data.UID))
                        _insertDataList.Add(data);
                    else
                        _updateDataList.Add(data);
                }
            }

            UDTSingleRecordDef uDT = dtFillDate.Tag as UDTSingleRecordDef;
            if (uDT == null)
            {
                uDT = new UDTSingleRecordDef();
                uDT.StudentID = _intStudentID;
                uDT.Key = dtFillDate.Name;
            }

            if (dtFillDate.IsEmpty)
                uDT.Data = "";
            else
                uDT.Data = dtFillDate.Value.ToShortDateString();

            if (string.IsNullOrEmpty(uDT.UID))
                _insertDataList.Add(uDT);
            else
                _updateDataList.Add(uDT);
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
            _dataDict.Clear();
            foreach (UDTSingleRecordDef data in UDTTransfer.ABUDTSingleRecordSelectByStudentIDList(_StudentIDList))
            {
                if (data.Key.Contains(GroupName))
                {
                    if (!_dataDict.ContainsKey(data.Key))
                        _dataDict.Add(data.Key, data);
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

        private void SetControlEnable(bool tf)
        {
            txt01.Enabled = txt02.Enabled = txt03.Enabled = tf;
            dtFillDate.Enabled = tf;        
        }

        private void cbxGradeYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(cbxGradeYear.Text))
            {
                ClearTxt();
                SetControlEnable(false);
            }
            else
            {
                SetControlEnable(true);
                dtFillDate.Name = _dtName + "_" + cbxGradeYear.Text;
                txt01.Name = _txt01Name + "_" + cbxGradeYear.Text;
                txt02.Name = _txt02Name + "_" + cbxGradeYear.Text;
                txt03.Name = _txt03Name + "_" + cbxGradeYear.Text;
                dtFillDate.Tag = null;
                txt01.Tag = txt02.Tag = txt03.Tag = null;

                _BGRun();
            }
        }

        private void dtFillDate_TextChanged(object sender, EventArgs e)
        {
            _txtdateV.Text = dtFillDate.Text;
        }

        private void StudABCard05Content_Load(object sender, EventArgs e)
        {

        }


    }
}
