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
    [FISCA.Permission.FeatureCode(PermissionCode.綜合表現紀錄表_資料項目, "綜合表現紀錄表-生活感想")]
    public partial class StudABCard06Content : FISCA.Presentation.DetailContent
    {
        private BackgroundWorker _bgWorker;
        bool _isBusy = false;
        ChangeListener _ChangeListener = new ChangeListener();
        Dictionary<string, UDTSingleRecordDef> _dataDict;
        List<string> _StudentIDList;
        int _intStudentID = 0;
        TextBox _txtdateV;                    
        List<UDTSingleRecordDef> _insertDataList;
        List<UDTSingleRecordDef> _updateDataList;
        string GroupName = "生活感想";
        string _txt01Name = "生活感想_內容1";
        string _txt02Name = "生活感想_內容2";
        string _txt03Name = "生活感想_內容3";
        string _dtDateName = "生活感想_填寫日期";
        Dictionary<string, string> _labelDict;

        public StudABCard06Content()
        {
            InitializeComponent();
            _txtdateV = new TextBox();
            _StudentIDList = new List<string>();
            _insertDataList = new List<UDTSingleRecordDef>();
            _updateDataList = new List<UDTSingleRecordDef>();
            _dataDict = new Dictionary<string, UDTSingleRecordDef>();
            _ChangeListener.Add(new TextBoxSource(txt01));
            _ChangeListener.Add(new TextBoxSource(txt02));
            _ChangeListener.Add(new TextBoxSource(txt03));
            _ChangeListener.Add(new TextBoxSource(_txtdateV));

            _labelDict = new Dictionary<string, string>();
            _labelDict.Add("11", "期望");
            _labelDict.Add("12", "為達理想，所需要的努力");
            _labelDict.Add("13", "期望師長給幫助");
            _labelDict.Add("21", "一年來的感想");
            _labelDict.Add("22", "今後努力的目標");
            _labelDict.Add("23", "期望師長給幫助");
            _labelDict.Add("31", "項目1");
            _labelDict.Add("32", "項目2");
            _labelDict.Add("33", "項目3");

            // 填入班級年級
            cbxGradeYear.Items.AddRange(Utility.GetClassGradeYear().ToArray());
            // 預設未選年級前無法使用
            SetControlEnable(false);
            _ChangeListener.StatusChanged += new EventHandler<ChangeEventArgs>(_ChangeListener_StatusChanged);
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);
            this.Group = "綜合表現紀錄表-生活感想";
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

        void _ChangeListener_StatusChanged(object sender, ChangeEventArgs e)
        {
            CancelButtonVisible = (e.Status == ValueStatus.Dirty);
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

        /// <summary>
        /// 設定控制項初始名稱
        /// </summary>
        private void SetDefaultControlName()
        {
            //lbl01.Text = lbl02.Text = lbl03.Text = "";
            txt01.Name = "txt01";
            txt02.Name = "txt02";
            txt03.Name = "txt03";
            dtFillDate.Name = "dtFill";
            //txt03.Size=txt02.Size= txt01.Size = new Size(460, 25);
            //txt01.Location = new Point(65, txt01.Location.Y);
            //txt02.Location = new Point(65, txt02.Location.Y);
            //txt03.Location = new Point(65, txt03.Location.Y);
            //lbl01.AutoSize = true;
            //lbl02.AutoSize = true;
            //lbl03.AutoSize = true;
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

        protected override void OnSaveButtonClick(EventArgs e)
        {
            SetData();
            if (_insertDataList.Count > 0)
                UDTTransfer.ABUDTSingleRecordInsert(_insertDataList);

            if (_updateDataList.Count > 0)
                UDTTransfer.ABUDTSingleRecordUpdate(_updateDataList);

            this.CancelButtonVisible = this.SaveButtonVisible = false;
            cbxGradeYear.Enabled = true;
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

        /// <summary>
        /// 設定控制項是否啟用
        /// </summary>
        /// <param name="tf"></param>
        private void SetControlEnable(bool tf)
        {
            txt01.Enabled = txt02.Enabled = txt03.Enabled = tf;
            dtFillDate.Enabled = tf;
        }

        protected override void OnCancelButtonClick(EventArgs e)
        {
            this.CancelButtonVisible = this.SaveButtonVisible = false;
            cbxGradeYear.Enabled = true;
            _BGRun();
        }

        private void ClearTxt()
        {            
            _txtdateV.Text = txt03.Text = txt02.Text = txt01.Text = "";
            dtFillDate.Text = "";
        }

        private void cbxGradeYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbl03.Text = lbl02.Text = lbl01.Text = "";

            if (string.IsNullOrEmpty(cbxGradeYear.Text))
            {
                ClearTxt();                
                SetControlEnable(false);
            }
            else
            {
                string k1 = cbxGradeYear.Text + "1";
                string k2 = cbxGradeYear.Text + "2";
                string k3 = cbxGradeYear.Text + "3";

                if (_labelDict.ContainsKey(k1))
                    lbl01.Text = _labelDict[k1];

                if (_labelDict.ContainsKey(k2))
                    lbl02.Text = _labelDict[k2];

                if (_labelDict.ContainsKey(k3))
                    lbl03.Text = _labelDict[k3];

                SetControlEnable(true);
                dtFillDate.Name = _dtDateName + "_" + cbxGradeYear.Text;
                txt01.Name = _txt01Name + "_" + cbxGradeYear.Text;
                txt02.Name = _txt02Name + "_" + cbxGradeYear.Text;
                txt03.Name = _txt03Name + "_" + cbxGradeYear.Text;
                dtFillDate.Tag = null;
                txt01.Tag = txt02.Tag = txt03.Tag = null;


                // 位置調整
                txt03.Size = txt02.Size = txt01.Size = new Size(527, 25);
                int x1 = lbl01.Location.X + lbl01.Size.Width+2;
                txt01.Location = new Point(x1, txt01.Location.Y);
                txt01.Size = new System.Drawing.Size(txt01.Size.Width - x1, txt01.Size.Height);

                int x2 = lbl02.Location.X + lbl02.Size.Width + 2;
                txt02.Location = new Point(x2, txt02.Location.Y);
                txt02.Size = new System.Drawing.Size(txt02.Size.Width - x2, txt02.Size.Height);

                int x3 = lbl03.Location.X + lbl03.Size.Width + 2;
                txt03.Location = new Point(x3, txt03.Location.Y);
                txt03.Size = new System.Drawing.Size(txt03.Size.Width - x3, txt03.Size.Height);

                _BGRun();
            }
        }

        private void dtFillDate_TextChanged(object sender, EventArgs e)
        {
            _txtdateV.Text = dtFillDate.Text;
        }

    }
}
