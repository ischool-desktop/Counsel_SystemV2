using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace Counsel_System2.Forms
{
    public partial class SetABCardAccessStartingDate : FISCA.Presentation.Controls.BaseForm
    {
        // 系統設定值
        DAO.UDT_SystemListDef _SystemListDef;
        BackgroundWorker _bgWorker;
        string _ConfigName = "ABCardAccessStartingDate";
        List<string> _GradeYearList;
        // 日期時間格式 2111/12/15 16:30:00
        string _DTFormat = "yyyy/MM/dd HH:mm:ss";
        string strStartDateTime = "StartDateTime";
        string strEndDateTime = "EndDateTime";

        public SetABCardAccessStartingDate()
        {
            InitializeComponent();
            this.MaximumSize = this.MinimumSize = this.Size;
            _SystemListDef = new DAO.UDT_SystemListDef();
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _GradeYearList = new List<string>();
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);

        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            lblMsg.Text = "日期與時間內容範例：2011/12/15 23:30:00";
            LoadUDTDataToDataGridView();

        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _GradeYearList = Utility.GetClassGradeYear();            
        }

        private void LoadUDTDataToDataGridView()
        {
            _SystemListDef = DAO.UDTTransfer.GetSystemListByName(_ConfigName);
            int rowIdx = 0;
            _GradeYearList.Sort();
            dgStatEndDateTime.Rows.Clear();
            if (!string.IsNullOrEmpty(_SystemListDef.Content))
            {
                XElement elmRoot = XElement.Parse(_SystemListDef.Content);

                foreach (string str in _GradeYearList)
                {
                    rowIdx = dgStatEndDateTime.Rows.Add();
                    List<XElement> elm = (from data in elmRoot.Elements("Item") where data.Attribute("GradeYear").Value==str select data).ToList();

                    // 年級
                    dgStatEndDateTime.Rows[rowIdx].Cells[colGradeYear.Index].Value = str;

                    if (elm.Count == 1)
                    {
                        DateTime dtStart, dtEnd;
                        if (elm[0].Attribute(strStartDateTime) != null)
                        {
                            if (DateTime.TryParse(elm[0].Attribute(strStartDateTime).Value, out dtStart))
                                dgStatEndDateTime.Rows[rowIdx].Cells[colStartDateTime.Index].Value = dtStart.ToString(_DTFormat);
                                                        
                        }

                        if(elm[0].Attribute(strEndDateTime) !=null)
                        {
                            if (DateTime.TryParse(elm[0].Attribute(strEndDateTime).Value, out dtEnd))
                                dgStatEndDateTime.Rows[rowIdx].Cells[colEndDateTime.Index].Value = dtEnd.ToString(_DTFormat);
                        }
                    }
                }
               

                // 年級唯讀
                dgStatEndDateTime.Columns[colGradeYear.Index].ReadOnly = true;                
            }
            else
            {
                foreach (string str in _GradeYearList)
                {
                    rowIdx = dgStatEndDateTime.Rows.Add();
                    // 年級
                    dgStatEndDateTime.Rows[rowIdx].Cells[colGradeYear.Index].Value = str;                    
                }
            }
        }

        private void SetUDTData()
        {
            bool isSave = true;

            // 資料是否有標錯誤檢查
            foreach (DataGridViewRow drv in dgStatEndDateTime.Rows)
            {
                if (drv.Cells[colStartDateTime.Index].ErrorText != "" || drv.Cells[colEndDateTime.Index].ErrorText!="")
                    isSave = false;
            }

            if (isSave)
            {
                // name
                if (_SystemListDef.Name != _ConfigName)
                    _SystemListDef.Name = _ConfigName;

                // content
                XElement elmRoot = new XElement("Content");

                // 轉換成 XML
                foreach (DataGridViewRow dgv in dgStatEndDateTime.Rows)
                {
                    if (dgv.IsNewRow)
                        continue;

                    XElement elm = new XElement("Item");
                    DateTime dts, dte;

                    // 年級
                    elm.SetAttributeValue("GradeYear", dgv.Cells[colGradeYear.Index].Value.ToString());

                    // 開始
                    if (dgv.Cells[colStartDateTime.Index].Value != null)
                    {
                        if (DateTime.TryParse(dgv.Cells[colStartDateTime.Index].Value.ToString(), out dts))
                            elm.SetAttributeValue(strStartDateTime, dts.ToString(_DTFormat));
                    }
                    else
                        elm.SetAttributeValue(strStartDateTime,"");

                    // 結束
                    if (dgv.Cells[colEndDateTime.Index].Value != null)
                    {
                        if (DateTime.TryParse(dgv.Cells[colEndDateTime.Index].Value.ToString(), out dte))
                            elm.SetAttributeValue(strEndDateTime, dte.ToString(_DTFormat));
                    }
                    else
                        elm.SetAttributeValue(strEndDateTime,"");



                    elmRoot.Add(elm);
                }

                _SystemListDef.Content = elmRoot.ToString();

                List<DAO.UDT_SystemListDef> dataList = new List<DAO.UDT_SystemListDef>();
                dataList.Add(_SystemListDef);
                if (string.IsNullOrEmpty(_SystemListDef.UID))
                    DAO.UDTTransfer.InsertSystemList(dataList);
                else
                    DAO.UDTTransfer.UpdateSystemList(dataList);

                FISCA.Presentation.Controls.MsgBox.Show("儲存完成.");
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("日期與時間設定有錯誤，無法儲存!");
                return;            
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            btnSave.Enabled = false;
            SetUDTData();
            _bgWorker.RunWorkerAsync();
            btnSave.Enabled = true;
            
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void SetABCardAccessStartingDate_Load(object sender, EventArgs e)
        {
            lblMsg.Text = "資料讀取中...";
            _bgWorker.RunWorkerAsync();
        }

        private void dgStatEndDateTime_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            DateTime dts, dte;
            foreach (DataGridViewRow drv in dgStatEndDateTime.Rows)
            {
                if (drv.Cells[colStartDateTime.Index].Value != null)
                    if (DateTime.TryParse(drv.Cells[colStartDateTime.Index].Value.ToString(), out dts))
                        drv.Cells[colStartDateTime.Index].ErrorText = "";
                    else
                        drv.Cells[colStartDateTime.Index].ErrorText = "輸入日期與時間有錯誤,無法儲存!";

                if (drv.Cells[colEndDateTime.Index].Value != null)
                    if (DateTime.TryParse(drv.Cells[colEndDateTime.Index].Value.ToString(), out dte))
                        drv.Cells[colEndDateTime.Index].ErrorText = "";
                    else
                        drv.Cells[colEndDateTime.Index].ErrorText = "輸入日期與時間有錯誤,無法儲存!";
            }
        }
    }
}
