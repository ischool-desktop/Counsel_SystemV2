using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Counsel_System2.DAO;
using System.Xml;
using Counsel_System2.UI;

namespace Counsel_System2.Contents
{

    /// <summary>
    /// 學生.AB表
    /// </summary>
    [FISCA.Permission.FeatureCode(PermissionCode.綜合表現紀錄表_資料項目, "綜合表現紀錄表")]
    public partial class BasicInfo : FISCA.Presentation.DetailContent
    {
        private List<UDT_ABCardTemplateDefinitionDef> template;
        private DAO.UDT_ABCardDataDef card_data;
        private SubjectUIMaker subjUim;

        private UDT_ABCardTemplateDefinitionDef currentTemplate;
        private bool isSomeoneWaiting = false;

        public BasicInfo()
        {
            InitializeComponent();
            
            this.Group = "綜合表現紀錄表(舊)";
        }

        private void loadTemplate()
        {
            FISCA.UDT.AccessHelper ah = new FISCA.UDT.AccessHelper();
            this.template = ah.Select<UDT_ABCardTemplateDefinitionDef>();
            this.template.Sort(delegate(UDT_ABCardTemplateDefinitionDef t1, UDT_ABCardTemplateDefinitionDef t2)
            {
                return t1.ToString().CompareTo(t2.ToString());
            }
            );
            //MessageBox.Show(this.template.Count.ToString());
            this.comboBoxEx1.Items.Clear();
            this.comboBoxEx1.Items.AddRange(this.template.ToArray());

            if (this.comboBoxEx1.Items.Count > 0)
                this.comboBoxEx1.SelectedIndex = 0;

            //initUI();
        }

        private void initUI()
        {
            this.pnlQArea.Controls.Clear();

            UDT_ABCardTemplateDefinitionDef subject_template = (UDT_ABCardTemplateDefinitionDef)this.comboBoxEx1.SelectedItem;

            this.currentTemplate = subject_template;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(subject_template.Content);
            this.subjUim = new Counsel_System2.UI.SubjectUIMaker(xmlDoc.DocumentElement, this.pnlQArea);
            this.subjUim.OnContentChange += new SubjectUIMaker.ContentChange(subjUim_OnContentChange);
            this.subjUim.ResetContent();
            this.panel1.Height = this.pnlQArea.Height;
            this.Height = this.panel1.Height + 50;
            

        }

        void subjUim_OnContentChange(object sender, EventArgs e)
        {
            this.showSaveButton(true);
        }

        /// <summary>
        /// 載入某位學生的 AB 卡資料
        /// </summary>
        private void loadABCardData()
        {
            this.Loading = true;

            if (this.backgroundWorker1.IsBusy)
                this.isSomeoneWaiting = true;
            else
                this.backgroundWorker1.RunWorkerAsync();
        }

        private void BasicInfo_Load(object sender, EventArgs e)
        {
            this.pnlQArea.Size = this.panel1.Size;
            this.pnlQArea.Location = new Point(0, 0);
            this.loadTemplate();
            this.SaveButtonClick += new EventHandler(BasicInfo_SaveButtonClick);
            this.CancelButtonClick += new EventHandler(BasicInfo_CancelButtonClick);
            this.PrimaryKeyChanged += new EventHandler(BasicInfo_PrimaryKeyChanged);
        }

        void BasicInfo_PrimaryKeyChanged(object sender, EventArgs e)
        {


            this.loadABCardData();
        }

        void BasicInfo_CancelButtonClick(object sender, EventArgs e)
        {
            this.showSaveButton(false);
        }

        void BasicInfo_SaveButtonClick(object sender, EventArgs e)
        {
            this.Loading = true;
            if (this.card_data == null)
            {
                this.card_data = new UDT_ABCardDataDef();
                this.card_data.StudentID = int.Parse(this.PrimaryKey);
                this.card_data.SubjectName = this.currentTemplate.SubjectName;
                this.card_data.TemplateID = int.Parse(this.currentTemplate.UID);
                this.card_data.Content = this.subjUim.GetAnswers();
            }
            else
            {
                this.card_data.Content = this.subjUim.GetAnswers();
            }

            this.backgroundWorker2.RunWorkerAsync();
        }

        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.initUI();
            if (!string.IsNullOrEmpty(this.PrimaryKey))
                this.loadABCardData();
        }

        private void showSaveButton(bool isVisible)
        {
            this.SaveButtonVisible = isVisible;
            this.CancelButtonVisible = isVisible;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            FISCA.UDT.AccessHelper ah = new FISCA.UDT.AccessHelper();
            string query = string.Format("ref_student_id={0} and ref_template_id ='{1}'", this.PrimaryKey, this.currentTemplate.UID);
            List<DAO.UDT_ABCardDataDef> records = ah.Select<DAO.UDT_ABCardDataDef>(query);
            if (records.Count > 0)
            {
                this.card_data = records[0];
            }
            else
            {
                this.card_data = null;
            }
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            FISCA.UDT.AccessHelper ah = new FISCA.UDT.AccessHelper();

            List<FISCA.UDT.ActiveRecord> objs = new List<FISCA.UDT.ActiveRecord>();
            objs.Add(this.card_data);

            if (string.IsNullOrEmpty(this.card_data.UID))
            {
                ah.InsertValues(objs);
            }
            else
            {
                ah.UpdateValues(objs);
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                if (this.isSomeoneWaiting)
                {
                    this.isSomeoneWaiting = false;
                    this.loadABCardData();
                }
                else
                {
                    this.isSomeoneWaiting = false;
                    this.subjUim.ResetContent();

                    if (this.card_data != null)
                    {
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.LoadXml(this.card_data.Content);
                        this.subjUim.SetAnswer(xmlDoc.DocumentElement);
                    }

                    this.showSaveButton(false);
                    this.Loading = false;
                }
            }
            catch(Exception ex)
            { }

        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                this.showSaveButton(false);
                this.loadABCardData();
            }
            catch(Exception ex)
            {}
        }

        private void BasicInfo_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Forms.SetABTemplateXMLForm sax = new Forms.SetABTemplateXMLForm();
            sax.ShowDialog();
        }

    }
}
