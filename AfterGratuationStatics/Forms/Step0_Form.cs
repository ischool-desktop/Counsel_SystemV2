using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using FISCA.Presentation.Controls;
using Aspose.Cells;
using K12.Data;


namespace AfterGratuationStatics_whsh.Forms
{
    public partial class Step0_Form : BaseForm
    {
        
        public Step0_Form()
        {
            InitializeComponent();
        }


        // 取消
        private void buttonX3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //下載
        private void buttonX2_Click(object sender, EventArgs e)
        {


            //Workbook wb = new Workbook(new MemoryStream(Properties.Resources.新編多元性向測驗文華高中測驗範例樣板_CSV));

            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.畢業生進路調查公務統計報表_教務處空樣板_));


            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "畢業生進路調查公務統計報表(教務處空樣板).xlsx";
            sd.Filter = "Excel檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    wb.Save(sd.FileName);
                    System.Diagnostics.Process.Start(sd.FileName);
                }
                catch
                {
                    MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }


            this.Close();
        }

       
    }
}
