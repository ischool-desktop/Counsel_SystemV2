using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Aspose.Words;
using Campus.Report;
using System.Windows.Forms;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// Word 樣板交換
    /// </summary>
    public class DocTemplateTransfer
    {
        private ReportConfiguration _Config;

        public DocTemplateTransfer(string GlobalUserDefineTemplateName)
        {
            _Config = new ReportConfiguration(GlobalUserDefineTemplateName);
            SetDefaultTemplate();
        }

        /// <summary>
        /// 設定用預設樣版
        /// </summary>
        public void SetDefaultTemplate()
        {
            if (_Config.Template == null)
            {
                ReportTemplate template = new ReportTemplate(Properties.Resources.輔導紀錄_樣板,TemplateType.Word);
                _Config.Template = template;
                _Config.Save();
            }
        }

        /// <summary>
        /// 上傳樣板
        /// </summary>
        public void UploadTemplate()
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Word (*.doc)|*.doc";
            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                FileInfo fileInfo = new FileInfo(openDialog.FileName);
                TemplateType type = TemplateType.Word;
                ReportTemplate template = new ReportTemplate(fileInfo, type);
                _Config.Template = template;
                _Config.Save();
            }
        
        }

        /// <summary>
        /// 取得目前的樣板 MemoryStream 型態
        /// </summary>
        public byte[] GetUsingTemplate()
        {
            return _Config.Template.ToBinary();
        }

        /// <summary>
        /// 移除樣板並使用預設
        /// </summary>
        public void RemoveAndUseDefaultTemplate()
        {
            if (FISCA.Presentation.Controls.MsgBox.Show("請問確定要移除自訂範本嗎?", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                _Config.Template = null;
                SetDefaultTemplate();
            }        
        }

        /// <summary>
        /// 下載目前樣版
        /// </summary>
        public void DownloadTemplate()
        {
            if (_Config.Template == null)
            {
                FISCA.Presentation.Controls.MsgBox.Show("目前沒有任何範本");
                return;
            }

            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Word (*.doc)|*.doc";
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _Config.Template.ToDocument().Save(saveDialog.FileName);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("儲存失敗。" + ex.Message);
                    return;
                }

                try
                {
                    System.Diagnostics.Process.Start(saveDialog.FileName);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("開啟失敗。" + ex.Message);
                    return;
                }

            }
        }
    }
}
