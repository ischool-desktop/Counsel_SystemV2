using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Permission;
using FISCA.Presentation;

namespace CounselTools
{
    public class Program
    {
        [FISCA.MainMethod()]
        public static void Main()
        {
            // 列印綜合紀錄表未輸入完整名單
            Catalog catalog1 = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog1.Add(new RibbonFeature("K12.CounselTools.ExportNotInputABCardComplete", "綜合紀錄表未輸入完整名單"));

            RibbonBarItem rbRptItemABNew = MotherForm.RibbonBarItems["學生", "輔導"];
            rbRptItemABNew["報表"]["綜合紀錄表未輸入完整名單"].Enable = UserAcl.Current["K12.CounselTools.ExportNotInputABCardComplete"].Executable;
            rbRptItemABNew["報表"]["綜合紀錄表未輸入完整名單"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    ExportNotInputABCardComplete enabCC = new ExportNotInputABCardComplete(K12.Presentation.NLDPanels.Student.SelectedSource);
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生");
                    return;                
                }            
            };
        }
    }
}
