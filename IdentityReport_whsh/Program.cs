using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Presentation;
using FISCA.Permission;
using System.Xml.Linq;
using System.ComponentModel;
using K12.Data;
using K12.Presentation;
using FISCA.Presentation.Controls;


namespace IdentityReport_whsh
{

    public class Program
    {
        [FISCA.MainMethod()]

        public static void Main()
        {

            //2017/7/4 穎驊新增
            // 列印輔導案量統計
            Catalog catalog27 = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog27.Add(new RibbonFeature("K12.Counsel_System2.statics", "身分報表"));

            RibbonBarItem rbRptItem = MotherForm.RibbonBarItems["學生", "輔導"];
            rbRptItem["報表"]["身分報表"]["特殊監護名冊"].Enable = UserAcl.Current["K12.Counsel_System2.statics"].Executable;
            rbRptItem["報表"]["身分報表"]["特殊監護名冊"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    SpecialGuardian sg = new SpecialGuardian(K12.Presentation.NLDPanels.Student.SelectedSource);
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生");
                    return;
                }
            };

            //2017/7/25 穎驊新增
            //經由與文華輔導室來回確認規格後，繼續把剩下的統計表完成。        
            rbRptItem["報表"]["身分報表"]["新住民統計表"].Enable = UserAcl.Current["K12.Counsel_System2.statics"].Executable;
            rbRptItem["報表"]["身分報表"]["新住民統計表"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    NewResident nr = new NewResident(K12.Presentation.NLDPanels.Student.SelectedSource);
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生");
                    return;
                }
            };

            rbRptItem["報表"]["身分報表"]["單親家庭統計表"].Enable = UserAcl.Current["K12.Counsel_System2.statics"].Executable;
            rbRptItem["報表"]["身分報表"]["單親家庭統計表"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    SingleParent sp = new SingleParent(K12.Presentation.NLDPanels.Student.SelectedSource);
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