using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using FISCA.Presentation;
using FISCA.Permission;
using System.Xml.Linq;
using System.ComponentModel;
using K12.Data;
using FISCA.Presentation.Controls;


namespace AfterGratuationStatics_whsh
{

    
    public class Program
    {
        [FISCA.MainMethod()]

        public static void Main()
        {

            Catalog catalog1b2 = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog1b2.Add(new RibbonFeature("K12.Student.AfterGratuationStatics_whsh(畢業後進路公務統計報表)", "畢業後進路公務統計報表"));


            RibbonBarItem rbRptItem1 = MotherForm.RibbonBarItems["學生", "輔導"];


            rbRptItem1["報表"]["畢業後進路公務統計報表 "]["0.下載畢業生進路調查樣板"].Enable = UserAcl.Current["K12.Student.AfterGratuationStatics_whsh(畢業後進路公務統計報表)"].Executable;

            rbRptItem1["報表"]["畢業後進路公務統計報表 "]["0.下載畢業生進路調查樣板"].Click += delegate
            {


                Forms.Step0_Form s0f = new Forms.Step0_Form();

                s0f.ShowDialog();

            };


            rbRptItem1["報表"]["畢業後進路公務統計報表 "]["1.匯入畢業生進路調查樣板"].Enable = UserAcl.Current["K12.Student.AfterGratuationStatics_whsh(畢業後進路公務統計報表)"].Executable;

            rbRptItem1["報表"]["畢業後進路公務統計報表 "]["1.匯入畢業生進路調查樣板"].Click += delegate
            {

                Forms.Step1_Form s1f = new Forms.Step1_Form();

                s1f.ShowDialog();

            };

            rbRptItem1["報表"]["畢業後進路公務統計報表 "]["2.匯入畢業生進路調查公務統計報表(學校分類)"].Enable = UserAcl.Current["K12.Student.AfterGratuationStatics_whsh(畢業後進路公務統計報表)"].Executable;

            rbRptItem1["報表"]["畢業後進路公務統計報表 "]["2.匯入畢業生進路調查公務統計報表(學校分類)"].Click += delegate
            {
                Forms.Step2_Form s2f = new Forms.Step2_Form();

                s2f.ShowDialog();

            };


            rbRptItem1["報表"]["畢業後進路公務統計報表 "]["3.匯入畢業生進路調查公務統計報表(學生系統分類)"].Enable = UserAcl.Current["K12.Student.AfterGratuationStatics_whsh(畢業後進路公務統計報表)"].Executable;

            rbRptItem1["報表"]["畢業後進路公務統計報表 "]["3.匯入畢業生進路調查公務統計報表(學生系統分類)"].Click += delegate
            {
                Forms.Step3_Form s3f = new Forms.Step3_Form();

                s3f.ShowDialog();

            };

        }


    }
}
