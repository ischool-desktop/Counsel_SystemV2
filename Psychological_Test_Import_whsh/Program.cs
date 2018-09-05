using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.DocumentValidator;
using FISCA.Presentation;
using FISCA.Permission;
using K12.Presentation;
using System.Xml.Linq;
using System.ComponentModel;
using FISCA.Data;
using FISCA.UDT;
using K12.Data;
using FISCA.Presentation.Controls;
using System.Windows.Forms;

namespace Psychological_Test_Import_whsh
{


    public class Program
    {
        [FISCA.MainMethod()]

        public static void Main()
        {


            #region 處理UDT Table沒有的問題

            // 初始化 強迫新增，可以避開 Web第一次使用找不到Table的錯誤
            AccessHelper _accessHelper = new AccessHelper();
            _accessHelper.Select<DAO.UDT_Aptitude_Test_Data_Def>("UID = '00000'");

            //2018/9/5 穎驊 註解，舊版 興趣測驗作廢，更新為新格式
            //_accessHelper.Select<DAO.UDT_Interest_Test_Data_Def>("UID = '00000'");

            //2018/9 新版的 興趣測驗
            _accessHelper.Select<DAO.UDT_Interest_Test_Data_Def_2018_09>("UID = '00000'");

            //2018/9/ 新文華項目學系探索量表測驗匯入
            _accessHelper.Select<DAO.UDT_MajorIn_Test_Data_Def>("UID = '00000'");

            #endregion



            // 匯入心理測驗(新編多元性向測驗文華高中測驗結果)
            Catalog catalog1b2 = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog1b2.Add(new RibbonFeature("K12.Student.StudQuizDataImport(新編多元性向測驗文華高中測驗結果)", "匯入測驗資料(新編多元性向測驗文華高中測驗結果)"));

            RibbonBarItem rbItem1 = MotherForm.RibbonBarItems["學生", "輔導"];// MotherForm.RibbonBarItems["輔導", "測驗"];

            rbItem1["匯入"].Items["匯入心理測驗(新編多元性向測驗文華高中測驗結果)"].Enable = UserAcl.Current["K12.Student.StudQuizDataImport(新編多元性向測驗文華高中測驗結果)"].Executable;
            rbItem1["匯入"].Items["匯入心理測驗(新編多元性向測驗文華高中測驗結果)"].Click += delegate
            {

                Forms.ImportStudent_AptitudeTest ISAT = new Forms.ImportStudent_AptitudeTest();

                ISAT.ShowDialog();

            };

            // 匯入興趣測驗

            //Catalog catalog1b3 = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog1b2.Add(new RibbonFeature("K12.Student.StudQuizDataImport(匯入興趣測驗資料)", "匯入興趣測驗資料"));


            RibbonBarItem rbItem2 = MotherForm.RibbonBarItems["學生", "輔導"];// MotherForm.RibbonBarItems["輔導", "測驗"];

            // 日後記得來調權限問題
            //rbItem2["匯入"].Items["匯入興趣測驗資料"].Enable = UserAcl.Current["K12.Student.StudQuizDataImport(匯入興趣測驗資料)"].Executable;

            rbItem2["匯入"].Items["匯入興趣測驗資料"].Click += delegate
            {

                Forms.ImportStudent_InterestTest ISIT = new Forms.ImportStudent_InterestTest();

                ISIT.ShowDialog();

            };

            rbItem2["匯入"].Items["匯入學系探索量表測驗"].Click += delegate
            {

                Forms.ImportStudent_MajorInTest ISMT = new Forms.ImportStudent_MajorInTest();

                ISMT.ShowDialog();

            };


        }


    }
}
