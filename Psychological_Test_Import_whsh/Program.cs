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
                 
            #endregion



            // 匯入測驗(新編多元性向測驗文華高中測驗結果)
            Catalog catalog1b2 = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog1b2.Add(new RibbonFeature("K12.Student.StudQuizDataImport(新編多元性向測驗文華高中測驗結果)", "匯入測驗資料(新編多元性向測驗文華高中測驗結果)"));

            RibbonBarItem rbItem1 = MotherForm.RibbonBarItems["學生", "輔導"];// MotherForm.RibbonBarItems["輔導", "測驗"];

            rbItem1["匯入"].Items["匯入心理測驗(新編多元性向測驗文華高中測驗結果)"].Enable = UserAcl.Current["K12.Student.StudQuizDataImport(新編多元性向測驗文華高中測驗結果)"].Executable;
            rbItem1["匯入"].Items["匯入心理測驗(新編多元性向測驗文華高中測驗結果)"].Click += delegate
            {

            

                Forms.ImportStudent_AptitudeTest ISAT = new Forms.ImportStudent_AptitudeTest();

                ISAT.ShowDialog();

                //DAO.UDT_Aptitude_Test_Data_Def UATDD = new DAO.UDT_Aptitude_Test_Data_Def();

                //UATDD.language_reasoning_original_score = 22;
                //UATDD.language_reasoning_scale_score = 14;
                //UATDD.language_reasoning_pr_level = 92;

                //UATDD.digit_reasoning_original_score = 16;
                //UATDD.digit_reasoning_scale_score = 11;
                //UATDD.digit_reasoning_pr_level = 68;

                //UATDD.image_reasoning_original_score = 15;
                //UATDD.image_reasoning_scale_score = 13;
                //UATDD.image_reasoning_pr_level = 82;

                //UATDD.mechanical_reasoning_original_score = 15;
                //UATDD.mechanical_reasoning_scale_score = 11;
                //UATDD.mechanical_reasoning_pr_level = 63;

                //UATDD.dimension_relation_original_score = 18;
                //UATDD.dimension_relation_scale_score = 12;
                //UATDD.dimension_relation_pr_level = 75;

                //UATDD.chineses_words_original_score = 44;
                //UATDD.chineses_words_scale_score = 12;
                //UATDD.chineses_words_pr_level = 79;

                //UATDD.english_words_original_score = 13;
                //UATDD.english_words_scale_score = 10;
                //UATDD.english_words_pr_level = 46;

                //UATDD.perception_time_original_score = 35;
                //UATDD.perception_time_scale_score = 14;
                //UATDD.perception_time_pr_level = 89;

                //UATDD.learning_aptitude_assemble_score = 117;
                //UATDD.learning_aptitude_pr_level = 87;

                //UATDD.science_aptitude_assemble_score = 113;
                //UATDD.science_aptitude_pr_level = 81;

                //UATDD.literal_aptitude_assemble_score = 117;
                //UATDD.literal_aptitude_pr_level = 87;

                //UATDD.perception_time_complete_quiz_count = 40;
                //UATDD.perception_time_correct_quiz_count = 35;



                //UATDD.StudentID = 483;

                //UATDD.Save();

            };


        }


    }
}
