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
using K12.Data;
using FISCA.Presentation.Controls;
using System.Windows.Forms;

namespace Counsel_System2
{
    /// <summary>
    /// 輔導
    /// </summary>
    public class Program
    {
        static BackgroundWorker _bkWork;
        static bool _hasABTemplate = false;
        

        [FISCA.MainMethod()]
        public static void Main()
        {
        

            _bkWork = new BackgroundWorker();
            _bkWork.DoWork += new DoWorkEventHandler(_bkWork_DoWork);
            _bkWork.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bkWork_RunWorkerCompleted);
            _bkWork.RunWorkerAsync();
        
            // 20120510 將輔導頁簽移除
            //// 輔導底下空白
            //MotherForm.AddPanel(CounselAdmin.Instance);


        

            // 輔導->測驗->管理測驗試別。
            //輔導->測驗->匯入測驗資料。
            //學生->資料項目->測驗資料。

            RibbonBarItem rbItem1 = MotherForm.RibbonBarItems["學生", "輔導"];// MotherForm.RibbonBarItems["輔導", "測驗"];
            rbItem1["設定"].Image = Properties.Resources.設定;
            rbItem1["設定"].Size = RibbonBarButton.MenuButtonSize.Large;
            rbItem1["設定"].Items["設定心理測驗試別"].Enable = UserAcl.Current["K12.Student.StudQuizDataManager"].Executable;
            rbItem1["設定"].Items["設定心理測驗試別"].Click += delegate
            {
                Forms.QuizForm qf = new Forms.QuizForm();
                qf.ShowDialog();

            };

            rbItem1["匯入"].Items["匯入心理測驗"].Enable = UserAcl.Current["K12.Student.StudQuizDataImport"].Executable;
            rbItem1["匯入"].Items["匯入心理測驗"].Click += delegate
            {
                Global._AllStudentNumberStatusIDTemp = Utility.GetAllStudenNumberStatusDict();
                Global._AllStudentClassSeatNoDictTemp = Utility.GetAllStudentClassSeatNoStatusDict();
                Global._StudentStatusDBDict = Utility.GetStudentStatusDBValDict();
                Forms.ImportStudQuizDataSelectQuizNameForm isqdsqf = new Forms.ImportStudQuizDataSelectQuizNameForm();
                isqdsqf.ShowDialog();

            };

            // 測驗
            if (Utility.CheckAddContent(PermissionCode.輔導相關測驗_資料項目))
                K12.Presentation.NLDPanels.Student.AddDetailBulider(new FISCA.Presentation.DetailBulider<Contents.StudQuizDataContent>());


            //2016/11/25 穎驊筆記，因應輔導2.0 table counsel.interview_record 調整，牽連到的功能先暫時關閉，日後調整再開啟。

            //// 晤談
            //if (Utility.CheckAddContent(PermissionCode.輔導晤談紀錄_資料項目))
            //    K12.Presentation.NLDPanels.Student.AddDetailBulider(new FISCA.Presentation.DetailBulider<Contents.StudInterviewDataContent>());

            // 優先關懷
            if (Utility.CheckAddContent(PermissionCode.輔導優先關懷紀錄_資料項目))
                K12.Presentation.NLDPanels.Student.AddDetailBulider(new FISCA.Presentation.DetailBulider<Contents.StudCareRecordContent>());

            // 個案會議
            if (Utility.CheckAddContent(PermissionCode.輔導個案會議_資料項目))
                K12.Presentation.NLDPanels.Student.AddDetailBulider(new FISCA.Presentation.DetailBulider<Contents.StudCaseMeetingRecordContent>());

            // 輔導老師
            if (Utility.CheckAddContent(PermissionCode.認輔老師及輔導老師_資料項目))
                K12.Presentation.NLDPanels.Student.AddDetailBulider(new FISCA.Presentation.DetailBulider<Contents.CounselStudentListContent>());

            if (Utility.CheckAddContent(PermissionCode.認輔學生_資料項目))
                K12.Presentation.NLDPanels.Teacher.AddDetailBulider(new FISCA.Presentation.DetailBulider<Contents.CounselStudentListAContent>());

            if (Utility.CheckAddContent(PermissionCode.輔導學生_資料項目))
                K12.Presentation.NLDPanels.Teacher.AddDetailBulider(new FISCA.Presentation.DetailBulider<Contents.CounselStudentListBContent>());

            // 輔導自訂欄位
            if (Utility.CheckAddContent(PermissionCode.輔導自訂欄位_資料項目))
                K12.Presentation.NLDPanels.Student.AddDetailBulider(new FISCA.Presentation.DetailBulider<Contents.CounselUserDefineDataItem>());


            // 新版綜合表現紀錄表使用 -- begin
            // 管理綜合表現題目
            RibbonBarItem rbItemA = MotherForm.RibbonBarItems["學生", "輔導"];
            rbItemA["設定"].Items["設定綜合紀錄表題目"].Enable = UserAcl.Current[PermissionCode.綜合表現紀錄表_題目管理].Executable;
            rbItemA["設定"].Items["設定綜合紀錄表題目"].Click += delegate
            {
                Forms.ABCardQuestionsForm abcqf = new Forms.ABCardQuestionsForm();
                abcqf.ShowDialog();

            };

            // 管理綜合表現題目
            Catalog catalog1abCard = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog1abCard.Add(new RibbonFeature(PermissionCode.綜合表現紀錄表_題目管理, "綜合表現紀錄表題目管理"));


            // 綜合表現紀錄
            // AB Card
            if (Utility.CheckAddContent(PermissionCode.綜合表現紀錄表_資料項目))
                K12.Presentation.NLDPanels.Student.AddDetailBulider<Counsel_System2.Contents.StudABCard01Content>();

            if (Utility.CheckAddContent(PermissionCode.綜合表現紀錄表_資料項目))
                K12.Presentation.NLDPanels.Student.AddDetailBulider<Counsel_System2.Contents.StudABCard02Content>();

            if (Utility.CheckAddContent(PermissionCode.綜合表現紀錄表_資料項目))
                K12.Presentation.NLDPanels.Student.AddDetailBulider<Counsel_System2.Contents.StudABCard03Content>();

            if (Utility.CheckAddContent(PermissionCode.綜合表現紀錄表_資料項目))
                K12.Presentation.NLDPanels.Student.AddDetailBulider<Counsel_System2.Contents.StudABCard04Content>();

            if (Utility.CheckAddContent(PermissionCode.綜合表現紀錄表_資料項目))
                K12.Presentation.NLDPanels.Student.AddDetailBulider<Counsel_System2.Contents.StudABCard05Content>();

            if (Utility.CheckAddContent(PermissionCode.綜合表現紀錄表_資料項目))
                K12.Presentation.NLDPanels.Student.AddDetailBulider<Counsel_System2.Contents.StudABCard06Content>();

            if (Utility.CheckAddContent(PermissionCode.綜合表現紀錄表_資料項目))
                K12.Presentation.NLDPanels.Student.AddDetailBulider<Counsel_System2.Contents.StudABCard07Content>();

            if (Utility.CheckAddContent(PermissionCode.綜合表現紀錄表_資料項目))
                K12.Presentation.NLDPanels.Student.AddDetailBulider<Counsel_System2.Contents.StudABCard08Content>();

            if (Utility.CheckAddContent(PermissionCode.綜合表現紀錄表_資料項目))
                K12.Presentation.NLDPanels.Student.AddDetailBulider<Counsel_System2.Contents.StudABCard09Content>();

            // 新版綜合表現紀錄表使用 -- end
            

            // 加入權限
            // 測驗
            Catalog catalog = RoleAclSource.Instance["學生"]["資料項目"];
            catalog.Add(new DetailItemFeature(PermissionCode.輔導相關測驗_資料項目, "心理測驗"));

            // 管理測驗
            Catalog catalog1a = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog1a.Add(new RibbonFeature("K12.Student.StudQuizDataManager", "管理測驗試別"));

            // 匯入測驗
            Catalog catalog1b = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog1b.Add(new RibbonFeature("K12.Student.StudQuizDataImport", "匯入測驗資料"));


            // 晤談
            Catalog catalog1 = RoleAclSource.Instance["學生"]["資料項目"];
            catalog1.Add(new DetailItemFeature(PermissionCode.輔導晤談紀錄_資料項目, "晤談紀錄"));

            // 優先關懷
            Catalog catalog2 = RoleAclSource.Instance["學生"]["資料項目"];
            catalog2.Add(new DetailItemFeature(PermissionCode.輔導優先關懷紀錄_資料項目, "優先關懷"));

            // 個案會議
            Catalog catalog3 = RoleAclSource.Instance["學生"]["資料項目"];
            catalog3.Add(new DetailItemFeature(PermissionCode.輔導個案會議_資料項目, "個案會議"));

            // 認輔老師及輔導老師
            Catalog catalog4 = RoleAclSource.Instance["學生"]["資料項目"];
            catalog4.Add(new DetailItemFeature(PermissionCode.認輔老師及輔導老師_資料項目, "認輔/輔導老師"));

            // 教師認輔學生
            Catalog catalog5 = RoleAclSource.Instance["教師"]["資料項目"];
            catalog5.Add(new DetailItemFeature(PermissionCode.認輔學生_資料項目, "認輔學生"));

            // 教師輔導學生
            Catalog catalog6 = RoleAclSource.Instance["教師"]["資料項目"];
            catalog6.Add(new DetailItemFeature(PermissionCode.輔導學生_資料項目, "輔導學生"));

            // 學生AB表
            Catalog catalog7 = RoleAclSource.Instance["學生"]["資料項目"];
            catalog7.Add(new DetailItemFeature(PermissionCode.綜合表現紀錄表_資料項目, "綜合表現紀錄表"));

            // 匯出晤談紀錄
            Catalog catalog8 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog8.Add(new RibbonFeature("K12.Student.CounselStudentExportInterViewData", "匯出晤談紀錄"));

            // 匯出優先關懷
            Catalog catalog9 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog9.Add(new RibbonFeature("K12.Student.CounselStudentExportCareRecord", "匯出優先關懷"));

            // 匯出個案會議
            Catalog catalog10 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog10.Add(new RibbonFeature("K12.Student.CounselStudentExportCaseMeetingRecord", "匯出個案會議"));

            // 匯入晤談紀錄
            Catalog catalog11 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog11.Add(new RibbonFeature("K12.Student.CounselStudentImportInterViewData", "匯入晤談紀錄"));

            // 匯入優先關懷
            Catalog catalog12 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog12.Add(new RibbonFeature("K12.Student.CounselStudentImportCareRecord", "匯入優先關懷"));

            // 匯入個案會議
            Catalog catalog13 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog13.Add(new RibbonFeature("K12.Student.CounselStudentImportCaseMeetingRecord", "匯入個案會議"));

            // 輔導自訂欄位
            Catalog catalog14 = RoleAclSource.Instance["學生"]["資料項目"];
            catalog14.Add(new DetailItemFeature(PermissionCode.輔導自訂欄位_資料項目, "輔導自訂欄位"));

            // 匯出自訂欄位
            Catalog catalog15 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog15.Add(new RibbonFeature("K12.Student.CounselStudentExportUserDefineData", "匯出輔導自訂欄位"));

            // 匯入自訂欄位
            Catalog catalog16 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog16.Add(new RibbonFeature("K12.Student.CounselStudentImportUserDefineData", "匯入輔導自訂欄位"));

            // 設定自訂欄位名稱
            Catalog catalog17 = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog17.Add(new RibbonFeature("K12.CounselStudentUserDefineDataContentSet", "設定自訂欄位名稱"));

            // 指定認輔老師
            Catalog catalog18 = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog18.Add(new RibbonFeature("K12.Teacher.CounselStudentListAContentSet", "指定認輔老師"));

            // 指定輔導老師
            Catalog catalog19 = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog19.Add(new RibbonFeature("K12.Teacher.CounselStudentListBContentSet", "指定輔導老師"));

            // 列印綜合資料紀錄表
            Catalog catalog20 = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog20.Add(new RibbonFeature("K12.Student.ABCardPrintForm", "綜合資料紀錄表"));

            //刪除輔導測驗資料
            Catalog catalog21 = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog21.Add(new RibbonFeature("K12.Student.DelStudQuizData", "刪除學生測驗資料"));


            // 列印晤談紀錄表
            Catalog catalog22 = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog22.Add(new RibbonFeature("K12.Student.StudInterviewDataReport", "學生.晤談紀錄表"));

            // 列印晤談紀錄表
            Catalog catalog23 = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog23.Add(new RibbonFeature("K12.Teacher.StudInterviewDataReport", "教師.晤談紀錄表"));

            // 列印晤談紀錄簽認表
            Catalog catalog24 = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalog23.Add(new RibbonFeature("K12.Teacher.StudInterviewData1Report", "教師.晤談紀錄簽認表"));

            // 匯出學生晤談紀錄篩選
            Catalog catalog25 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog25.Add(new RibbonFeature("K12.Student.CounselStudentExport_interviewfilter", "學生晤談紀錄篩選"));

            // 匯出家庭聯繫紀錄篩選
            Catalog catalog26 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog26.Add(new RibbonFeature("K12.Student.CounselStudentExport_home_visitfilter", "家庭聯繫紀錄篩選"));

            //RibbonBarItem rbRptItem = MotherForm.RibbonBarItems["學生", "輔導"];
            //rbRptItem["報表"].Image = Properties.Resources.Report;
            //rbRptItem["報表"].Size = RibbonBarButton.MenuButtonSize.Large;
            //rbRptItem["報表"]["綜合資料紀錄表"].Enable = UserAcl.Current["K12.Student.ABCardPrintForm"].Executable;
            //rbRptItem["報表"]["綜合資料紀錄表"].Click += delegate
            //{
            //    if (NLDPanels.Student.SelectedSource.Count > 0)
            //    {
            //        Forms.ABCardPrintForm abcpf = new Forms.ABCardPrintForm();
            //        abcpf.ShowDialog();
            //    }
            //    else
            //        FISCA.Presentation.Controls.MsgBox.Show("請選擇學生.");
            //};

            
            RibbonBarItem rbRptItemABNew = MotherForm.RibbonBarItems["學生", "輔導"];
            rbRptItemABNew["報表"].Image = Properties.Resources.Report;
            rbRptItemABNew["報表"].Size = RibbonBarButton.MenuButtonSize.Large;
            rbRptItemABNew["報表"]["綜合資料紀錄表"].Enable = UserAcl.Current["K12.Student.ABCardPrintForm"].Executable;

            rbRptItemABNew["報表"]["綜合資料紀錄表"].Click += delegate
            {
                if (NLDPanels.Student.SelectedSource.Count > 0)
                {
                    List<string> sidList = Utility.GetStudentIDListByStudentID(K12.Presentation.NLDPanels.Student.SelectedSource);
                    Forms.StudABCardReportForm sabcrf = new Forms.StudABCardReportForm(Forms.StudABCardReportForm.SelectType.學生, sidList);
                    sabcrf.ShowDialog();
                }
                else
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生.");
            };


            
            RibbonBarItem rbRptItem1 = MotherForm.RibbonBarItems["學生", "輔導"];
            rbRptItem1["報表"].Image = Properties.Resources.Report;
            rbRptItem1["報表"].Size = RibbonBarButton.MenuButtonSize.Large;
            //rbRptItem1["報表"]["晤談紀錄表"].Enable = UserAcl.Current["K12.Student.StudInterviewDataReport"].Executable;

            //2016/11/25 穎驊筆記，因應輔導2.0 table counsel.interview_record 調整，牽連到的功能先暫時關閉，日後調整再開啟。

            rbRptItem1["報表"]["晤談紀錄表"].Enable = false;

            rbRptItem1["報表"]["晤談紀錄表"].Click += delegate
            {
                if (NLDPanels.Student.SelectedSource.Count > 0)
                {
                    Forms.StudInterviewDataReportForm sidrf = new Forms.StudInterviewDataReportForm(Forms.StudInterviewDataReportForm.SelectType.學生);
                    sidrf.ShowDialog();
                }
                else
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生.");
            };


            RibbonBarItem rbRptItem2 = MotherForm.RibbonBarItems["教師", "輔導"];
            rbRptItem2["報表"].Image = Properties.Resources.Report;
            rbRptItem2["報表"].Size = RibbonBarButton.MenuButtonSize.Large;
            rbRptItem2["報表"]["晤談紀錄表"].Enable = UserAcl.Current["K12.Teacher.StudInterviewDataReport"].Executable;
            rbRptItem2["報表"]["晤談紀錄表"].Click += delegate
            {
                if (NLDPanels.Teacher.SelectedSource.Count > 0)
                {
                    Forms.StudInterviewDataReportForm sidrf = new Forms.StudInterviewDataReportForm(Forms.StudInterviewDataReportForm.SelectType.教師);
                    sidrf.ShowDialog();
                }
                else
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇教師.");
            };

            RibbonBarItem rbRptItem3 = MotherForm.RibbonBarItems["教師", "輔導"];
            rbRptItem3["報表"].Image = Properties.Resources.Report;
            rbRptItem3["報表"].Size = RibbonBarButton.MenuButtonSize.Large;
            rbRptItem3["報表"]["晤談紀錄簽認表"].Enable = UserAcl.Current["K12.Teacher.StudInterviewData1Report"].Executable;
            rbRptItem3["報表"]["晤談紀錄簽認表"].Click += delegate
            {
                if (NLDPanels.Teacher.SelectedSource.Count > 0)
                {
                    Forms.StudInterviewDataReport1Form sidrf = new Forms.StudInterviewDataReport1Form();
                    sidrf.ShowDialog();
                }
                else
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇教師.");
            };

            // 刪除學生測驗資料
            NLDPanels.Student.ListPaneContexMenu["輔導"]["刪除學生測驗資料"].Enable = UserAcl.Current["K12.Student.DelStudQuizData"].Executable;
            NLDPanels.Student.ListPaneContexMenu["輔導"]["刪除學生測驗資料"].Click += delegate
            {
                if (NLDPanels.Student.SelectedSource.Count > 0)
                {
                    Forms.DelStudentQuizDataForm dsqdf = new Forms.DelStudentQuizDataForm(NLDPanels.Student.SelectedSource);
                    dsqdf.ShowDialog();
                }
                else
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生後再指定.");
            };


            RibbonBarItem rbItem = MotherForm.RibbonBarItems["學生", "輔導"];
            rbItem["指定"].Image = Properties.Resources.parent_coordinator_next_64;
            rbItem["指定"].Size = RibbonBarButton.MenuButtonSize.Large;
            rbItem["指定"]["指定認輔老師"].Enable = UserAcl.Current["K12.Teacher.CounselStudentListAContentSet"].Executable;
            rbItem["指定"]["指定認輔老師"].Click += delegate
            {
                if (NLDPanels.Student.SelectedSource.Count > 0)
                {
                    Forms.SetCounselTeacherForm stf = new Forms.SetCounselTeacherForm(DAO.CounselTeacherRecord.CounselTeacherType.認輔老師, NLDPanels.Student.SelectedSource);
                    stf.ShowDialog();
                }
                else
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生後再指定.");
            };

            rbItem["指定"]["指定輔導老師"].Enable = UserAcl.Current["K12.Teacher.CounselStudentListBContentSet"].Executable;
            rbItem["指定"]["指定輔導老師"].Click += delegate
            {
                if (NLDPanels.Student.SelectedSource.Count > 0)
                {
                    Forms.SetCounselTeacherForm stf = new Forms.SetCounselTeacherForm(DAO.CounselTeacherRecord.CounselTeacherType.輔導老師, NLDPanels.Student.SelectedSource);
                    stf.ShowDialog();
                }
                else
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生後再指定.");
            };

            RibbonBarItem rbItem2 = MotherForm.RibbonBarItems["學生", "輔導"];// MotherForm.RibbonBarItems["輔導", "自訂欄位"];
            rbItem2["設定"]["設定自訂欄位名稱"].Enable = UserAcl.Current["K12.CounselStudentUserDefineDataContentSet"].Executable;
            rbItem2["設定"]["設定自訂欄位名稱"].Click += delegate
            {
                Forms.SetUserDefineDataForm suddf = new Forms.SetUserDefineDataForm();
                suddf.ShowDialog();
            };

            NLDPanels.Student.ListPaneContexMenu["輔導"]["指定認輔老師"].Enable = UserAcl.Current["K12.Teacher.CounselStudentListAContentSet"].Executable;
            NLDPanels.Student.ListPaneContexMenu["輔導"]["指定認輔老師"].Click += delegate
            {
                if (NLDPanels.Student.SelectedSource.Count > 0)
                {
                    Forms.SetCounselTeacherForm stf = new Forms.SetCounselTeacherForm(DAO.CounselTeacherRecord.CounselTeacherType.認輔老師, NLDPanels.Student.SelectedSource);
                    stf.ShowDialog();
                }
                else
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生後再指定.");
            };

            NLDPanels.Student.ListPaneContexMenu["輔導"]["指定輔導老師"].Enable = UserAcl.Current["K12.Teacher.CounselStudentListBContentSet"].Executable;
            NLDPanels.Student.ListPaneContexMenu["輔導"]["指定輔導老師"].Click += delegate
            {
                if (NLDPanels.Student.SelectedSource.Count > 0)
                {
                    Forms.SetCounselTeacherForm stf = new Forms.SetCounselTeacherForm(DAO.CounselTeacherRecord.CounselTeacherType.輔導老師, NLDPanels.Student.SelectedSource);
                    stf.ShowDialog();
                }
                else
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生後再指定.");
            };



            // 匯出優先關懷
            NLDPanels.Student.RibbonBarItems["輔導"]["匯出"]["匯出優先關懷"].Enable = UserAcl.Current["K12.Student.CounselStudentExportCareRecord"].Executable;
            NLDPanels.Student.RibbonBarItems["輔導"]["匯出"]["匯出優先關懷"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new Counsel_System2.ImportExport.ExportCareRecord();
                ImportExport.ExportStudentV2 wizard = new ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };
            NLDPanels.Student.RibbonBarItems["輔導"]["匯出"].Image = Properties.Resources.Export_Image;
            NLDPanels.Student.RibbonBarItems["輔導"]["匯出"].Size = RibbonBarButton.MenuButtonSize.Large;

            // 匯出個案會議
            NLDPanels.Student.RibbonBarItems["輔導"]["匯出"]["匯出個案會議"].Enable = UserAcl.Current["K12.Student.CounselStudentExportCaseMeetingRecord"].Executable;
            NLDPanels.Student.RibbonBarItems["輔導"]["匯出"]["匯出個案會議"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new Counsel_System2.ImportExport.ExportCaseMeetingRecord();
                ImportExport.ExportStudentV2 wizard = new ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            // 匯出晤談紀錄
            NLDPanels.Student.RibbonBarItems["輔導"]["匯出"]["匯出晤談紀錄"].Enable = UserAcl.Current["K12.Student.CounselStudentExportInterViewData"].Executable;
            NLDPanels.Student.RibbonBarItems["輔導"]["匯出"]["匯出晤談紀錄"].Click += delegate
            {
                //2016/11/25 穎驊筆記，因應輔導2.0 table counsel.interview_record 調整，牽連到的功能先暫時關閉，日後調整再開啟。

                MsgBox.Show("晤談紀錄系統調整中，功能暫不開放", "警告",MessageBoxButtons.OKCancel,MessageBoxIcon.Warning);

                //SmartSchool.API.PlugIn.Export.Exporter exporter = new Counsel_System2.ImportExport.ExportInterViewData();
                //ImportExport.ExportStudentV2 wizard = new ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
                //exporter.InitializeExport(wizard);
                //wizard.ShowDialog();
            };


            // 匯出輔導自訂欄位
            NLDPanels.Student.RibbonBarItems["輔導"]["匯出"]["匯出輔導自訂欄位"].Enable = UserAcl.Current["K12.Student.CounselStudentExportUserDefineData"].Executable;
            NLDPanels.Student.RibbonBarItems["輔導"]["匯出"]["匯出輔導自訂欄位"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new Counsel_System2.ImportExport.ExportStudentUserDefData();
                ImportExport.ExportStudentV2 wizard = new ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            // 學生晤談紀錄篩選
            NLDPanels.Student.RibbonBarItems["輔導"]["匯出"]["學生晤談紀錄篩選"].Enable = UserAcl.Current["K12.Student.CounselStudentExport_interviewfilter"].Executable;
            NLDPanels.Student.RibbonBarItems["輔導"]["匯出"]["學生晤談紀錄篩選"].Click += delegate
            {
                if (NLDPanels.Student.SelectedSource.Count > 0)
                {
                    List<string> student_id_List = Utility.GetStudentIDListByStudentID(K12.Presentation.NLDPanels.Student.SelectedSource);
                    Forms.StudInterviewFilterForm Form = new Forms.StudInterviewFilterForm(student_id_List);
                    Form.ShowDialog();
                }
                else 
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生.");
                }                
       
            };


            // 學生家長聯繫篩選
            NLDPanels.Student.RibbonBarItems["輔導"]["匯出"]["家庭聯繫紀錄篩選"].Enable = UserAcl.Current["K12.Student.CounselStudentExport_home_visitfilter"].Executable;
            NLDPanels.Student.RibbonBarItems["輔導"]["匯出"]["家庭聯繫紀錄篩選"].Click += delegate
            {
                if (NLDPanels.Student.SelectedSource.Count > 0)
                {
                    List<string> student_id_List = Utility.GetStudentIDListByStudentID(K12.Presentation.NLDPanels.Student.SelectedSource);
                    Forms.StudHome_VisitFilterForm Form = new Forms.StudHome_VisitFilterForm(student_id_List);
                    Form.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇學生.");
                }     
           

            };


            NLDPanels.Student.RibbonBarItems["輔導"]["匯入"].Image = Properties.Resources.Import_Image;
            NLDPanels.Student.RibbonBarItems["輔導"]["匯入"].Size = RibbonBarButton.MenuButtonSize.Large;


            // 匯入優先關懷
            NLDPanels.Student.RibbonBarItems["輔導"]["匯入"]["匯入優先關懷"].Enable = UserAcl.Current["K12.Student.CounselStudentImportCareRecord"].Executable;
            NLDPanels.Student.RibbonBarItems["輔導"]["匯入"]["匯入優先關懷"].Click += delegate
            {

                Global._AllStudentNumberStatusIDTemp = Utility.GetAllStudenNumberStatusDict();
                Global._AllTeacherNameIdDictTemp = Utility.GetAllTeacherNameIDDict();
                Global._StudentStatusDBDict = Utility.GetStudentStatusDBValDict();
                ImportExport.ImportStudentCareRecord iscr = new ImportExport.ImportStudentCareRecord();
                iscr.Execute();
            };


            // 匯入個案會議
            NLDPanels.Student.RibbonBarItems["輔導"]["匯入"]["匯入個案會議"].Enable = UserAcl.Current["K12.Student.CounselStudentImportCaseMeetingRecord"].Executable;
            NLDPanels.Student.RibbonBarItems["輔導"]["匯入"]["匯入個案會議"].Click += delegate
            {
                Global._AllStudentNumberStatusIDTemp = Utility.GetAllStudenNumberStatusDict();
                Global._AllTeacherNameIdDictTemp = Utility.GetAllTeacherNameIDDict();
                Global._StudentStatusDBDict = Utility.GetStudentStatusDBValDict();
                ImportExport.ImportStudentCaseMeetingRecord iscmr = new ImportExport.ImportStudentCaseMeetingRecord();
                iscmr.Execute();
            };

            // 匯入晤談紀錄
            NLDPanels.Student.RibbonBarItems["輔導"]["匯入"]["匯入晤談紀錄"].Enable = UserAcl.Current["K12.Student.CounselStudentImportInterViewData"].Executable;
            NLDPanels.Student.RibbonBarItems["輔導"]["匯入"]["匯入晤談紀錄"].Click += delegate
            {
                //2016/11/25 穎驊筆記，因應輔導2.0 table counsel.interview_record 調整，牽連到的功能先暫時關閉，日後調整再開啟。

                MsgBox.Show("晤談紀錄系統調整中，功能暫不開放", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

                //Global._AllStudentNumberStatusIDTemp = Utility.GetAllStudenNumberStatusDict();
                //Global._AllTeacherNameIdDictTemp = Utility.GetAllTeacherNameIDDict();
                //Global._StudentStatusDBDict = Utility.GetStudentStatusDBValDict();
                //ImportExport.ImportStudentInterViewData isivd = new ImportExport.ImportStudentInterViewData();
                //isivd.Execute();
            };

            // 匯入輔導自訂欄位
            NLDPanels.Student.RibbonBarItems["輔導"]["匯入"]["匯入輔導自訂欄位"].Enable = UserAcl.Current["K12.Student.CounselStudentImportUserDefineData"].Executable;
            NLDPanels.Student.RibbonBarItems["輔導"]["匯入"]["匯入輔導自訂欄位"].Click += delegate
            {
                Global._AllStudentNumberStatusIDTemp = Utility.GetAllStudenNumberStatusDict();
                Global._StudentStatusDBDict = Utility.GetStudentStatusDBValDict();
                ImportExport.ImportStudentUserDefData isudd = new ImportExport.ImportStudentUserDefData();
                isudd.Execute();
            };

            // 管理開放時間
            RibbonBarItem rbItemAB = MotherForm.RibbonBarItems["學生", "輔導"];// MotherForm.RibbonBarItems["輔導", "綜合資料表"];
            rbItemAB["設定"].Items["設定綜合紀錄表開放時間"].Enable = UserAcl.Current["K12.Student.SetABCardAccessStartingDate"].Executable;
            rbItemAB["設定"].Items["設定綜合紀錄表開放時間"].Click += delegate
            {
                Forms.SetABCardAccessStartingDate sabcsd = new Forms.SetABCardAccessStartingDate();
                sabcsd.ShowDialog();

            };

            // 管理開放時間
            Catalog catalogAB = RoleAclSource.Instance["輔導"]["功能按鈕"];
            catalogAB.Add(new RibbonFeature("K12.Student.SetABCardAccessStartingDate", "管理開放時間"));
        }

        static void _bkWork_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            _hasABTemplate = true;
            if (Global._ErrorMessageList.Length > 0)
            {
                FISCA.Presentation.Controls.MsgBox.Show(Global._ErrorMessageList.ToString());
            }
        }

        static void _bkWork_DoWork(object sender, DoWorkEventArgs e)
        {
            // 載入綜合表現題目檢查
            //DAO.ABCardQuestionDataManager man = new DAO.ABCardQuestionDataManager();           


            try
            {

                // 檢查與清空綜合表現題目 (Beta 用，修正題目用
                int cot = DAO.UDTTransfer.ABUDTQuestionsDataSelectAll().Count;
                if (cot < 80 || cot > 120)
                    DAO.UDTTransfer.ABUDTQuestionsDataDeleteAll();

                // 更新綜合紀錄表題目
                DAO.UDTTransfer.CreateCounselUDTTable();
                Utility.UpdateABQuestions();

                // 更新 UDS UDT 方式            
                if (!FISCA.RTContext.IsDiagMode)
                    FISCA.ServerModule.AutoManaged("http://module.ischool.com.tw:8080/module/137/CounselSystem_dep/udm.xml");


                #region 自訂驗證規則
                FactoryProvider.RowFactory.Add(new ValidationRule.CounselRowValidatorFactory());
                #endregion

                //// 檢查是否有樣板
                //List<DAO.UDT_ABCardTemplateDefinitionDef> ABCardTemplate = DAO.UDTTransfer.GetABCardTemplate();
                //// 沒有樣板時
                //if (ABCardTemplate.Count == 0)
                //{
                //    List<DAO.UDT_ABCardTemplateDefinitionDef> insertUDT = new List<DAO.UDT_ABCardTemplateDefinitionDef>();
                //    XElement elmRoot = XElement.Parse(Properties.Resources.ABCardTemplate);
                //    foreach (XElement elm in elmRoot.Elements("Subject"))
                //    {
                //        bool checkInsert = true;

                //        foreach (DAO.UDT_ABCardTemplateDefinitionDef rec in ABCardTemplate)
                //            if (rec.SubjectName.Trim() == elm.Attribute("label").Value.Trim())
                //                checkInsert = false;

                //        if (checkInsert)
                //        {
                //            DAO.UDT_ABCardTemplateDefinitionDef abRec = new DAO.UDT_ABCardTemplateDefinitionDef();
                //            abRec.SubjectName = elm.Attribute("label").Value.Trim();
                //            abRec.Content = elm.ToString();
                //            int i;
                //            if (int.TryParse(abRec.SubjectName.Substring(0, 1), out i))
                //                abRec.Priority = i;

                //            insertUDT.Add(abRec);
                //        }
                //    }
                //    if (insertUDT.Count > 0)
                //        DAO.UDTTransfer.InsertABCardTemplate(insertUDT);
                //}

                //// 將樣板傳入
                //Global._ABCardTemplateTransfer.LoadAllTemplate(ABCardTemplate);

                // 檢查是否有輔導相關設定標籤，沒有自動加入:輔導:認輔老師,輔導:輔導主任,輔導:輔導老師,
                List<TagConfigRecord> tagList = TagConfig.SelectByCategoryAndPrefix(TagCategory.Teacher, "輔導");
                List<string> pNameList = new List<string>();
                pNameList.Add("認輔老師");
                pNameList.Add("輔導主任");
                pNameList.Add("輔導老師");

                foreach (string str in pNameList)
                {
                    bool chkeckAdd = true;
                    foreach (TagConfigRecord tt in tagList)
                        if (tt.Name == str && tt.Prefix == "輔導")
                            chkeckAdd = false;

                    if (chkeckAdd)
                    {
                        TagConfigRecord TaRec = new TagConfigRecord();
                        TaRec.Prefix = "輔導";
                        TaRec.Name = str;
                        TaRec.Category = "Teacher";
                        TagConfig.Insert(TaRec);
                    }
                }

            }
            catch (Exception ex)
            {
                Global._ErrorMessageList.AppendLine("載入輔導系統發生錯誤："+ex.Message);
            }
        }//

    }
}
