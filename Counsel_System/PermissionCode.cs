using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2
{
    /// <summary>
    /// 輔導相關共用權限代碼
    /// </summary>
    class PermissionCode
    {
        public const string 輔導相關測驗_資料項目="K12.Student.StudQuizDataContent";
        public const string 輔導晤談紀錄_資料項目 ="K12.Student.StudInterviewDataContent";
        public const string 輔導個案會議_資料項目 ="K12.Student.StudCaseMeetingRecordContent";
        public const string 輔導優先關懷紀錄_資料項目 ="K12.Student.StudCareRecordContent";
        public const string 輔導自訂欄位_資料項目 ="K12.CounselStudentUserDefineDataContent";
        public const string 認輔老師及輔導老師_資料項目 ="K12.Student.CounselStudentListContent";
        public const string 輔導學生_資料項目 ="K12.Teacher.CounselStudentListBContent";
        public const string 認輔學生_資料項目 ="K12.Teacher.CounselStudentListAContent";
        public const string 綜合表現紀錄表_資料項目 ="K12.Student.CounselStudentABCard";
        public const string 綜合表現紀錄表_題目管理 = "K12.Student.CounselStudentABCardQuestionManager";
    }
}
