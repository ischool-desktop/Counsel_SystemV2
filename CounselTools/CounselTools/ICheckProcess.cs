using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CounselTools
{
    /// <summary>
    /// 檢查綜合紀錄表使用
    /// </summary>
    public interface ICheckProcess
    {
        void SetGroupName(string GroupName);

        void SetStudent(ClassStudent Student);

        Dictionary<string, string> GetErrorData();
        int GetErrorCount();
        int GetTotalCount();
         void Start();

         string GetMessage();
    }
}
