using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Counsel_System2
{
    public partial class CounselAdmin : FISCA.Presentation.BlankPanel
    {
        public CounselAdmin()
        {
            InitializeComponent();
            Group = "輔導系統";
            //if (!helpContentPanel1.NavigateBySetting(Group))
            //    helpContentPanel1.Naviate("https://sites.google.com/a/ischool.com.tw/ischool-jh/home/ischool-jh-jiao-wu-xiang-guan-fu-zhu-shuo-ming");
        }
        private static CounselAdmin _Counsel_Admin;

        public static CounselAdmin Instance
        {
            get { if(_Counsel_Admin == null) _Counsel_Admin = new CounselAdmin (); return _Counsel_Admin;}

        }
    }
}
