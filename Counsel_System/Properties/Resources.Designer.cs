﻿//------------------------------------------------------------------------------
// <auto-generated>
//     這段程式碼是由工具產生的。
//     執行階段版本:4.0.30319.42000
//
//     對這個檔案所做的變更可能會造成錯誤的行為，而且如果重新產生程式碼，
//     變更將會遺失。
// </auto-generated>
//------------------------------------------------------------------------------

namespace Counsel_System2.Properties {
    using System;
    
    
    /// <summary>
    ///   用於查詢當地語系化字串等的強類型資源類別。
    /// </summary>
    // 這個類別是自動產生的，是利用 StronglyTypedResourceBuilder
    // 類別透過 ResGen 或 Visual Studio 這類工具。
    // 若要加入或移除成員，請編輯您的 .ResX 檔，然後重新執行 ResGen
    // (利用 /str 選項)，或重建您的 VS 專案。
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   傳回這個類別使用的快取的 ResourceManager 執行個體。
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("Counsel_System2.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   覆寫目前執行緒的 CurrentUICulture 屬性，對象是所有
        ///   使用這個強類型資源類別的資源查閱。
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   查詢類似 &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;
        ///&lt;Subjects&gt;
        ///  &lt;Subject label=&quot;1.本人概況&quot;&gt;
        ///    &lt;!-- 大題 --&gt;
        ///    &lt;QG label=&quot;血型&quot;&gt;
        ///      &lt;!-- 題組 --&gt;
        ///      &lt;Choices&gt;
        ///        &lt;!-- 選項清單 --&gt;
        ///        &lt;Item label=&quot;A&quot; /&gt;
        ///        &lt;Item label=&quot;B&quot; /&gt;
        ///        &lt;Item label=&quot;O&quot; selected=&quot;true&quot; /&gt;
        ///        &lt;!-- default selected --&gt;
        ///        &lt;Item label=&quot;AB&quot; /&gt;
        ///        &lt;Item label=&quot;其它&quot; hasText=&quot;true&quot; /&gt;
        ///      &lt;/Choices&gt;
        ///      &lt;Qs&gt;
        ///        &lt;!-- 題目清單 --&gt;
        ///        &lt;Q type=&quot;combobox&quot; name=&quot;AAA10000001&quot; label=&quot;&quot; /&gt;
        ///        &lt;!-- 題目 --&gt;
        ///      &lt;/Qs&gt;
        ///  [字串的其餘部分已遭截斷]&quot;; 的當地語系化字串。
        /// </summary>
        internal static string ABCardTemplate {
            get {
                return ResourceManager.GetString("ABCardTemplate", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查詢類似 &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;
        ///&lt;ABCardTransferDataOldMapping&gt;
        ///  &lt;Item GroupName=&quot;1.本人概況&quot; QuestionName=&quot;血型&quot; Key=&quot;AAA10000001&quot; AType=&quot;&quot; TableName=&quot;本人概況&quot; ColumnName=&quot;血型&quot; /&gt;
        ///  &lt;Item GroupName=&quot;1.本人概況&quot; QuestionName=&quot;宗教&quot; Key=&quot;AAA10000002&quot; AType=&quot;&quot; TableName=&quot;本人概況&quot; ColumnName=&quot;宗教&quot; /&gt;
        ///  &lt;Item GroupName=&quot;1.本人概況&quot; QuestionName=&quot;身高一上&quot; Key=&quot;AAA10000005&quot; AType=&quot;&quot; TableName=&quot;本人概況&quot; ColumnName=&quot;身高一上&quot; /&gt;
        ///  &lt;Item GroupName=&quot;1.本人概況&quot; QuestionName=&quot;身高二上&quot; Key=&quot;AAA10000006&quot; AType=&quot;&quot; TableName=&quot;本人概況&quot; ColumnName=&quot;身高二上&quot; /&gt;
        ///   [字串的其餘部分已遭截斷]&quot;; 的當地語系化字串。
        /// </summary>
        internal static string ABCardTransferDataOldMapping {
            get {
                return ResourceManager.GetString("ABCardTransferDataOldMapping", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        /// </summary>
        internal static System.Drawing.Bitmap Export_Image {
            get {
                object obj = ResourceManager.GetObject("Export_Image", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        /// </summary>
        internal static System.Drawing.Bitmap Import_Image {
            get {
                object obj = ResourceManager.GetObject("Import_Image", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   查詢類似 &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot; ?&gt;
        ///&lt;?xml-stylesheet type=&quot;text/xsl&quot; href=&quot;format.xsl&quot; ?&gt;
        ///&lt;ValidateRule Name=&quot;學生優先關懷&quot;&gt;
        ///  &lt;DuplicateDetection&gt;
        ///    &lt;Detector Name=&quot;PrimaryKey1&quot;&gt;
        ///      &lt;Field Name=&quot;學號&quot; /&gt;
        ///	    &lt;Field Name=&quot;狀態&quot; /&gt;
        ///      &lt;Field Name=&quot;立案日期&quot; /&gt;
        ///      &lt;Field Name=&quot;個案類別&quot; /&gt;
        ///    &lt;/Detector&gt;
        ///  &lt;/DuplicateDetection&gt;
        ///  &lt;FieldList&gt;
        ///
        ///    &lt;Field Required=&quot;False&quot; Name=&quot;代號&quot;&gt;      
        ///    &lt;/Field&gt;
        ///    
        ///    &lt;Field Required=&quot;True&quot; Name=&quot;學號&quot;&gt;
        ///      &lt;Validate AutoCorrect=&quot;False&quot; Description=&quot;「學號」不允許空白。&quot; ErrorTyp [字串的其餘部分已遭截斷]&quot;; 的當地語系化字串。
        /// </summary>
        internal static string ImportStudentCareRecordVal {
            get {
                return ResourceManager.GetString("ImportStudentCareRecordVal", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查詢類似 &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot; ?&gt;
        ///&lt;?xml-stylesheet type=&quot;text/xsl&quot; href=&quot;format.xsl&quot; ?&gt;
        ///&lt;ValidateRule Name=&quot;學生個案會議&quot;&gt;
        ///  &lt;DuplicateDetection&gt;
        ///    &lt;Detector Name=&quot;PrimaryKey1&quot;&gt;
        ///      &lt;Field Name=&quot;學號&quot; /&gt;
        ///  	  &lt;Field Name=&quot;狀態&quot; /&gt;
        ///      &lt;Field Name=&quot;會議日期&quot; /&gt;
        ///      &lt;Field Name=&quot;會議事由&quot;/&gt;
        ///    &lt;/Detector&gt;
        ///  &lt;/DuplicateDetection&gt;
        ///  &lt;FieldList&gt;
        ///
        ///    &lt;Field Required=&quot;True&quot; Name=&quot;學號&quot;&gt;
        ///      &lt;Validate AutoCorrect=&quot;False&quot; Description=&quot;「學號」不允許空白。&quot; ErrorType=&quot;Error&quot; Validator=&quot;不可空白&quot; When=&quot;&quot; /&gt;
        ///    &lt;/Field&gt;
        ///
        ///    &lt;Field R [字串的其餘部分已遭截斷]&quot;; 的當地語系化字串。
        /// </summary>
        internal static string ImportStudentCaseMeetingRecordVal {
            get {
                return ResourceManager.GetString("ImportStudentCaseMeetingRecordVal", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查詢類似 &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot; ?&gt;
        ///&lt;?xml-stylesheet type=&quot;text/xsl&quot; href=&quot;format.xsl&quot; ?&gt;
        ///&lt;ValidateRule Name=&quot;晤談紀錄&quot;&gt;
        ///  &lt;DuplicateDetection&gt;
        ///    &lt;Detector Name=&quot;PrimaryKey1&quot;&gt;
        ///      &lt;Field Name=&quot;學號&quot; /&gt;
        ///	    &lt;Field Name=&quot;狀態&quot; /&gt;
        ///      &lt;Field Name=&quot;日期&quot; /&gt;      
        ///      &lt;Field Name=&quot;晤談事由&quot; /&gt;
        ///    &lt;/Detector&gt;
        ///  &lt;/DuplicateDetection&gt;
        ///  &lt;FieldList&gt;
        ///
        ///    &lt;Field Required=&quot;True&quot; Name=&quot;學號&quot;&gt;
        ///      &lt;Validate AutoCorrect=&quot;False&quot; Description=&quot;「學號」不允許空白。&quot; ErrorType=&quot;Error&quot; Validator=&quot;不可空白&quot; When=&quot;&quot; /&gt;
        ///    &lt;/Field&gt;
        ///
        ///    &lt;Fiel [字串的其餘部分已遭截斷]&quot;; 的當地語系化字串。
        /// </summary>
        internal static string ImportStudentInterViewDataVal {
            get {
                return ResourceManager.GetString("ImportStudentInterViewDataVal", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查詢類似 &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot; ?&gt;
        ///&lt;?xml-stylesheet type=&quot;text/xsl&quot; href=&quot;format.xsl&quot; ?&gt;
        ///&lt;ValidateRule Name=&quot;學生使用者自訂欄位&quot;&gt;
        ///  &lt;DuplicateDetection&gt;
        ///    &lt;Detector Name=&quot;PrimaryKey1&quot;&gt;
        ///      &lt;Field Name=&quot;學號&quot; /&gt;
        ///      &lt;Field Name=&quot;狀態&quot; /&gt;
        ///      &lt;Field Name=&quot;欄位名稱&quot; /&gt;
        ///    &lt;/Detector&gt;
        ///  &lt;/DuplicateDetection&gt;
        ///  &lt;FieldList&gt;
        ///
        ///    &lt;Field Required=&quot;True&quot; Name=&quot;學號&quot;&gt;
        ///      &lt;Validate AutoCorrect=&quot;False&quot; Description=&quot;「學號」不允許空白。&quot; ErrorType=&quot;Error&quot; Validator=&quot;不可空白&quot; When=&quot;&quot; /&gt;
        ///    &lt;/Field&gt;
        ///
        ///    &lt;Field Required=&quot;True&quot; Name=&quot;狀態 [字串的其餘部分已遭截斷]&quot;; 的當地語系化字串。
        /// </summary>
        internal static string ImportStudentUserDefDataVal {
            get {
                return ResourceManager.GetString("ImportStudentUserDefDataVal", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查詢類似 &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot; ?&gt;
        ///&lt;?xml-stylesheet type=&quot;text/xsl&quot; href=&quot;format.xsl&quot; ?&gt;
        ///&lt;ValidateRule Name=&quot;學生測驗資料&quot;&gt;
        ///  &lt;DuplicateDetection&gt;
        ///    &lt;Detector Name=&quot;PrimaryKey1&quot;&gt;
        ///      &lt;Field Name=&quot;班級&quot; /&gt;
        ///      &lt;Field Name=&quot;座號&quot; /&gt;
        ///      &lt;Field Name=&quot;狀態&quot; /&gt;
        ///    &lt;/Detector&gt;
        ///  &lt;/DuplicateDetection&gt;
        ///  &lt;FieldList&gt;
        ///    &lt;Field Required=&quot;True&quot; Name=&quot;班級&quot;&gt;
        ///      &lt;Validate AutoCorrect=&quot;False&quot; Description=&quot;「班級」不允許空白。&quot; ErrorType=&quot;Error&quot; Validator=&quot;不可空白&quot; When=&quot;&quot; /&gt;
        ///    &lt;/Field&gt;
        ///    &lt;Field Required=&quot;True&quot; Name=&quot;座號&quot;&gt;
        ///     [字串的其餘部分已遭截斷]&quot;; 的當地語系化字串。
        /// </summary>
        internal static string ImportStudQuizDataVa_CSeatNo {
            get {
                return ResourceManager.GetString("ImportStudQuizDataVa_CSeatNo", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查詢類似 &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot; ?&gt;
        ///&lt;?xml-stylesheet type=&quot;text/xsl&quot; href=&quot;format.xsl&quot; ?&gt;
        ///&lt;ValidateRule Name=&quot;學生測驗資料&quot;&gt;
        ///  &lt;DuplicateDetection&gt;
        ///    &lt;Detector Name=&quot;PrimaryKey1&quot;&gt;
        ///      &lt;Field Name=&quot;學號&quot; /&gt;
        ///      &lt;Field Name=&quot;狀態&quot; /&gt;
        ///    &lt;/Detector&gt;
        ///  &lt;/DuplicateDetection&gt;
        ///  &lt;FieldList&gt;
        ///    &lt;Field Required=&quot;True&quot; Name=&quot;學號&quot;&gt;
        ///      &lt;Validate AutoCorrect=&quot;False&quot; Description=&quot;「學號」不允許空白。&quot; ErrorType=&quot;Error&quot; Validator=&quot;不可空白&quot; When=&quot;&quot; /&gt;
        ///    &lt;/Field&gt;
        ///    &lt;Field Required=&quot;True&quot; Name=&quot;狀態&quot;&gt;
        ///      &lt;Validate AutoCorrect=&quot;F [字串的其餘部分已遭截斷]&quot;; 的當地語系化字串。
        /// </summary>
        internal static string ImportStudQuizDataVal_SNum {
            get {
                return ResourceManager.GetString("ImportStudQuizDataVal_SNum", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        /// </summary>
        internal static System.Drawing.Bitmap parent_coordinator_next_64 {
            get {
                object obj = ResourceManager.GetObject("parent_coordinator_next_64", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   查詢類似 &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;
        ///&lt;Questions&gt;
        ///  &lt;Question&gt;
        ///    &lt;Group&gt;本人概況&lt;/Group&gt;
        ///    &lt;Name&gt;血型&lt;/Name&gt;
        ///    &lt;QuestionType&gt;SINGLE_ANSWER&lt;/QuestionType&gt;
        ///    &lt;ControlType&gt;RADIO_BUTTON&lt;/ControlType&gt;
        ///    &lt;CanPrint&gt;True&lt;/CanPrint&gt;
        ///    &lt;CanTeacherEdit&gt;True&lt;/CanTeacherEdit&gt;
        ///    &lt;CanStudentEdit&gt;True&lt;/CanStudentEdit&gt;
        ///    &lt;DisplayOrder&gt;1&lt;/DisplayOrder&gt;
        ///    &lt;Items&gt;
        ///      &lt;item key=&quot;A&quot; has_remark=&quot;False&quot; /&gt;
        ///      &lt;item key=&quot;B&quot; has_remark=&quot;False&quot; /&gt;
        ///      &lt;item key=&quot;O&quot; has_remark=&quot;False&quot; /&gt;
        ///      &lt;item key=&quot;AB&quot; has_r [字串的其餘部分已遭截斷]&quot;; 的當地語系化字串。
        /// </summary>
        internal static string Questions {
            get {
                return ResourceManager.GetString("Questions", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        /// </summary>
        internal static System.Drawing.Bitmap Report {
            get {
                object obj = ResourceManager.GetObject("Report", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   查詢類型 System.Byte[] 的當地語系化資源。
        /// </summary>
        internal static byte[] 學生晤談紀錄篩選 {
            get {
                object obj = ResourceManager.GetObject("學生晤談紀錄篩選", resourceCulture);
                return ((byte[])(obj));
            }
        }
        
        /// <summary>
        ///   查詢類型 System.Byte[] 的當地語系化資源。
        /// </summary>
        internal static byte[] 學生晤談記錄表_新範本 {
            get {
                object obj = ResourceManager.GetObject("學生晤談記錄表_新範本", resourceCulture);
                return ((byte[])(obj));
            }
        }
        
        /// <summary>
        ///   查詢類型 System.Byte[] 的當地語系化資源。
        /// </summary>
        internal static byte[] 晤談紀錄簽認表樣版 {
            get {
                object obj = ResourceManager.GetObject("晤談紀錄簽認表樣版", resourceCulture);
                return ((byte[])(obj));
            }
        }
        
        /// <summary>
        ///   查詢類型 System.Byte[] 的當地語系化資源。
        /// </summary>
        internal static byte[] 綜合表現紀錄表_樣版 {
            get {
                object obj = ResourceManager.GetObject("綜合表現紀錄表_樣版", resourceCulture);
                return ((byte[])(obj));
            }
        }
        
        /// <summary>
        ///   查詢類型 System.Byte[] 的當地語系化資源。
        /// </summary>
        internal static byte[] 綜合表現紀錄表合併欄位說明 {
            get {
                object obj = ResourceManager.GetObject("綜合表現紀錄表合併欄位說明", resourceCulture);
                return ((byte[])(obj));
            }
        }
        
        /// <summary>
        ///   查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        /// </summary>
        internal static System.Drawing.Bitmap 設定 {
            get {
                object obj = ResourceManager.GetObject("設定", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   查詢類型 System.Byte[] 的當地語系化資源。
        /// </summary>
        internal static byte[] 輔導紀錄_樣板 {
            get {
                object obj = ResourceManager.GetObject("輔導紀錄_樣板", resourceCulture);
                return ((byte[])(obj));
            }
        }
        
        /// <summary>
        ///   查詢類型 System.Byte[] 的當地語系化資源。
        /// </summary>
        internal static byte[] 輔導資料紀錄表範本 {
            get {
                object obj = ResourceManager.GetObject("輔導資料紀錄表範本", resourceCulture);
                return ((byte[])(obj));
            }
        }
    }
}
