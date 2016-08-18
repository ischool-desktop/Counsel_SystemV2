using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using System.Data;
using System.Xml.Linq;
using K12.Data;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// word 資料與樣板合併
    /// </summary>
    public class DocumentMerge
    {
        // 樣板
        byte[] _DocTemplate;
        BackgroundWorker _bgWork;
        // 檢查是否儲存後開啟
        bool _CheckSaveOpen = false;
        // 合併欄位名稱
        List<string> _MergeNameList;
        bool _pass = true;
        string _FilePathAndName = "";
        List<Dictionary<string, string>> _DocDataList;

        private DocumentBuilder _builder;
        Document docTemplate;
               /// <summary>
        /// Word文件合併
        /// </summary>
        /// <param name="合併項目名稱"></param>
        /// <param name="合併資料"></param>
        /// <param name="合併Word樣板"></param>
        /// <param name="儲存路徑"></param>
        /// <param name="是否直接開啟"></param>
        public DocumentMerge(List<string> MergeNameList, List<Dictionary<string, string>> DocDataList, byte[] WordTemplateStream, string SaveFilePathAndName, bool OpenFile)
        {
            try
            {
                _DocTemplate = WordTemplateStream;

                if (_DocTemplate == null)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("樣板解析失敗!");
                    return;
                }
                _FilePathAndName = SaveFilePathAndName;

                if (string.IsNullOrEmpty(_FilePathAndName))
                    _FilePathAndName = Application.StartupPath + "\\Doc1.doc";
                else
                {
                    // 檢查是否有副檔名
                    if (_FilePathAndName.ToLower().IndexOf(".DOC") < 1)
                        _FilePathAndName = _FilePathAndName + ".doc";
                }

                _CheckSaveOpen = OpenFile;
                _MergeNameList = new List<string>();
                _DocDataList = DocDataList;
                foreach (string str in MergeNameList)
                    if (!_MergeNameList.Contains(str))
                        _MergeNameList.Add(str);

                if (_MergeNameList == null)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("沒有合併名稱!");
                    return;               
                }

                if (_DocDataList == null)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("沒有合併資料!");
                    return;                
                }

            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show(ex.Message + "無法讀取樣板檔。");
                return;
            }        
        }

                /// <summary>
        /// 執行合併資料並儲存
        /// </summary>
        /// <returns></returns>
        public bool Merge()
        {
            _bgWork = new BackgroundWorker();
            _bgWork.DoWork += new DoWorkEventHandler(_bgWork_DoWork);
            _bgWork.RunWorkerCompleted+=new RunWorkerCompletedEventHandler(_bgWork_RunWorkerCompleted);
            _bgWork.RunWorkerAsync();
            return _pass;
        }

        void _bgWork_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Document doc = (Document)e.Result;

            // 當沒有設定檔案名稱
            if (string.IsNullOrEmpty(_FilePathAndName))
            {
                SaveFileDialog sd = new SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = "Doc1.doc";
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        doc.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                        if (_CheckSaveOpen)
                            System.Diagnostics.Process.Start(sd.FileName);
                    }
                    catch
                    {
                        MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        _pass = false;
                        return;
                    }
                }
            }
            else
            {
                try
                {
                    doc.Save(_FilePathAndName, SaveFormat.Doc);
                    if (_CheckSaveOpen)
                        System.Diagnostics.Process.Start(_FilePathAndName);
                }
                catch
                {
                    SaveFileDialog sd = new SaveFileDialog();
                    sd.Title = "另存新檔";
                    sd.FileName = "Doc1.doc";
                    sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                    if (sd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            doc.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                            System.Diagnostics.Process.Start(sd.FileName);
                        }
                        catch
                        {
                            MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            _pass = false;
                            return;
                        }
                    }
                }
            }
         
        }

        void _bgWork_DoWork(object sender, DoWorkEventArgs e)
        {
            Document doc = new Document();
            doc.Sections.Clear();
          

            // 比對欄位名稱放值
            foreach (Dictionary<string, string> data in _DocDataList)
            {
                DataTable dt = new DataTable();
                // 建立欄位名稱
                foreach (string name in _MergeNameList)
                    dt.Columns.Add(name, typeof(string));

                DataRow dr = dt.NewRow();
                foreach (string name in _MergeNameList)
                {
                    if (data.ContainsKey(name))
                        dr[name] = data[name];
                }
                dt.Rows.Add(dr);
                docTemplate = new Document(new MemoryStream(_DocTemplate));
                //_builder = new DocumentBuilder(docTemplate);
                docTemplate.MailMerge.FieldMergingCallback = new InsertDocumentAtMailMergeHandler();
                //docTemplate.MailMerge.MergeField += new Aspose.Words.Reporting.MergeFieldEventHandler(MailMerge_MergeField);
                docTemplate.MailMerge.RemoveEmptyParagraphs = true;
                docTemplate.MailMerge.Execute(dt);
                docTemplate.MailMerge.DeleteFields();
                doc.Sections.Add(doc.ImportNode(docTemplate.Sections[0], true));
            }
            e.Result = doc;
        }

        private class InsertDocumentAtMailMergeHandler : IFieldMergingCallback
        {

            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
               try
                {
                    DocumentBuilder _builder = new DocumentBuilder(e.Document);
                    // 這段在處理單純 List 要產生動態表格
                    if (e.FieldName.IndexOf("GL_") > -1 && e.FieldValue.ToString() != "")
                    {
                        if (_builder.MoveToMergeField(e.FieldName))
                        {
                            Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                            cell.CellFormat.LeftPadding = 0;
                            cell.CellFormat.RightPadding = 0;
                            double width = cell.CellFormat.Width;
                            int columnCount = 1;
                            double miniUnitWitdh = width / (double)columnCount;

                            Table table = _builder.StartTable();

                            //(table.ParentNode.ParentNode as Row).RowFormat.LeftIndent = 0;
                            double p = _builder.RowFormat.LeftIndent;
                            _builder.RowFormat.HeightRule = HeightRule.Exactly;
                            _builder.RowFormat.Height = 18.0;
                            _builder.RowFormat.LeftIndent = 0;
                            int count = 0;
                            // 答案
                            List<string> dataList = e.FieldValue.ToString().Split(';').ToList();
                            if (dataList != null)
                            {
                                foreach (string data in dataList)
                                {
                                    Cell c1 = _builder.InsertCell();
                                    c1.CellFormat.Width = miniUnitWitdh;
                                    c1.CellFormat.WrapText = true;

                                    _builder.Write(data);
                                    _builder.EndRow();
                                }

                                _builder.EndTable();

                                //去除表格四邊的線
                                foreach (Cell c in table.FirstRow.Cells)
                                    c.CellFormat.Borders.Top.LineStyle = LineStyle.None;

                                foreach (Cell c in table.LastRow.Cells)
                                    c.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;

                                foreach (Row r in table.Rows)
                                {
                                    r.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                    r.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                }
                            }
                            _builder.RowFormat.LeftIndent = p;
                        }
                    }

                    // 這段主要在處理解析XML結構與相對應產生動態表格
                    if (e.FieldName.IndexOf("GR_") > -1 && e.FieldValue.ToString() != "")
                    {

                        XElement elmData = XElement.Parse(e.FieldValue.ToString());

                        List<string> ColNames = new List<string>();
                        List<Dictionary<string, string>> DataDict = new List<Dictionary<string, string>>();
                        foreach (XElement elm1 in elmData.Elements("Item"))
                        {
                            Dictionary<string, string> dict = new Dictionary<string, string>();
                            foreach (XElement elm2 in elm1.Elements("Field"))
                            {
                                if (!ColNames.Contains(elm2.Attribute("key").Value))
                                    ColNames.Add(elm2.Attribute("key").Value);
                                if (!dict.ContainsKey(elm2.Attribute("key").Value))
                                    dict.Add(elm2.Attribute("key").Value, elm2.Attribute("value").Value);
                            }
                            DataDict.Add(dict);
                        }


                        if (_builder.MoveToMergeField(e.FieldName))
                        {

                            Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                            cell.CellFormat.LeftPadding = 0;
                            cell.CellFormat.RightPadding = 0;
                            double width = cell.CellFormat.Width;
                            int columnCount = ColNames.Count;
                            double miniUnitWitdh = width / (double)columnCount;

                            Table table = _builder.StartTable();
                            double li = _builder.RowFormat.LeftIndent;
                            _builder.RowFormat.HeightRule = HeightRule.Exactly;
                            _builder.RowFormat.Height = 18.0;
                            _builder.RowFormat.LeftIndent = 0;
                            if (ColNames.Count > 0)
                            {
                                // 欄位名稱
                                foreach (string name in ColNames)
                                {
                                    Cell c1 = _builder.InsertCell();
                                    c1.CellFormat.Width = miniUnitWitdh;
                                    c1.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                    _builder.Write(name);
                                }
                                _builder.EndRow();
                            }


                            // 答案
                            if (DataDict.Count > 0)
                            {
                                foreach (Dictionary<string, string> data in DataDict)
                                {
                                    foreach (KeyValuePair<string, string> da in data)
                                    {
                                        Cell c2 = _builder.InsertCell();
                                        c2.CellFormat.Width = miniUnitWitdh;
                                        c2.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                        _builder.Write(da.Value);

                                    }
                                    _builder.EndRow();

                                }
                                _builder.EndTable();

                                //去除表格四邊的線
                                foreach (Cell c in table.FirstRow.Cells)
                                    c.CellFormat.Borders.Top.LineStyle = LineStyle.None;

                                foreach (Cell c in table.LastRow.Cells)
                                    c.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;

                                foreach (Row r in table.Rows)
                                {
                                    r.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                    r.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                }
                            }
                            _builder.RowFormat.LeftIndent = li;
                        }
                    }

                    #region 學習歷程資料讀取

                    // 學期歷程資料讀取
                    if ((e.FieldName == "學習歷程" || e.FieldName == "學期對照") && e.FieldValue.ToString() != "")
                    {
                        string studID = e.FieldValue.ToString();
                        if (Global._StudentSemesterHistoryItemDict.ContainsKey(studID))
                        {
                            if (_builder.MoveToMergeField(e.FieldName))
                            {
                                Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                                cell.CellFormat.LeftPadding = 0;
                                cell.CellFormat.RightPadding = 0;
                                double width = cell.CellFormat.Width;
                                int columnCount = 4;
                                double miniUnitWitdh = width / (double)columnCount;

                                Table table = _builder.StartTable();
                                double li = _builder.RowFormat.LeftIndent;
                                _builder.RowFormat.HeightRule = HeightRule.Exactly;
                                _builder.RowFormat.Height = 18.0;
                                _builder.RowFormat.LeftIndent = 0;

                                // 欄位名稱
                                Cell tc1 = _builder.InsertCell();
                                tc1.CellFormat.Width = miniUnitWitdh;
                                tc1.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                _builder.Write("年級學期");

                                Cell tc2 = _builder.InsertCell();
                                tc2.CellFormat.Width = miniUnitWitdh;
                                tc2.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                _builder.Write("班級");

                                Cell tc3 = _builder.InsertCell();
                                tc3.CellFormat.Width = miniUnitWitdh;
                                tc3.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                _builder.Write("座號");

                                Cell tc4 = _builder.InsertCell();
                                tc4.CellFormat.Width = miniUnitWitdh;
                                tc4.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                _builder.Write("導師");
                                _builder.EndRow();


                                // 學習歷程內容
                                if (Global._StudentSemesterHistoryItemDict[studID].Count > 0)
                                {
                                    foreach (SemesterHistoryItem shi in Global._StudentSemesterHistoryItemDict[studID])
                                    {

                                        // 年級學期
                                        Cell c1 = _builder.InsertCell();
                                        c1.CellFormat.Width = miniUnitWitdh;
                                        c1.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                        _builder.Write(Utility.ConvertGradeYearSemester(shi.GradeYear, shi.Semester));

                                        // 班級
                                        Cell c2 = _builder.InsertCell();
                                        c2.CellFormat.Width = miniUnitWitdh;
                                        c2.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                        _builder.Write(shi.ClassName);

                                        // 座號
                                        Cell c3 = _builder.InsertCell();
                                        c3.CellFormat.Width = miniUnitWitdh;
                                        c3.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                        if (shi.SeatNo.HasValue)
                                            _builder.Write(shi.SeatNo.Value.ToString());
                                        else
                                            _builder.Write("");

                                        Cell c4 = _builder.InsertCell();
                                        c4.CellFormat.Width = miniUnitWitdh;
                                        c4.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                        _builder.Write(shi.Teacher);

                                        _builder.EndRow();
                                    }

                                    _builder.EndTable();

                                    //去除表格四邊的線
                                    foreach (Cell c in table.FirstRow.Cells)
                                        c.CellFormat.Borders.Top.LineStyle = LineStyle.None;

                                    foreach (Cell c in table.LastRow.Cells)
                                        c.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;

                                    foreach (Row r in table.Rows)
                                    {
                                        r.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                        r.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                    }
                                }
                                _builder.RowFormat.LeftIndent = li;
                            }
                        }
                    }
                    #endregion
                    
                    #region 導師評語(日常生活表現具體建議)
                    if ((e.FieldName == "導師評語" || e.FieldName == "具體建議") && e.FieldValue.ToString() != "")
                    {
                        string studID = e.FieldValue.ToString();
                        if (Global._ABCard_StudentTextScoreDict.ContainsKey(studID))
                        {
                            if (_builder.MoveToMergeField(e.FieldName))
                            {
                                Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                                cell.CellFormat.LeftPadding = 0;
                                cell.CellFormat.RightPadding = 0;
                                double width = cell.CellFormat.Width;
                                int columnCount = 5;
                                double miniUnitWitdh = width / (double)columnCount;

                                Table table = _builder.StartTable();
                                double li = _builder.RowFormat.LeftIndent;
                                _builder.RowFormat.HeightRule = HeightRule.Exactly;
                                _builder.RowFormat.Height = 18.0;
                                _builder.RowFormat.LeftIndent = 0;

                                // 導師評語、日常生活具體表現內容
                                if (Global._ABCard_StudentTextScoreDict[studID].Count > 0)
                                {
                                    foreach (ABCard_StudentTextScore asts in Global._ABCard_StudentTextScoreDict[studID])
                                    {
                                        // 年級學期
                                        Cell c1 = _builder.InsertCell();
                                        c1.CellFormat.Width = miniUnitWitdh;
                                        c1.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                        _builder.Write(Utility.ConvertGradeYearSemester(asts.GradeYear, asts.Semester));

                                        // 導師評語、日常生活具體表現內容
                                        Cell c2 = _builder.InsertCell();
                                        c2.CellFormat.Width = miniUnitWitdh * 4;
                                        c2.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                        if (asts.sb_Comment == "")
                                            _builder.Write(asts.DailyLifeRecommend);    // 國中小
                                        else
                                            _builder.Write(asts.DailyLifeRecommend + "," + asts.sb_Comment); // 高中部分舊制支援
                                        _builder.EndRow();
                                    }

                                    _builder.EndTable();

                                    //去除表格四邊的線
                                    foreach (Cell c in table.FirstRow.Cells)
                                        c.CellFormat.Borders.Top.LineStyle = LineStyle.None;

                                    foreach (Cell c in table.LastRow.Cells)
                                        c.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;

                                    foreach (Row r in table.Rows)
                                    {
                                        r.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                        r.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                    }
                                }
                                _builder.RowFormat.LeftIndent = li;
                            }
                        }
                    }

                    # endregion

                    #region 畢業資料
                    if (e.FieldName == "畢業資料" && e.FieldValue.ToString() != "")
                    {
                        string studID = e.FieldValue.ToString();
                        if (Global._ABCard_StudentGraduateDict.ContainsKey(studID))
                        {
                            if (_builder.MoveToMergeField(e.FieldName))
                            {
                                Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                                cell.CellFormat.LeftPadding = 0;
                                cell.CellFormat.RightPadding = 0;
                                _builder.Write("畢業學年度:" + Global._ABCard_StudentGraduateDict[studID].SchoolYear + ",畢業證書字號:" + Global._ABCard_StudentGraduateDict[studID].DiplomaNumber);
                            }
                        }
                    }
                    #endregion


                    #region 異動資料
                    if (e.FieldName == "異動資料" && e.FieldValue.ToString() != "")
                    {
                        string studID = e.FieldValue.ToString();
                        if (Global._ABCard_StudentUpdateRecDict.ContainsKey(studID))
                        {
                            if (_builder.MoveToMergeField(e.FieldName))
                            {
                                Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                                cell.CellFormat.LeftPadding = 0;
                                cell.CellFormat.RightPadding = 0;
                                double width = cell.CellFormat.Width;
                                int columnCount = 8;
                                double miniUnitWitdh = width / (double)columnCount;

                                Table table = _builder.StartTable();
                                double li = _builder.RowFormat.LeftIndent;
                                _builder.RowFormat.HeightRule = HeightRule.Exactly;
                                _builder.RowFormat.Height = 18.0;
                                _builder.RowFormat.LeftIndent = 0;

                                // 欄位名稱
                                Cell tc1 = _builder.InsertCell();
                                tc1.CellFormat.Width = miniUnitWitdh * 2;
                                tc1.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                _builder.Write("日期");

                                Cell tc2 = _builder.InsertCell();
                                tc2.CellFormat.Width = miniUnitWitdh;
                                tc2.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                _builder.Write("類別");

                                Cell tc3 = _builder.InsertCell();
                                tc3.CellFormat.Width = miniUnitWitdh * 5;
                                tc3.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                _builder.Write("事由");


                                // 異動資料內容
                                if (Global._ABCard_StudentUpdateRecDict[studID].Count > 0)
                                {
                                    foreach (AB_StudDateString ass in Global._ABCard_StudentUpdateRecDict[studID])
                                    {

                                        // 日期
                                        Cell c1 = _builder.InsertCell();
                                        c1.CellFormat.Width = miniUnitWitdh * 2;
                                        c1.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                        _builder.Write(ass.Date.ToShortDateString());

                                        // 類別
                                        Cell c2 = _builder.InsertCell();
                                        c2.CellFormat.Width = miniUnitWitdh;
                                        c2.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                        _builder.Write(ass.Text1);

                                        // 事由
                                        Cell c3 = _builder.InsertCell();
                                        c3.CellFormat.Width = miniUnitWitdh * 5;
                                        c3.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                        _builder.Write(ass.Text2);

                                        _builder.EndRow();
                                    }

                                    _builder.EndTable();

                                    //去除表格四邊的線
                                    foreach (Cell c in table.FirstRow.Cells)
                                        c.CellFormat.Borders.Top.LineStyle = LineStyle.None;

                                    foreach (Cell c in table.LastRow.Cells)
                                        c.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;

                                    foreach (Row r in table.Rows)
                                    {
                                        r.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                        r.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                    }
                                }
                                _builder.RowFormat.LeftIndent = li;
                            }
                        }
                    }


                    #endregion

                    #region 獎懲資料
                    if (e.FieldName == "獎懲資料" && e.FieldValue.ToString() != "")
                    {
                        string studID = e.FieldValue.ToString();
                        if (Global._ABCard_StudentMDRecordDict.ContainsKey(studID))
                        {
                            if (_builder.MoveToMergeField(e.FieldName))
                            {
                                Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                                cell.CellFormat.LeftPadding = 0;
                                cell.CellFormat.RightPadding = 0;
                                double width = cell.CellFormat.Width;
                                int columnCount = 8;
                                double miniUnitWitdh = width / (double)columnCount;

                                Table table = _builder.StartTable();
                                double li = _builder.RowFormat.LeftIndent;
                                _builder.RowFormat.HeightRule = HeightRule.Exactly;
                                _builder.RowFormat.Height = 18.0;
                                _builder.RowFormat.LeftIndent = 0;

                                // 欄位名稱
                                Cell tc1 = _builder.InsertCell();
                                tc1.CellFormat.Width = miniUnitWitdh * 2;
                                tc1.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                _builder.Write("日期");

                                Cell tc2 = _builder.InsertCell();
                                tc2.CellFormat.Width = miniUnitWitdh;
                                tc2.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                _builder.Write("類別");

                                Cell tc3 = _builder.InsertCell();
                                tc3.CellFormat.Width = miniUnitWitdh * 5;
                                tc3.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                _builder.Write("事由");


                                // 學習歷程內容
                                if (Global._ABCard_StudentMDRecordDict[studID].Count > 0)
                                {
                                    foreach (AB_StudDateString ass in Global._ABCard_StudentMDRecordDict[studID])
                                    {
                                        // 日期
                                        Cell c1 = _builder.InsertCell();
                                        c1.CellFormat.Width = miniUnitWitdh * 2;
                                        c1.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                        _builder.Write(ass.Date.ToShortDateString());

                                        // 類別
                                        Cell c2 = _builder.InsertCell();
                                        c2.CellFormat.Width = miniUnitWitdh;
                                        c2.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                        _builder.Write(ass.Text1);

                                        // 事由
                                        Cell c3 = _builder.InsertCell();
                                        c3.CellFormat.Width = miniUnitWitdh * 5;
                                        c3.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                        _builder.Write(ass.Text2);

                                        _builder.EndRow();
                                    }

                                    _builder.EndTable();

                                    //去除表格四邊的線
                                    foreach (Cell c in table.FirstRow.Cells)
                                        c.CellFormat.Borders.Top.LineStyle = LineStyle.None;

                                    foreach (Cell c in table.LastRow.Cells)
                                        c.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;

                                    foreach (Row r in table.Rows)
                                    {
                                        r.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                        r.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                    }
                                }
                                _builder.RowFormat.LeftIndent = li;
                            }
                        }
                    }


                    #endregion

                    #region 測驗資料處理
                    if (e.FieldName == "測驗紀錄" && e.FieldValue.ToString() != "")
                    {
                        string studID = e.FieldValue.ToString();
                        if (StudentInfoTransfer.StudQuizDataDefDict.ContainsKey(studID))
                        {
                            if (_builder.MoveToMergeField(e.FieldName))
                            {
                                Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                                cell.CellFormat.LeftPadding = 0;
                                cell.CellFormat.RightPadding = 0;
                                double width = cell.CellFormat.Width;
                                int columnCount = 10;
                                double miniUnitWitdh = width / (double)columnCount;

                                Table table = _builder.StartTable();
                                double li = _builder.RowFormat.LeftIndent;
                                _builder.RowFormat.HeightRule = HeightRule.Exactly;
                                _builder.RowFormat.Height = 18.0;
                                _builder.RowFormat.LeftIndent = 0;

                                // 欄位名稱
                                Cell tc1 = _builder.InsertCell();
                                tc1.CellFormat.Width = miniUnitWitdh * 2;
                                tc1.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                _builder.Write("測驗名稱");

                                Cell tc2 = _builder.InsertCell();
                                tc2.CellFormat.Width = miniUnitWitdh;
                                tc2.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                _builder.Write("測驗日期");

                                Cell tc3 = _builder.InsertCell();
                                tc3.CellFormat.Width = miniUnitWitdh * 7;
                                tc3.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                _builder.Write("測驗結果");

                                _builder.EndRow();


                                Dictionary<string, UDT_StudQuizDataDef> DataDict = new Dictionary<string, UDT_StudQuizDataDef>();

                                DataDict = StudentInfoTransfer.StudQuizDataDefDict[studID];



                                // 答案
                                if (DataDict.Count > 0)
                                {
                                    List<string> tmp = new List<string>();
                                    foreach (KeyValuePair<string, UDT_StudQuizDataDef> da in DataDict)
                                    {
                                        Cell c1 = _builder.InsertCell();
                                        c1.CellFormat.Width = miniUnitWitdh * 2;
                                        c1.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                        _builder.Write(da.Key);


                                        Cell c2 = _builder.InsertCell();
                                        c2.CellFormat.Width = miniUnitWitdh;
                                        c2.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                        if (da.Value.ImplementationDate.HasValue)
                                            _builder.Write(da.Value.ImplementationDate.Value.ToShortDateString());
                                        else
                                            _builder.Write("");


                                        Cell c3 = _builder.InsertCell();
                                        c3.CellFormat.Width = miniUnitWitdh * 7;
                                        c3.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                        if (string.IsNullOrEmpty(da.Value.Content))
                                            _builder.Write("");
                                        else
                                        {
                                            StringBuilder sb = new StringBuilder();
                                            sb.Append("<root>"); sb.Append(da.Value.Content); sb.Append("</root>");
                                            XElement elmC = XElement.Parse(sb.ToString());

                                            tmp.Clear();
                                            foreach (XElement el in elmC.Elements("Item"))
                                                tmp.Add(el.Attribute("name").Value + ":" + el.Attribute("value").Value);

                                            _builder.Write(string.Join(",", tmp.ToArray()));
                                        }
                                        _builder.EndRow();
                                    }



                                    _builder.EndTable();

                                    //去除表格四邊的線
                                    foreach (Cell c in table.FirstRow.Cells)
                                        c.CellFormat.Borders.Top.LineStyle = LineStyle.None;

                                    foreach (Cell c in table.LastRow.Cells)
                                        c.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;

                                    foreach (Row r in table.Rows)
                                    {
                                        r.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                        r.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                    }
                                }
                                _builder.RowFormat.LeftIndent = li;
                            }
                        }
                    }
                    #endregion

                    #region 參加社團
                    // 參加社團
                    if (e.FieldName == "參加社團" && e.FieldValue.ToString() != "")
                    {
                        string studID = e.FieldValue.ToString();
                        if (Global._ABCard_StudentSpecCourseDict.ContainsKey(studID))
                        {
                            if (_builder.MoveToMergeField(e.FieldName))
                            {
                                Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                                cell.CellFormat.LeftPadding = 0;
                                cell.CellFormat.RightPadding = 0;
                                double width = cell.CellFormat.Width;
                                int columnCount = 4;
                                double miniUnitWitdh = width / (double)columnCount;

                                Table table = _builder.StartTable();
                                double li = _builder.RowFormat.LeftIndent;
                                _builder.RowFormat.HeightRule = HeightRule.Exactly;
                                _builder.RowFormat.Height = 18.0;
                                _builder.RowFormat.LeftIndent = 0;


                                // 參加社團內容
                                foreach (AB_StudText ast in Global._ABCard_StudentSpecCourseDict[studID])
                                {

                                    // 年級學期
                                    Cell c1 = _builder.InsertCell();
                                    c1.CellFormat.Width = miniUnitWitdh;
                                    c1.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                    _builder.Write(Utility.ConvertGradeYearSemester(ast.GradeYear, ast.Semester));

                                    // 內容
                                    Cell c2 = _builder.InsertCell();
                                    c2.CellFormat.Width = miniUnitWitdh * 3;
                                    c2.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                    _builder.Write(ast.Text1);

                                    _builder.EndRow();
                                }

                                _builder.EndTable();

                                //去除表格四邊的線
                                foreach (Cell c in table.FirstRow.Cells)
                                    c.CellFormat.Borders.Top.LineStyle = LineStyle.None;

                                foreach (Cell c in table.LastRow.Cells)
                                    c.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;

                                foreach (Row r in table.Rows)
                                {
                                    r.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                    r.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                }

                                _builder.RowFormat.LeftIndent = li;
                            }
                        }
                    }
                    #endregion

                    #region 擔任幹部
                    // 擔任幹部
                    if (e.FieldName == "擔任幹部" && e.FieldValue.ToString() != "")
                    {
                        string studID = e.FieldValue.ToString();
                        if (Global._ABCard_StudentTheCadreDict.ContainsKey(studID))
                        {
                            if (_builder.MoveToMergeField(e.FieldName))
                            {
                                Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                                cell.CellFormat.LeftPadding = 0;
                                cell.CellFormat.RightPadding = 0;
                                double width = cell.CellFormat.Width;
                                int columnCount = 4;
                                double miniUnitWitdh = width / (double)columnCount;

                                Table table = _builder.StartTable();
                                double li = _builder.RowFormat.LeftIndent;
                                _builder.RowFormat.HeightRule = HeightRule.Exactly;
                                _builder.RowFormat.Height = 18.0;
                                _builder.RowFormat.LeftIndent = 0;

                                // 擔任幹部內容
                                foreach (AB_StudText ast in Global._ABCard_StudentTheCadreDict[studID])
                                {

                                    // 年級學期
                                    Cell c1 = _builder.InsertCell();
                                    c1.CellFormat.Width = miniUnitWitdh;
                                    c1.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                    _builder.Write(Utility.ConvertGradeYearSemester(ast.GradeYear, ast.Semester));

                                    // 內容
                                    Cell c2 = _builder.InsertCell();
                                    c2.CellFormat.Width = miniUnitWitdh * 3;
                                    c2.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                    _builder.Write(ast.Text1);

                                    _builder.EndRow();
                                }

                                _builder.EndTable();

                                //去除表格四邊的線
                                foreach (Cell c in table.FirstRow.Cells)
                                    c.CellFormat.Borders.Top.LineStyle = LineStyle.None;

                                foreach (Cell c in table.LastRow.Cells)
                                    c.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;

                                foreach (Row r in table.Rows)
                                {
                                    r.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                    r.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                }

                                _builder.RowFormat.LeftIndent = li;
                            }
                        }
                    }
                    #endregion

                    #region 學習領域與畢業成績
                    // 學習領域與畢業成績
                    if (e.FieldName == "學習領域與畢業成績" && e.FieldValue.ToString() != "")
                    {
                        string studID = e.FieldValue.ToString();

                        if (Global._AB_StudSemsDomainScoreDict.ContainsKey(studID))
                        {
                            // 取得該位學生資料
                            AB_StudSemsDomainScore studData = Global._AB_StudSemsDomainScoreDict[studID];

                            if (_builder.MoveToMergeField(e.FieldName))
                            {
                                Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                                cell.CellFormat.LeftPadding = 0;
                                cell.CellFormat.RightPadding = 0;
                                double width = cell.CellFormat.Width;
                                // 學習領域+動態名稱+畢業成績
                                int columnCount = studData.GradeYearList.Count + 2;
                                double miniUnitWitdh = width / (double)columnCount;

                                Table table = _builder.StartTable();
                                double li = _builder.RowFormat.LeftIndent;
                                _builder.RowFormat.HeightRule = HeightRule.Exactly;
                                _builder.RowFormat.Height = 18.0;
                                _builder.RowFormat.LeftIndent = 0;

                                // 處理報表欄位名稱
                                // 學習領域
                                _builder.InsertCell();
                                _builder.CellFormat.Width = miniUnitWitdh;

                                _builder.Write("學習領域");

                                int tmpSchoolYear = 0;
                                // 處理學年度
                                foreach (DAO.AB_RptColIdx cidx in studData.GradeYearList)
                                {
                                    if (tmpSchoolYear != cidx.SchoolYear)
                                    {
                                        _builder.InsertCell().Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                        _builder.Write(cidx.SchoolYear + "學年度");
                                    }
                                    else
                                        _builder.InsertCell();

                                    tmpSchoolYear = cidx.SchoolYear;
                                }

                                // 畢業成績
                                _builder.InsertCell();
                                _builder.CellFormat.Width = miniUnitWitdh;
                                _builder.Write("畢業成績");
                                _builder.EndRow();

                                //處理第2列學期
                                // 學習領域不填合併用
                                _builder.InsertCell();

                                // 處理學期
                                foreach (DAO.AB_RptColIdx cidx in studData.GradeYearList)
                                {
                                    if (cidx.Semester == 1)
                                    {
                                        _builder.InsertCell().Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                        _builder.Write("上學期");
                                    }

                                    if (cidx.Semester == 2)
                                    {
                                        _builder.InsertCell().Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                        _builder.Write("下學期");
                                    }
                                }

                                // 畢業成績不填合併用
                                _builder.InsertCell();
                                _builder.CellFormat.Width = miniUnitWitdh;
                                _builder.EndRow();

                                int domainRowidx = 1, LangMgBeginRow = 2, LangMgEndRow = 0;
                                // 檢查是否語文合併
                                bool checkMgLang = false;
                                // 放入成績
                                // 處理有語文類需要展開
                                foreach (KeyValuePair<string, List<string>> domainName in studData.DomainNameDict)
                                {
                                    // 沒有科目跳出不處理
                                    if (domainName.Value.Count == 0)
                                        continue;

                                    checkMgLang = true;

                                    domainRowidx += 1;


                                    LangMgEndRow = LangMgBeginRow;
                                    // 領域成績使用科目
                                    foreach (string SubjName in domainName.Value)
                                    {
                                        // 領域名稱
                                        _builder.CellFormat.Width = miniUnitWitdh / 2;
                                        _builder.InsertCell();
                                        _builder.Write(domainName.Key);

                                        // 科目名稱
                                        _builder.InsertCell();
                                        _builder.Write(SubjName);

                                        // 讀取並放入成績
                                        foreach (DAO.AB_RptColIdx cidx in studData.GradeYearList)
                                        {
                                            // 判斷是否有資料
                                            bool hasData = false;

                                            foreach (DAO.AB_SemsDomainScore semsScore in (from data in studData.SemsDomainScoreList where data.DomainName == domainName.Key && data.SchoolYear == cidx.SchoolYear && data.Semester == cidx.Semester && data.LangSubjDict.ContainsKey(SubjName) select data))
                                            {
                                                _builder.InsertCell();
                                                _builder.CellFormat.Width = miniUnitWitdh;
                                                if (semsScore.LangSubjDict[SubjName].HasValue)
                                                    _builder.Write(semsScore.LangSubjDict[SubjName].Value.ToString());

                                                hasData = true;
                                            }

                                            if (hasData == false)
                                                _builder.InsertCell();
                                        }

                                        // 畢業
                                        _builder.InsertCell();
                                        //         _builder.Write("test-" + LangMgEndRow);
                                        _builder.EndRow();
                                        LangMgEndRow++;
                                    }
                                }


                                // 領域不需要動態展開
                                foreach (KeyValuePair<string, List<string>> domainName in studData.DomainNameDict)
                                {
                                    domainRowidx += 1;
                                    // 領域名稱
                                    _builder.CellFormat.Width = miniUnitWitdh;
                                    _builder.InsertCell();
                                    _builder.Write(domainName.Key);


                                    // 領域成績
                                    foreach (DAO.AB_RptColIdx cidx in studData.GradeYearList)
                                    {
                                        // 判斷是否有資料
                                        bool hasData = false;

                                        foreach (DAO.AB_SemsDomainScore semsScore in (from data in studData.SemsDomainScoreList where data.DomainName == domainName.Key && data.SchoolYear == cidx.SchoolYear && data.Semester == cidx.Semester select data))
                                        {
                                            _builder.InsertCell();
                                            if (semsScore.DomainScore.HasValue)
                                                _builder.Write(semsScore.DomainScore.Value.ToString());

                                            hasData = true;
                                        }

                                        if (hasData == false)
                                            _builder.InsertCell();
                                    }

                                    // 領域畢業成績
                                    if (studData.GraduateScoreDict.ContainsKey(domainName.Key))
                                    {
                                        _builder.InsertCell();
                                        if (studData.GraduateScoreDict[domainName.Key].HasValue)
                                            _builder.Write(studData.GraduateScoreDict[domainName.Key].Value.ToString());

                                    }
                                    else
                                    {
                                        _builder.InsertCell();
                                        _builder.Write("");
                                    }
                                    _builder.EndRow();
                                }

                                _builder.EndTable();

                                // 處理學期領域合併
                                Cell top = table.Rows[0].Cells[0];
                                Cell bottom = table.Rows[1].Cells[0];
                                Utility.VerticallyMergeCells(top, bottom);
                                top.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;

                                // 處理學年度合併
                                Dictionary<int, bool> tmpMegDict = new Dictionary<int, bool>();
                                tmpSchoolYear = 0;
                                foreach (DAO.AB_RptColIdx cidx in studData.GradeYearList)
                                {
                                    if (tmpMegDict.ContainsKey(cidx.SchoolYear))
                                        tmpMegDict[cidx.SchoolYear] = true;
                                    else
                                        tmpMegDict.Add(cidx.SchoolYear, false);
                                }

                                int MgLeftIdx = 1;

                                foreach (KeyValuePair<int, bool> mg in tmpMegDict)
                                {
                                    // 需要合併
                                    if (mg.Value)
                                    {
                                        Cell left = table.Rows[0].Cells[MgLeftIdx];
                                        Cell right = table.Rows[0].Cells[MgLeftIdx + 1];
                                        Utility.HorizontallyMergeCells(left, right);
                                        left.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                        MgLeftIdx += 2;
                                    }
                                    else
                                        MgLeftIdx += 1;
                                }

                                // 處理畢業領域合併
                                int tmpCellsCount = table.Rows[0].Cells.Count - 1;
                                top = table.Rows[0].Cells[tmpCellsCount];
                                bottom = table.Rows[1].Cells[tmpCellsCount];
                                Utility.VerticallyMergeCells(top, bottom);
                                top.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;

                                // 處理語文要合併，因為合併列印一次只能1格，所以需要使用迴圈處理。
                                if (checkMgLang)
                                {
                                    for (int i = 1; i < (LangMgEndRow - LangMgBeginRow); i++)
                                    {
                                        top = table.Rows[LangMgBeginRow].Cells[0];
                                        bottom = table.Rows[LangMgBeginRow + i].Cells[0];
                                        Utility.VerticallyMergeCells(top, bottom);
                                    }
                                    top.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;

                                    tmpCellsCount = table.Rows[LangMgBeginRow].Cells.Count - 1;

                                    for (int i = 1; i < (LangMgEndRow - LangMgBeginRow); i++)
                                    {
                                        top = table.Rows[LangMgBeginRow].Cells[tmpCellsCount];
                                        bottom = table.Rows[LangMgBeginRow + i].Cells[tmpCellsCount];
                                        Utility.VerticallyMergeCells(top, bottom);
                                    }
                                    top.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                }

                                //去除表格四邊的線
                                foreach (Cell c in table.FirstRow.Cells)
                                    c.CellFormat.Borders.Top.LineStyle = LineStyle.None;

                                foreach (Cell c in table.LastRow.Cells)
                                    c.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;

                                foreach (Row r in table.Rows)
                                {
                                    r.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                    r.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                }

                                _builder.RowFormat.LeftIndent = li;
                            }
                        }
                    }
                    #endregion

                    #region 輔導個案會議
                    if (e.FieldName == "輔導個案會議" && e.FieldValue.ToString() != "")
                    {
                        string studID = e.FieldValue.ToString();

                        if (Global._AB_CaseMeetingRecordToABRptDataDict.ContainsKey(studID))
                        {
                            bool MergeOtherData = false;
                            // 其它內容
                            string FieldName = e.FieldName + "其它內容";
                            if (_builder.MoveToMergeField(FieldName))
                                MergeOtherData = true;

                            // 資料
                            if (_builder.MoveToMergeField(e.FieldName))
                            {
                                Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                                cell.CellFormat.LeftPadding = 0;
                                cell.CellFormat.RightPadding = 0;
                                double width = cell.CellFormat.Width;

                                int columnCount = 5;
                                double miniUnitWitdh = width / (double)columnCount;

                                Table table = _builder.StartTable();
                                double li = _builder.RowFormat.LeftIndent;
                                _builder.RowFormat.HeightRule = HeightRule.Exactly;
                                _builder.RowFormat.Height = 18.0;
                                _builder.RowFormat.LeftIndent = 0;

                                // 處理報表欄位名稱
                                // 標題
                                _builder.InsertCell().Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                _builder.CellFormat.Width = miniUnitWitdh;
                                _builder.Write("會議日期");
                                _builder.InsertCell();
                                _builder.CellFormat.Width = miniUnitWitdh * 3;
                                _builder.Write("會議事由");
                                _builder.InsertCell();
                                _builder.CellFormat.Width = miniUnitWitdh;
                                _builder.Write("記錄者帳號");
                                _builder.EndRow();

                                foreach (AB_RptCounselData csData in Global._AB_CaseMeetingRecordToABRptDataDict[studID])
                                {
                                    // 內容
                                    _builder.InsertCell().CellFormat.Width = miniUnitWitdh;
                                    if (csData.DataDict.ContainsKey("會議日期"))
                                        _builder.Write(csData.DataDict["會議日期"]);

                                    _builder.InsertCell().CellFormat.Width = miniUnitWitdh * 3;
                                    if (csData.DataDict.ContainsKey("會議事由"))
                                        _builder.Write(csData.DataDict["會議事由"]);

                                    _builder.InsertCell().CellFormat.Width = miniUnitWitdh;
                                    if (csData.DataDict.ContainsKey("記錄者帳號"))
                                        _builder.Write(csData.DataDict["記錄者帳號"]);

                                    _builder.EndRow();

                                    // 其它內容
                                    if (MergeOtherData)
                                    {
                                        _builder.InsertCell().CellFormat.Width = miniUnitWitdh * 5;
                                        _builder.Write(string.Join(",", csData.OtherDataList.ToArray()));
                                        _builder.EndRow();
                                    }
                                }
                                _builder.EndTable();

                                //去除表格四邊的線
                                foreach (Cell c in table.FirstRow.Cells)
                                    c.CellFormat.Borders.Top.LineStyle = LineStyle.None;

                                foreach (Cell c in table.LastRow.Cells)
                                    c.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;

                                foreach (Row r in table.Rows)
                                {
                                    r.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                    r.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                }

                                _builder.RowFormat.LeftIndent = li;
                            }
                        }
                    }


                    #endregion

                    #region 輔導優先關懷
                    if (e.FieldName == "輔導優先關懷" && e.FieldValue.ToString() != "")
                    {
                        string studID = e.FieldValue.ToString();

                        if (Global._AB_CareRecordToABRptDataDict.ContainsKey(studID))
                        {
                            bool MergeOtherData = false;
                            // 其它內容
                            string FieldName = e.FieldName + "其它內容";
                            if (_builder.MoveToMergeField(FieldName))
                                MergeOtherData = true;

                            // 主要
                            if (_builder.MoveToMergeField(e.FieldName))
                            {
                                Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                                cell.CellFormat.LeftPadding = 0;
                                cell.CellFormat.RightPadding = 0;
                                double width = cell.CellFormat.Width;

                                int columnCount = 5;
                                double miniUnitWitdh = width / (double)columnCount;

                                Table table = _builder.StartTable();
                                double li = _builder.RowFormat.LeftIndent;
                                _builder.RowFormat.HeightRule = HeightRule.Exactly;
                                _builder.RowFormat.Height = 18.0;
                                _builder.RowFormat.LeftIndent = 0;

                                // 標題
                                _builder.InsertCell().CellFormat.Width = miniUnitWitdh;
                                _builder.Write("立案日期");
                                _builder.InsertCell().CellFormat.Width = miniUnitWitdh * 3;
                                _builder.Write("個案類別");
                                _builder.InsertCell().CellFormat.Width = miniUnitWitdh;
                                _builder.Write("記錄者帳號");
                                _builder.EndRow();

                                foreach (AB_RptCounselData csData in Global._AB_CareRecordToABRptDataDict[studID])
                                {
                                    // 內容
                                    _builder.InsertCell().CellFormat.Width = miniUnitWitdh;
                                    if (csData.DataDict.ContainsKey("立案日期"))
                                        _builder.Write(csData.DataDict["立案日期"]);

                                    _builder.InsertCell().CellFormat.Width = miniUnitWitdh * 3;
                                    if (csData.DataDict.ContainsKey("個案類別"))
                                        _builder.Write(csData.DataDict["個案類別"]);

                                    _builder.InsertCell().CellFormat.Width = miniUnitWitdh;
                                    if (csData.DataDict.ContainsKey("記錄者帳號"))
                                        _builder.Write(csData.DataDict["記錄者帳號"]);

                                    _builder.EndRow();

                                    // 其它內容
                                    if (MergeOtherData)
                                    {
                                        _builder.InsertCell().CellFormat.Width = miniUnitWitdh * 5;
                                        _builder.Write(string.Join(",", csData.OtherDataList.ToArray()));
                                        _builder.EndRow();
                                    }
                                }
                                _builder.EndTable();


                                //去除表格四邊的線
                                foreach (Cell c in table.FirstRow.Cells)
                                    c.CellFormat.Borders.Top.LineStyle = LineStyle.None;

                                foreach (Cell c in table.LastRow.Cells)
                                    c.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;

                                foreach (Row r in table.Rows)
                                {
                                    r.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                    r.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                }

                                _builder.RowFormat.LeftIndent = li;
                            }
                        }
                    }

                    #endregion

                    #region 輔導晤談紀錄
                    if (e.FieldName == "輔導晤談紀錄" && e.FieldValue.ToString() != "")
                    {
                        string studID = e.FieldValue.ToString();

                        if (Global._AB_InterViewDataToABRptDataDict.ContainsKey(studID))
                        {
                            bool MergeOtherData = false;
                            // 其它內容
                            string FieldName = e.FieldName + "其它內容";
                            if (_builder.MoveToMergeField(FieldName))
                                MergeOtherData = true;

                            // 主要
                            if (_builder.MoveToMergeField(e.FieldName))
                            {
                                Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                                cell.CellFormat.LeftPadding = 0;
                                cell.CellFormat.RightPadding = 0;
                                double width = cell.CellFormat.Width;

                                int columnCount = 5;
                                double miniUnitWitdh = width / (double)columnCount;

                                Table table = _builder.StartTable();
                                double li = _builder.RowFormat.LeftIndent;
                                _builder.RowFormat.HeightRule = HeightRule.Exactly;
                                _builder.RowFormat.Height = 18.0;
                                _builder.RowFormat.LeftIndent = 0;

                                foreach (AB_RptCounselData csData in Global._AB_InterViewDataToABRptDataDict[studID])
                                {
                                    _builder.InsertCell().CellFormat.Width = miniUnitWitdh;
                                    if (csData.DataDict.ContainsKey("晤談日期"))
                                        _builder.Write(csData.DataDict["晤談日期"]);

                                    _builder.InsertCell().CellFormat.Width = miniUnitWitdh;
                                    if (csData.DataDict.ContainsKey("晤談對象"))
                                        _builder.Write(csData.DataDict["晤談對象"]);

                                    _builder.InsertCell().CellFormat.Width = miniUnitWitdh * 2;
                                    if (csData.DataDict.ContainsKey("晤談事由"))
                                        _builder.Write(csData.DataDict["晤談事由"]);


                                    _builder.InsertCell().CellFormat.Width = miniUnitWitdh;
                                    if (csData.DataDict.ContainsKey("記錄者帳號"))
                                        _builder.Write(csData.DataDict["記錄者帳號"]);

                                    _builder.EndRow();

                                    if (MergeOtherData)
                                    {
                                        _builder.InsertCell().CellFormat.Width = miniUnitWitdh * 5;
                                        _builder.Write(string.Join(",", csData.OtherDataList.ToArray()));
                                    }
                                }

                                //去除表格四邊的線
                                foreach (Cell c in table.FirstRow.Cells)
                                    c.CellFormat.Borders.Top.LineStyle = LineStyle.None;

                                foreach (Cell c in table.LastRow.Cells)
                                    c.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;

                                foreach (Row r in table.Rows)
                                {
                                    r.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                    r.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                }

                                _builder.RowFormat.LeftIndent = li;
                            }
                        }
                    }
                    #endregion

                    #region 學生照片
                    if ((e.FieldName == "入學照片" || e.FieldName == "畢業照片") && e.FieldValue.ToString() != "")
                    {
                        string studID = e.FieldValue.ToString();

                        // 入學照片
                        if (Global._AB_StudentFreshmanDict.ContainsKey(studID))
                        {
                            if (_builder.MoveToMergeField("入學照片"))
                            {
                                byte[] photo = Convert.FromBase64String(Global._AB_StudentFreshmanDict[studID]);
                                if (photo != null)
                                {
                                    Shape photoShape = new Shape(e.Document, ShapeType.Image);
                                    photoShape.ImageData.SetImage(photo);
                                    photoShape.WrapType = WrapType.Inline;

                                    Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                                    cell.CellFormat.LeftPadding = 0;
                                    cell.CellFormat.RightPadding = 0;
                                    photoShape.Height = (cell.ParentNode as Row).RowFormat.Height;
                                    photoShape.Width = cell.CellFormat.Width;

                                    _builder.InsertNode(photoShape);
                                }
                            }
                        }

                        // 畢業照片
                        if (Global._AB_StudentGraduateDict.ContainsKey(studID))
                        {
                            if (_builder.MoveToMergeField("畢業照片"))
                            {
                                byte[] photo = Convert.FromBase64String(Global._AB_StudentGraduateDict[studID]);
                                if (photo != null)
                                {
                                    Shape photoShape = new Shape(e.Document, ShapeType.Image);
                                    photoShape.ImageData.SetImage(photo);
                                    photoShape.WrapType = WrapType.Inline;

                                    Cell cell1 = _builder.CurrentParagraph.ParentNode as Cell;
                                    cell1.CellFormat.LeftPadding = 0;
                                    cell1.CellFormat.RightPadding = 0;
                                    photoShape.Height = (cell1.ParentNode as Row).RowFormat.Height;
                                    photoShape.Width = cell1.CellFormat.Width;
                                    _builder.InsertNode(photoShape);
                                }
                            }
                        }

                    }
                    #endregion


                }
                catch (Exception ex)
                {
                    // FISCA.Presentation.Controls.MsgBox.Show(ex.Message);
                }
            
            
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
              
            }
        }

        //void MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
        //{
          
        //}
    }
}
