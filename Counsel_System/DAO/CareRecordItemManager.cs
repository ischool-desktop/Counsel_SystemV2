using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 輔導優先關懷項目管理
    /// </summary>
    public class CareRecordItemManager
    {   
        public enum ItemType {個案類別,個案來源}

        public CareRecordItemManager()
        {
            // 檢查是否有預設資料
            CheckDefault();
        
        }

        /// <summary>
        /// 檢查預設是否存在
        /// </summary>
        private void CheckDefault()
        {
            string name1 = ItemType.個案類別.ToString();
            string name2 = ItemType.個案來源.ToString();
            UDT_SystemListDef def1 = UDTTransfer.GetSystemListByName(name1);
            UDT_SystemListDef def2 = UDTTransfer.GetSystemListByName(name2);
            List<UDT_SystemListDef> insert = new List<UDT_SystemListDef>();
            List<UDT_SystemListDef> update = new List<UDT_SystemListDef>();
            if (string.IsNullOrEmpty(def1.Content))
            {
                bool isInsert = false;
                if (def1.Content == null)
                    isInsert = true;

                def1 = new UDT_SystemListDef ();
                def1.Name = name1;
                def1.Content = GetDefaultItemXmlString(ItemType.個案類別);

                if (isInsert)
                    insert.Add(def1);
                else
                    update.Add(def1);
            }
            if (string.IsNullOrEmpty(def2.Content))
            {
                bool isInsert = false;
                if (def2.Content == null)
                    isInsert = true;
                
                def2 = new UDT_SystemListDef();
                def2.Name = name2;
                def2.Content = GetDefaultItemXmlString(ItemType.個案來源);
                if (isInsert)
                    insert.Add(def2);
                else
                    update.Add(def2);

            }

            if (insert.Count > 0)
                UDTTransfer.InsertSystemList(insert);

            if (update.Count > 0)
                UDTTransfer.UpdateSystemList(update);
        }

        /// <summary>
        /// 取得項目
        /// </summary>
        /// <param name="name"></param>
        /// <param name="itmType"></param>
        /// <returns></returns>
        public List<string> GetItemList(ItemType itmType)
        {
            List<string> retVal = new List<string>();

            UDT_SystemListDef data=UDTTransfer.GetSystemListByName(itmType.ToString());
            XElement elmData = XElement.Parse(data.Content);
            if (elmData != null)
                retVal = (from elm in elmData.Elements("item") select elm.Attribute("name").Value).ToList();
            return retVal;
        }

        /// <summary>
        /// 新增項目
        /// </summary>
        /// <param name="itemName"></param>
        /// <param name="itmType"></param>
        public void AddItem(string itemName, ItemType itmType)
        { 
            UDT_SystemListDef data=UDTTransfer.GetSystemListByName(itmType.ToString());
            XElement elmData = XElement.Parse(data.Content);
            List<UDT_SystemListDef> update= new List<UDT_SystemListDef>();
            if (elmData != null)
            {
                XElement elm = new XElement("item");
                elm.SetAttributeValue("name", itemName);
                elmData.Add(elm);
            }
            data.Content = elmData.ToString();
            update.Add(data);
            UDTTransfer.UpdateSystemList(update);
        }

        /// <summary>
        /// 取得預設資料
        /// </summary>
        /// <param name="itmType"></param>
        /// <returns></returns>
        private string GetDefaultItemXmlString(ItemType itmType)
        {
            XElement retVal = new XElement("items");
            List<string> itemList = new List<string>();
            if (itmType == ItemType.個案類別)
            {
                string strCaseCategoryItems = "學習適應困難個案,人際適應困難個案,嚴重缺曠個案,重大韋規個案,網路沉迷個案,家庭暴力個案,自傷及自殺個案,性侵害及性騷擾個案,涉及犯罪案件個案,精神疾患個案,家庭功能低落個案,情感困擾個案,其它";
                itemList = strCaseCategoryItems.Split(',').ToList();                
            }

            if (itmType == ItemType.個案來源)
            {
                string strCaseOriginItems = "導師轉介,任課老師轉介,家長轉介,個案主動求助,心理測驗篩檢,校外單位轉介,輔導人員發掘,其它";
                itemList = strCaseOriginItems.Split(',').ToList();
            }

            foreach (string item in itemList)
            {
                XElement elm = new XElement("item");
                elm.SetAttributeValue("name", item);
                retVal.Add(elm);
            }

            return retVal.ToString();
        }
    }
}
