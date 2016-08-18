using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System.DAO
{
    /// <summary>
    /// 輔導-綜合紀錄表-家庭狀況
    /// </summary>
    [TableName("ischool.counsel.family_status")]
    public class UDT_AB_FamilyStatus:ActiveRecord
    {
        ///<summary>
        /// 學生編號
        ///</summary>
        [Field(Field = "ref_student_id", Indexed = false)]
        public int StudentID { get; set; }

        ///<summary>
        /// 父母關係一年
        ///</summary>
        [Field(Field = "pr1", Indexed = false)]
        public string ParentRelationship1 { get; set; }

        ///<summary>
        /// 父母關係二年
        ///</summary>
        [Field(Field = "pr2", Indexed = false)]
        public string ParentRelationship2 { get; set; }

        ///<summary>
        /// 父母關係三年
        ///</summary>
        [Field(Field = "pr3", Indexed = false)]
        public string ParentRelationship3 { get; set; }

        ///<summary>
        /// 父母關係四年
        ///</summary>
        [Field(Field = "pr4", Indexed = false)]
        public string ParentRelationship4 { get; set; }

        ///<summary>
        /// 父母關係五年
        ///</summary>
        [Field(Field = "pr5", Indexed = false)]
        public string ParentRelationship5 { get; set; }

        ///<summary>
        /// 父母關係六年
        ///</summary>
        [Field(Field = "pr6", Indexed = false)]
        public string ParentRelationship6 { get; set; }

        ///<summary>
        /// 家庭氣氛一年
        ///</summary>
        [Field(Field = "fa1", Indexed = false)]
        public string FamilyAtmosphere1 { get; set; }

        ///<summary>
        /// 家庭氣氛二年
        ///</summary>
        [Field(Field = "fa2", Indexed = false)]
        public string FamilyAtmosphere2 { get; set; }

        ///<summary>
        /// 家庭氣氛三年
        ///</summary>
        [Field(Field = "fa3", Indexed = false)]
        public string FamilyAtmosphere3 { get; set; }

        ///<summary>
        /// 家庭氣氛四年
        ///</summary>
        [Field(Field = "fa4", Indexed = false)]
        public string FamilyAtmosphere4 { get; set; }

        ///<summary>
        /// 家庭氣氛五年
        ///</summary>
        [Field(Field = "fa5", Indexed = false)]
        public string FamilyAtmosphere5 { get; set; }

        ///<summary>
        /// 家庭氣氛六年
        ///</summary>
        [Field(Field = "fa6", Indexed = false)]
        public string FamilyAtmosphere6 { get; set; }

        ///<summary>
        /// 父親管教一年
        ///</summary>
        [Field(Field = "fd1", Indexed = false)]
        public string FatherDiscipline1 { get; set; }

        ///<summary>
        /// 父親管教二年
        ///</summary>
        [Field(Field = "fd2", Indexed = false)]
        public string FatherDiscipline2 { get; set; }

        ///<summary>
        /// 父親管教三年
        ///</summary>
        [Field(Field = "fd3", Indexed = false)]
        public string FatherDiscipline3 { get; set; }

        ///<summary>
        /// 父親管教四年
        ///</summary>
        [Field(Field = "fd4", Indexed = false)]
        public string FatherDiscipline4 { get; set; }

        ///<summary>
        /// 父親管教五年
        ///</summary>
        [Field(Field = "fd5", Indexed = false)]
        public string FatherDiscipline5 { get; set; }

        ///<summary>
        /// 父親管教六年
        ///</summary>
        [Field(Field = "fd6", Indexed = false)]
        public string FatherDiscipline6 { get; set; }

        ///<summary>
        /// 母親管教一年
        ///</summary>
        [Field(Field = "md1", Indexed = false)]
        public string MotherDiscipline1 { get; set; }

        ///<summary>
        /// 母親管教二年
        ///</summary>
        [Field(Field = "md2", Indexed = false)]
        public string MotherDiscipline2 { get; set; }

        ///<summary>
        /// 母親管教三年
        ///</summary>
        [Field(Field = "md3", Indexed = false)]
        public string MotherDiscipline3 { get; set; }

        ///<summary>
        /// 母親管教四年
        ///</summary>
        [Field(Field = "md4", Indexed = false)]
        public string MotherDiscipline4 { get; set; }

        ///<summary>
        /// 母親管教五年
        ///</summary>
        [Field(Field = "md5", Indexed = false)]
        public string MotherDiscipline5 { get; set; }

        ///<summary>
        /// 母親管教六年
        ///</summary>
        [Field(Field = "md6", Indexed = false)]
        public string MotherDiscipline6 { get; set; }

        ///<summary>
        /// 居住環境一年
        ///</summary>
        [Field(Field = "le1", Indexed = false)]
        public string LivingEnvironment1 { get; set; }

        ///<summary>
        /// 居住環境二年
        ///</summary>
        [Field(Field = "le2", Indexed = false)]
        public string LivingEnvironment2 { get; set; }

        ///<summary>
        /// 居住環境三年
        ///</summary>
        [Field(Field = "le3", Indexed = false)]
        public string LivingEnvironment3 { get; set; }

        ///<summary>
        /// 居住環境四年
        ///</summary>
        [Field(Field = "le4", Indexed = false)]
        public string LivingEnvironment4 { get; set; }

        ///<summary>
        /// 居住環境五年
        ///</summary>
        [Field(Field = "le5", Indexed = false)]
        public string LivingEnvironment5 { get; set; }

        ///<summary>
        /// 居住環境六年
        ///</summary>
        [Field(Field = "le6", Indexed = false)]
        public string LivingEnvironment6 { get; set; }

        ///<summary>
        /// 本人住宿一年
        ///</summary>
        [Field(Field = "stay1", Indexed = false)]
        public string Stay1 { get; set; }

        ///<summary>
        /// 本人住宿二年
        ///</summary>
        [Field(Field = "stay2", Indexed = false)]
        public string Stay2 { get; set; }

        ///<summary>
        /// 本人住宿三年
        ///</summary>
        [Field(Field = "stay3", Indexed = false)]
        public string Stay3 { get; set; }

        ///<summary>
        /// 本人住宿四年
        ///</summary>
        [Field(Field = "stay4", Indexed = false)]
        public string Stay4 { get; set; }

        ///<summary>
        /// 本人住宿五年
        ///</summary>
        [Field(Field = "stay5", Indexed = false)]
        public string Stay5 { get; set; }

        ///<summary>
        /// 本人住宿六年
        ///</summary>
        [Field(Field = "stay6", Indexed = false)]
        public string Stay6 { get; set; }

        ///<summary>
        /// 家庭經濟狀況一年
        ///</summary>
        [Field(Field = "fe1", Indexed = false)]
        public string FamilyEeconomic1 { get; set; }

        ///<summary>
        /// 家庭經濟狀況二年
        ///</summary>
        [Field(Field = "fe2", Indexed = false)]
        public string FamilyEeconomic2 { get; set; }

        ///<summary>
        /// 家庭經濟狀況三年
        ///</summary>
        [Field(Field = "fe3", Indexed = false)]
        public string FamilyEeconomic3 { get; set; }

        ///<summary>
        /// 家庭經濟狀況四年
        ///</summary>
        [Field(Field = "fe4", Indexed = false)]
        public string FamilyEeconomic4 { get; set; }

        ///<summary>
        /// 家庭經濟狀況五年
        ///</summary>
        [Field(Field = "fe5", Indexed = false)]
        public string FamilyEeconomic5 { get; set; }

        ///<summary>
        /// 家庭經濟狀況六年
        ///</summary>
        [Field(Field = "fe6", Indexed = false)]
        public string FamilyEeconomic6 { get; set; }

        ///<summary>
        /// 每星期零用錢一年
        ///</summary>
        [Field(Field = "wpm1", Indexed = false)]
        public int WeekPocketMoney1 { get; set; }

        ///<summary>
        /// 每星期零用錢二年
        ///</summary>
        [Field(Field = "wpm2", Indexed = false)]
        public int WeekPocketMoney2 { get; set; }

        ///<summary>
        /// 每星期零用錢三年
        ///</summary>
        [Field(Field = "wpm3", Indexed = false)]
        public int WeekPocketMoney3 { get; set; }

        ///<summary>
        /// 每星期零用錢四年
        ///</summary>
        [Field(Field = "wpm4", Indexed = false)]
        public int WeekPocketMoney4 { get; set; }

        ///<summary>
        /// 每星期零用錢五年
        ///</summary>
        [Field(Field = "wpm5", Indexed = false)]
        public int WeekPocketMoney5 { get; set; }

        ///<summary>
        /// 每星期零用錢六年
        ///</summary>
        [Field(Field = "wpm6", Indexed = false)]
        public int WeekPocketMoney6 { get; set; }

        ///<summary>
        /// 零用錢是否足夠一年
        ///</summary>
        [Field(Field = "mws1", Indexed = false)]
        public bool MoneyWhetherSufficient1 { get; set; }

        ///<summary>
        /// 零用錢是否足夠二年
        ///</summary>
        [Field(Field = "mws2", Indexed = false)]
        public bool MoneyWhetherSufficient2 { get; set; }

        ///<summary>
        /// 零用錢是否足夠三年
        ///</summary>
        [Field(Field = "mws3", Indexed = false)]
        public bool MoneyWhetherSufficient3 { get; set; }

        ///<summary>
        /// 零用錢是否足夠四年
        ///</summary>
        [Field(Field = "mws4", Indexed = false)]
        public bool MoneyWhetherSufficient4 { get; set; }

        ///<summary>
        /// 零用錢是否足夠五年
        ///</summary>
        [Field(Field = "mws5", Indexed = false)]
        public bool MoneyWhetherSufficient5 { get; set; }

    }
}
