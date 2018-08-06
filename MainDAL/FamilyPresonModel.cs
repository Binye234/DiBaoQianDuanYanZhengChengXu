using System;
using System.Collections.Generic;
using System.Text;


namespace MainDAL
{
    /// <summary>
    /// 家庭成员模型
    /// </summary>
    public class FamilyPresonModel
    {
        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { set; get; }
        /// <summary>
        /// 身份证号码
        /// </summary>
        public string IDcare { set; get; }
        /// <summary>
        /// 申请人号码
        /// </summary>
      //  public string ApplicantIDcare { set; get; }
        /// <summary>
        /// 与申请人关系
        /// </summary>
        public string RelationshipBetween { set; get; }
    }
}
