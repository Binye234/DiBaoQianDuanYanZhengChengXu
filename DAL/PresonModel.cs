using System;
using System.Collections.Generic;
using System.Text;


namespace DAL
{
    /// <summary>
    /// 申请人信息模型
    /// </summary>
    public class PresonModel
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
        /// 至困原因
        /// </summary>
        public string TrappedReason { set; get; }
        /// <summary>
        /// 工作单位
        /// </summary>
        public string WorkUnits { set; get; }
    }
}
