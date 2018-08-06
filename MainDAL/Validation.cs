using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace MainDAL
{
    static public class Validation
    {
        /// <summary>
        /// 身份证号码前17位乘数
        /// </summary>
        static private  int[] MulNum = { 7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2 };    
        /// <summary>
        /// 结果位数核对表
        /// </summary>
        static private  string ResultNum = "10x98765432";

        /// <summary>
        ///  身份证校验方法
        /// </summary>
        /// <param name="input">需要验证的身份证号</param>
        /// <returns>返回对错</returns>
        static public bool IsCareID(string input)
        {
            input = input.ToLower();
            int tmp = 0;
            for (int i = 0; i < 17; i++)
            {
                tmp += int.Parse(input[i].ToString()) * MulNum[i];
            }
            int r = tmp % 11;
            if (input[17].ToString() == ResultNum[r].ToString())
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 验证身份证号码格式是否正确
        /// </summary>
        /// <param name="input">需要验证的身份证号</param>
        /// <returns>返回对错</returns>
        static public bool IsCareNumX(string input)
        {
            //匹配字符串17位数字加上1位数字或者大小写X
            string regexString = "^[0-9]{17}[0-9xX]{1}$";
            return Regex.IsMatch(input, regexString);
        }
        /// <summary>
        /// 核对是否为中文
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        static public bool IsChinese(string input)
        {
            string regexString = "[\u4e00-\u9fa5]{2,30}";
            return Regex.IsMatch(input, regexString);
        }
    }
}
