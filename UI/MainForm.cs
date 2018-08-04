using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DAL;

namespace UI
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void FamilyOneCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (FamilyOneCheck.Checked == true)
            {
                FamilyOneNameTxt.Enabled = true;
                FamilyOneIDcareTxt.Enabled = true;
                FamilyOneBwteenCom.Enabled = true;
            }
            else
            {
                FamilyOneNameTxt.Enabled = false;
                FamilyOneIDcareTxt.Enabled = false;
                FamilyOneBwteenCom.Enabled = false;
            }
        }

        private void FamilyTwoCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (FamilyTwoCheck.Checked == true)
            {
                FamilyTwoNameTxt.Enabled = true;
                FamilyTwoIDcareTxt.Enabled = true;
                FamilyTwoBwteenCom.Enabled = true;
            }
            else
            {
                FamilyTwoNameTxt.Enabled = false;
                FamilyTwoIDcareTxt.Enabled = false;
                FamilyTwoBwteenCom.Enabled = false;
            }
        }

        private void FamilyThreeCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (FamilyThreeCheck.Checked == true)
            {
                FamilyThreeNameTxt.Enabled = true;
                FamilyThreeIDcareTxt.Enabled = true;
                FamilyThreeBwteenCom.Enabled = true;
            }
            else
            {
                FamilyThreeNameTxt.Enabled = false;
                FamilyThreeIDcareTxt.Enabled = false;
                FamilyThreeBwteenCom.Enabled = false;
            }
        }

        private void FamilyFourCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (FamilyFourCheck.Checked == true)
            {
                FamilyFourNameTxt.Enabled = true;
                FamilyFourIDcareTxt.Enabled = true;
                FamilyFourBwteenCom.Enabled = true;
            }
            else
            {
                FamilyFourNameTxt.Enabled = false;
                FamilyFourIDcareTxt.Enabled = false;
                FamilyFourBwteenCom.Enabled = false;
            }
        }

        private void FamilyFiveCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (FamilyFiveCheck.Checked == true)
            {
                FamilyFiveNameTxt.Enabled = true;
                FamilyFiveIDcareTxt.Enabled = true;
                FamilyFiveBwteenCom.Enabled = true;
            }
            else
            {
                FamilyFiveNameTxt.Enabled = false;
                FamilyFiveIDcareTxt.Enabled = false;
                FamilyFiveBwteenCom.Enabled = false;
            }
        }

        private void FamilySixCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (FamilySixCheck.Checked == true)
            {
                FamilySixNameTxt.Enabled = true;
                FamilySixIDcareTxt.Enabled = true;
                FamilySixBwteenCom.Enabled = true;
            }
            else
            {
                FamilySixNameTxt.Enabled = false;
                FamilySixIDcareTxt.Enabled = false;
                FamilySixBwteenCom.Enabled = false;
            }
        }

        private void FamilySevenCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (FamilySevenCheck.Checked == true)
            {
                FamilySevenNameTxt.Enabled = true;
                FamilySevenIDcareTxt.Enabled = true;
                FamilySevenBwteenCom.Enabled = true;
            }
            else
            {
                FamilySevenNameTxt.Enabled = false;
                FamilySevenIDcareTxt.Enabled = false;
                FamilySevenBwteenCom.Enabled = false;
            }
        }

        private void FamilyEightCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (FamilyEightCheck.Checked == true)
            {
                FamilyEightNameTxt.Enabled = true;
                FamilyEightIDcareTxt.Enabled = true;
                FamilyEightBwteenCom.Enabled = true;
            }
            else
            {
                FamilyEightNameTxt.Enabled = false;
                FamilyEightIDcareTxt.Enabled = false;
                FamilyEightBwteenCom.Enabled = false;
            }
        }

        private void CreateBtn_Click(object sender, EventArgs e)
        {
            PresonModel preson = new PresonModel();
            List<FamilyPresonModel> list = new List<FamilyPresonModel>();
           
            if (MainNameTxt.Text == "" || !Validation.IsChinese(MainNameTxt.Text))
            {
                MessageBox.Show("申请人姓名错误");
                return;
            }
            if(MainIDcareTxt.Text==""|| !Validation.IsCareNumX(MainIDcareTxt.Text) || !Validation.IsCareID(MainIDcareTxt.Text))
            {
                MessageBox.Show("申请人身份证错误");
                return;
            }
            if (MainTrappedReasonCom.Text == "")
            {
                MessageBox.Show("申请人致困原因未填");
                return;
            }
            if (MainWorkUnitsCom.Text == "")
            {
                MessageBox.Show("申请人工作单位未填");
                return;
            }
            preson.Name = MainNameTxt.Text;
            preson.IDcare = MainIDcareTxt.Text;
            preson.TrappedReason = MainTrappedReasonCom.Text;
            preson.WorkUnits = MainWorkUnitsCom.Text;

            //家庭成员
            if (FamilyOneCheck.Checked == true)
            {
                if (FamilyOneNameTxt.Text == "" || !Validation.IsChinese(FamilyOneNameTxt.Text))
                {
                    MessageBox.Show("家庭成员一姓名错误");
                    return;
                }
                if (FamilyOneIDcareTxt.Text == "" ||FamilyOneIDcareTxt.Text ==MainIDcareTxt.Text || !Validation.IsCareNumX(FamilyOneIDcareTxt.Text) || !Validation.IsCareID(FamilyOneIDcareTxt.Text))
                {
                    MessageBox.Show("家庭成员一身份证错误");
                    return;
                }
                if (FamilyOneBwteenCom.Text == "")
                {
                    MessageBox.Show("家庭成员一关系未填");
                    return;
                }
                FamilyPresonModel family = new FamilyPresonModel() { Name = FamilyOneNameTxt.Text, IDcare = FamilyOneIDcareTxt.Text, ApplicantIDcare = MainIDcareTxt.Text, RelationshipBetween = FamilyOneBwteenCom.Text };
                if (list.Where(p => p.IDcare == family.IDcare).ToList().Count != 0)
                {
                 
                    MessageBox.Show("家庭成员一身份证号码重复");
                    return;
                }
                list.Add(family);
            }

            if (FamilyTwoCheck.Checked == true)
            {
                if (FamilyTwoNameTxt.Text == "" || !Validation.IsChinese(FamilyTwoNameTxt.Text))
                {
                    MessageBox.Show("家庭成员二姓名错误");
                    return;
                }
                if (FamilyTwoIDcareTxt.Text == "" || FamilyTwoIDcareTxt.Text == MainIDcareTxt.Text || !Validation.IsCareNumX(FamilyTwoIDcareTxt.Text) || !Validation.IsCareID(FamilyTwoIDcareTxt.Text))
                {
                    MessageBox.Show("家庭成员二身份证错误");
                    return;
                }
                if (FamilyTwoBwteenCom.Text == "")
                {
                    MessageBox.Show("家庭成员二关系未填");
                    return;
                }
                FamilyPresonModel family = new FamilyPresonModel() { Name = FamilyTwoNameTxt.Text, IDcare = FamilyTwoIDcareTxt.Text, ApplicantIDcare = MainIDcareTxt.Text, RelationshipBetween = FamilyTwoBwteenCom.Text };
                if (list.Where(p => p.IDcare == family.IDcare).ToList().Count != 0)
                {
                    MessageBox.Show("家庭成员二身份证号码重复");
                    return;
                }
                list.Add(family);
            }

            if (FamilyThreeCheck.Checked == true)
            {
                if (FamilyThreeNameTxt.Text == "" || !Validation.IsChinese(FamilyThreeNameTxt.Text))
                {
                    MessageBox.Show("家庭成员三姓名错误");
                    return;
                }
                if (FamilyThreeIDcareTxt.Text == "" || FamilyThreeIDcareTxt.Text == MainIDcareTxt.Text || !Validation.IsCareNumX(FamilyThreeIDcareTxt.Text) || !Validation.IsCareID(FamilyThreeIDcareTxt.Text))
                {
                    MessageBox.Show("家庭成员三身份证错误");
                    return;
                }
                if (FamilyThreeBwteenCom.Text == "")
                {
                    MessageBox.Show("家庭成员三关系未填");
                    return;
                }
                FamilyPresonModel family = new FamilyPresonModel() { Name = FamilyThreeNameTxt.Text, IDcare = FamilyThreeIDcareTxt.Text, ApplicantIDcare = MainIDcareTxt.Text, RelationshipBetween = FamilyThreeBwteenCom.Text };
                if (list.Where(p => p.IDcare == family.IDcare).ToList().Count != 0)
                {
                    MessageBox.Show("家庭成员三身份证号码重复");
                    return;
                }
                list.Add(family);
            }

            if (FamilyFourCheck.Checked == true)
            {
                if (FamilyFourNameTxt.Text == "" || !Validation.IsChinese(FamilyFourNameTxt.Text))
                {
                    MessageBox.Show("家庭成员四姓名错误");
                    return;
                }
                if (FamilyFourIDcareTxt.Text == "" || FamilyFourIDcareTxt.Text == MainIDcareTxt.Text || !Validation.IsCareNumX(FamilyFourIDcareTxt.Text) || !Validation.IsCareID(FamilyFourIDcareTxt.Text))
                {
                    MessageBox.Show("家庭成员四身份证错误");
                    return;
                }
                if (FamilyFourBwteenCom.Text == "")
                {
                    MessageBox.Show("家庭成员四关系未填");
                    return;
                }
                FamilyPresonModel family = new FamilyPresonModel() { Name = FamilyFourNameTxt.Text, IDcare = FamilyFourIDcareTxt.Text, ApplicantIDcare = MainIDcareTxt.Text, RelationshipBetween = FamilyFourBwteenCom.Text };
                if (list.Where(p => p.IDcare == family.IDcare).ToList().Count != 0)
                {
                    MessageBox.Show("家庭成员四身份证号码重复");
                    return;
                }
                list.Add(family);
            }

            if (FamilyFiveCheck.Checked == true)
            {
                if (FamilyFiveNameTxt.Text == "" || !Validation.IsChinese(FamilyFiveNameTxt.Text))
                {
                    MessageBox.Show("家庭成员五姓名错误");
                    return;
                }
                if (FamilyFiveIDcareTxt.Text == "" || FamilyFiveIDcareTxt.Text == MainIDcareTxt.Text || !Validation.IsCareNumX(FamilyFiveIDcareTxt.Text) || !Validation.IsCareID(FamilyFiveIDcareTxt.Text))
                {
                    MessageBox.Show("家庭成员五身份证错误");
                    return;
                }
                if (FamilyFiveBwteenCom.Text == "")
                {
                    MessageBox.Show("家庭成员五关系未填");
                    return;
                }
                FamilyPresonModel family = new FamilyPresonModel() { Name = FamilyFiveNameTxt.Text, IDcare = FamilyFiveIDcareTxt.Text, ApplicantIDcare = MainIDcareTxt.Text, RelationshipBetween = FamilyFiveBwteenCom.Text };
                if (list.Where(p => p.IDcare == family.IDcare).ToList().Count != 0)
                {
                    MessageBox.Show("家庭成员五身份证号码重复");
                    return;
                }
                list.Add(family);
            }

            if (FamilySixCheck.Checked == true)
            {
                if (FamilySixNameTxt.Text == "" || !Validation.IsChinese(FamilySixNameTxt.Text))
                {
                    MessageBox.Show("家庭成员六姓名错误");
                    return;
                }
                if (FamilySixIDcareTxt.Text == "" || FamilyOneIDcareTxt.Text == MainIDcareTxt.Text || !Validation.IsCareNumX(FamilySixIDcareTxt.Text) || !Validation.IsCareID(FamilySixIDcareTxt.Text))
                {
                    MessageBox.Show("家庭成员六身份证错误");
                    return;
                }
                if (FamilySixBwteenCom.Text == "")
                {
                    MessageBox.Show("家庭成员六关系未填");
                    return;
                }
                FamilyPresonModel family = new FamilyPresonModel() { Name = FamilySixNameTxt.Text, IDcare = FamilySixIDcareTxt.Text, ApplicantIDcare = MainIDcareTxt.Text, RelationshipBetween = FamilySixBwteenCom.Text };
                if (list.Where(p => p.IDcare == family.IDcare).ToList().Count != 0)
                {
                    MessageBox.Show("家庭成员六身份证号码重复");
                    return;
                }
                list.Add(family);
            }

            if (FamilySevenCheck.Checked == true)
            {
                if (FamilySevenNameTxt.Text == "" || !Validation.IsChinese(FamilySevenNameTxt.Text))
                {
                    MessageBox.Show("家庭成员七姓名错误");
                    return;
                }
                if (FamilySevenIDcareTxt.Text == "" || FamilySevenIDcareTxt.Text == MainIDcareTxt.Text || !Validation.IsCareNumX(FamilySevenIDcareTxt.Text) || !Validation.IsCareID(FamilySevenIDcareTxt.Text))
                {
                    MessageBox.Show("家庭成员七身份证错误");
                    return;
                }
                if (FamilySevenBwteenCom.Text == "")
                {
                    MessageBox.Show("家庭成员七关系未填");
                    return;
                }
                FamilyPresonModel family = new FamilyPresonModel() { Name = FamilySevenNameTxt.Text, IDcare = FamilySevenIDcareTxt.Text, ApplicantIDcare = MainIDcareTxt.Text, RelationshipBetween = FamilySevenBwteenCom.Text };
                if (list.Where(p => p.IDcare == family.IDcare).ToList().Count != 0)
                {
                    MessageBox.Show("家庭成员七身份证号码重复");
                    return;
                }
                list.Add(family);
            }

            if (FamilyEightCheck.Checked == true)
            {
                if (FamilyEightNameTxt.Text == "" || !Validation.IsChinese(FamilyEightNameTxt.Text))
                {
                    MessageBox.Show("家庭成员八姓名错误");
                    return;
                }
                if (FamilyEightIDcareTxt.Text == "" || FamilyEightIDcareTxt.Text == MainIDcareTxt.Text || !Validation.IsCareNumX(FamilyEightIDcareTxt.Text) || !Validation.IsCareID(FamilyEightIDcareTxt.Text))
                {
                    MessageBox.Show("家庭成员八身份证错误");
                    return;
                }
                if (FamilyEightBwteenCom.Text == "")
                {
                    MessageBox.Show("家庭成员八关系未填");
                    return;
                }
                FamilyPresonModel family = new FamilyPresonModel() { Name = FamilyEightNameTxt.Text, IDcare = FamilyEightIDcareTxt.Text, ApplicantIDcare = MainIDcareTxt.Text, RelationshipBetween = FamilyEightBwteenCom.Text };
                if (list.Where(p => p.IDcare == family.IDcare).ToList().Count != 0)
                {
                    MessageBox.Show("家庭成员八身份证号码重复");
                    return;
                }
                list.Add(family);
            }

            MyExcel myExcel = new MyExcel();
            myExcel.CreateExcel(@"D:\" + MainWorkUnitsCom.Text + MainIDcareTxt.Text + ".xlsx", preson, list);
            MessageBox.Show("成功");
        }
    }
}
