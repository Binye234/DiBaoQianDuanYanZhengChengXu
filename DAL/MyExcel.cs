using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;

namespace DAL
{
    public class MyExcel
    {
        public void CreateExcel(string path, PresonModel preson, List<FamilyPresonModel> familyPresonModels)
        {
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                XSSFWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("sheet1");
                //设置第1行列的宽度
                sheet.SetColumnWidth(0, 10 * 256);
                sheet.SetColumnWidth(1,30 * 256);
                sheet.SetColumnWidth(2, 30 * 256);
                sheet.SetColumnWidth(3, 30 * 256);
                IRow row = sheet.CreateRow(0);
                row.CreateCell(0).SetCellValue(preson.Name);
                row.CreateCell(1).SetCellValue(preson.IDcare);
                row.CreateCell(2).SetCellValue(preson.TrappedReason);
                row.CreateCell(3).SetCellValue(preson.WorkUnits);
                for(int i = 1; i < familyPresonModels.Count + 1; i++)
                {
                    FamilyPresonModel family = familyPresonModels[i - 1];
                    IRow cells= sheet.CreateRow(i);
                    cells.CreateCell(0).SetCellValue(family.Name);
                    cells.CreateCell(1).SetCellValue(family.IDcare);
                    cells.CreateCell(2).SetCellValue(family.ApplicantIDcare);
                    cells.CreateCell(3).SetCellValue(family.RelationshipBetween);
                }
                workbook.Write(fs);
            }
        }
    }
}
