﻿using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.UserModel;
using System.IO;


namespace MainDAL
{
    public class MyExcel
    {
        public void CreateExcel(string path, PresonModel preson, List<FamilyPresonModel> familyPresonModels)
        {
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
                {
                    HSSFWorkbook workbook = new HSSFWorkbook();
                    HSSFSheet sheet = (HSSFSheet)workbook.CreateSheet("sheet1");
                    //设置第1行列的宽度
                    sheet.SetColumnWidth(0, 10 * 256);
                    sheet.SetColumnWidth(1, 30 * 256);
                    sheet.SetColumnWidth(2, 30 * 256);
                    sheet.SetColumnWidth(3, 30 * 256);
                    HSSFRow row = (HSSFRow)sheet.CreateRow(0);
                    row.CreateCell(0).SetCellValue(preson.Name);
                    row.CreateCell(1).SetCellValue(preson.IDcare);
                    row.CreateCell(2).SetCellValue(preson.TrappedReason);
                    row.CreateCell(3).SetCellValue(preson.WorkUnits);
                    sheet = (HSSFSheet)workbook.CreateSheet("sheet2");
                    sheet.SetColumnWidth(0, 10 * 256);
                    sheet.SetColumnWidth(1, 30 * 256);
                    sheet.SetColumnWidth(2, 30 * 256);
                    sheet.SetColumnWidth(3, 30 * 256);
                    for (int i = 0; i < familyPresonModels.Count ; i++)
                    {
                        FamilyPresonModel family = familyPresonModels[i];
                        HSSFRow cells = (HSSFRow)sheet.CreateRow(i);
                        cells.CreateCell(0).SetCellValue(family.Name);
                        cells.CreateCell(1).SetCellValue(family.IDcare);
                      //  cells.CreateCell(2).SetCellValue(family.ApplicantIDcare);
                        cells.CreateCell(2).SetCellValue(family.RelationshipBetween);
                    }
                    workbook.Write(fs);
                }
            }catch(Exception e)
            {
                throw;
            }
          
        }
    }
}
