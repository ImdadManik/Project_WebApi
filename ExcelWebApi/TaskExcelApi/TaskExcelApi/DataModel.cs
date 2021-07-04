using IronXL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;

namespace TaskExcelApi
{
    public class ExcelModel
    {
        public string ExcelFilePath { get; set; }
        public List<Validationpolicy> ValidationPolicy { get; set; }
    }

    public class Validationpolicy
    {
        public string type { get; set; }
        public List<string> values { get; set; }
        public string SheetName { get; set; }
    }

    public class Response
    {
        public List<string> messages = new List<string>();
        public string status { get; set; }
      
    }
    public static class cDAL
    {
        public static Response ReadExcel(ExcelModel model)
        {
            List<string> lstSheet = new List<string>();
            List<string> lstColumn = null;
            Response response = new Response();
            response.status = "Success";
            WorkBook _workbook = new WorkBook(model.ExcelFilePath);

            foreach (WorkSheet sheet in _workbook.WorkSheets)
                lstSheet.Add(sheet.Name.Trim());

            foreach (Validationpolicy valid in model.ValidationPolicy)
            {
                if (valid.type == "RequiredSheet")
                {
                    var lstDiffSheet = valid.values.Except(lstSheet);
                    foreach (string str in lstDiffSheet)
                    {
                        response.status = "Failure";
                        response.messages.Add("Sheet " + str + " is required but not found.");
                    }
                }
                else if (valid.type == "RequiredColumns")
                {
                    lstColumn = new List<string>();
                    foreach (WorkSheet sheet in _workbook.WorkSheets)
                    {
                        if (sheet.Name.Trim() == valid.SheetName)
                        {
                            foreach (RangeColumn col in sheet.Columns)
                            {
                                if (col.Value.ToString().Length > 0)
                                    lstColumn.Add(col.Value.ToString());
                            }

                            var lstDiffColumn = valid.values.Except(lstColumn);
                            foreach (string str in lstDiffColumn)
                            {
                                response.status = "Failure";
                                response.messages.Add(str + " not found in the " + sheet.Name);
                            }
                        }
                    }
                }
            }
           

            return response;

        }
    }
}