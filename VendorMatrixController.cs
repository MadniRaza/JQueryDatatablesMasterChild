using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using IP.Models;
namespace IP.Controllers
{
    [OutputCache(Duration = 0)]
    public class VendorMatrixController : Controller
    {
        // GET: VendorMatrix
        public ActionResult Index(string rptCode, string menuTitle)
        {
            VendorMatrixModel oModel = new VendorMatrixModel();
            oModel.ReportTitle = menuTitle;
            oModel.ReportCode = rptCode;
            TempData["ReportTitle"] = menuTitle;
            TempData["RptCode"] = rptCode;
            if (TempData["ReportTitle"] != null && TempData["RptCode"] != null)
            {
                TempData.Keep();
                cLog oLog = new cLog();
                oLog.SaveLog(menuTitle, Request.Url.PathAndQuery, rptCode);
            }
            return View(oModel);
        }
        [HttpGet]
        public JsonResult GetPartList(string report_code)
        {
            Dictionary<string, object> response = new Dictionary<string, object>();
            VendorMatrixModel oModel = new VendorMatrixModel();
            oModel.ReportCode = report_code;
            oModel.Get_Part_List();
            response["employee_rights"] = oModel.get_employee_rights(report_code);
            response["lst_all_vendors"] = oModel.lst_all_vendors;
            response["lst_parts"] = oModel.lst_parts;
            response["IsValid"] = true;
            return Json(response, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult GetEmployeeRights(string report_code)
        {
            Dictionary<string, object> response = new Dictionary<string, object>();
            VendorMatrixModel oModel = new VendorMatrixModel();
            response["employee_rights"] = oModel.get_employee_rights(report_code);
            response["IsValid"] = true;
            return Json(response, JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public FileResult GetExportedFile()
        {
            VendorMatrixModel oModel = new VendorMatrixModel();
            string employee_name = Session["EmpName"].ToString();
            ClosedXML.Excel.XLWorkbook wbook = new ClosedXML.Excel.XLWorkbook();
            wbook.Worksheets.Add(oModel.GetDtForExport(),"list");

            using (MemoryStream memoryStream = new MemoryStream())
            {
                wbook.SaveAs(memoryStream);
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "vendor_matrix_"+ employee_name + ".xlsx");
            }
        }

        [HttpPost]
        public JsonResult LoadExcelFile(HttpPostedFileBase excel_file)
        {
            Dictionary<string, object> response = new Dictionary<string, object>();
            VendorMatrixModel oModel = new VendorMatrixModel();
            string excel_file_path = "";
            if (excel_file.ContentLength > 0)
            {
                var file_name = Path.GetFileName(excel_file.FileName);
                excel_file_path = Path.Combine(Server.MapPath("~/App_Data"), file_name);

                if (System.IO.File.Exists(excel_file_path))
                    System.IO.File.Delete(excel_file_path);

                excel_file.SaveAs(excel_file_path);

                oModel.Load_Excel_File(excel_file_path);
                response["error_messages"] = oModel.error_messages;
            }
            else
            {
                response["error_messages"] = "";
            }
            return Json(response, JsonRequestBehavior.AllowGet);
        }
        [HttpGet]
        public JsonResult SaveVendors(string vendors, string part_num)
        {
            Dictionary<string, object> response = new Dictionary<string, object>();
            VendorMatrixModel oModel = new VendorMatrixModel();
            oModel.part_num = part_num;
            response["IsValid"] = oModel.Save_Part_Vendors(vendors);
            return Json(response, JsonRequestBehavior.AllowGet);
        }
    }
}