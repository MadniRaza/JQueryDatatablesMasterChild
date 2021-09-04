using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;



namespace IP.Models
{
    public class VendorMatrixModel
    {
        #region Data Fields
        public string ReportTitle { get; set; }
        public string ReportCode { get; set; }
        public List<ArrayList> lst_parts { get; set; }
        public IList<SelectListItem> lst_all_vendors { get; set; }
        public string part_num { get; set; }
        public List<string> invalid_vendors { get; set; }
        public List<string> invalid_parts { get; set; }
        public List<string> invalid_alloc_values { get; set; }
        public List<string> invalid_UP_values { get; set; }
        public List<string> invalid_LT_values { get; set; }
        public List<string> invalid_order_values { get; set; }
        public List<string> invalid_part_alloc { get; set; }
        public List<string> duplicate_vendors { get; set; }
        public List<string> error_messages { get; set; }


        #endregion

        #region Methods
        public ArrayList get_employee_rights(string report_code)
        {
            if (report_code.Equals(string.Empty))
            {
                return null;
            }
            return cEmployee.GetEmployeeAxs(report_code);
        }
        public void Get_Part_List()
        {
            cDAL oDAL = new cDAL(cDAL.ConnectionType.CWDB);
            string query = string.Empty;
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("SELECT  '' as default_column, PA.PartNum, PA.PartDesc, VA.VendorID,");
            sb.AppendLine("VenOrder, Alloc, UnitPrice, VendorLT, CASE WHEN  Sourcing IS NULL THEN 'N' ELSE SOURCING END SOURCING ");
            sb.AppendLine("FROM Mrp.Part_Attributes PA ");
            sb.AppendLine("LEFT OUTER JOIN zVendor_Attributes VA ON VA.Company = PA.Company ");
            sb.AppendLine("AND VA.Plant = PA.Plant ");
            sb.AppendLine("AND VA.PartNum = PA.PartNum ");
            sb.AppendLine("WHERE PA.Company = '" + HttpContext.Current.Session["EpicorCompany"].ToString() + "' ");
            sb.AppendLine("AND PA.Plant = '" + HttpContext.Current.Session["EpicorPlant"].ToString() + "' ");
            sb.AppendLine("GROUP BY PA.PartNum, PA.PartDesc, VA.VendorID,");
            sb.AppendLine("VenOrder, Alloc, UnitPrice, VendorLT, Sourcing");
            sb.AppendLine();
            DataTable dt_parts = oDAL.GetData(sb.ToString());
            if (!ReportCode.Equals(string.Empty))
            {
                //For SQL Documentation
                cLog oLog = new cLog();
                oLog.AddSqlQuery(ReportCode, sb.ToString(), string.Empty, false);
            }


            query = "SELECT VendorID, Name as vendor_name FROM Erp.Vendor WHERE InActive = 0 ";
            query += "AND Company = '" + HttpContext.Current.Session["EpicorCompany"].ToString() + "' ";
            lst_all_vendors = oDAL.GetData(query).AsEnumerable().Select(dataRow => new SelectListItem
            {
                Text = dataRow["VendorID"].ToString() + ";" + dataRow["vendor_name"].ToString(),
                Value = dataRow["VendorID"].ToString(),
            }).ToList<SelectListItem>();

            DataTable dt_parts_tmp = dt_parts.DefaultView.ToTable(true, "default_column", "PartNum", "PartDesc", "Sourcing");
            DataColumn vendor_column = new DataColumn();
            vendor_column.ColumnName = "Vendors";
            dt_parts_tmp.Columns.Add(vendor_column);

            string part_num = "";
            string vendors_params = string.Empty;
            for (int i = 0; i < dt_parts_tmp.Rows.Count; i++)
            {
                part_num = dt_parts_tmp.Rows[i]["PartNum"].ToString();

                dt_parts.DefaultView.RowFilter = "PartNum = '" + part_num + "'";
                DataTable dt_part_vendors = dt_parts.DefaultView.ToTable();
                int total_vendors = dt_part_vendors.Rows.Count;
                vendors_params = string.Empty;
                for (int j = 0; j < dt_part_vendors.Rows.Count; j++)
                {
                    string vendor_id = dt_part_vendors.Rows[j]["VendorID"].ToString();
                    if (vendor_id == string.Empty || vendor_id.Trim().Length == 0)
                        continue;
                    vendors_params += dt_part_vendors.Rows[j]["VendorID"].ToString();
                    vendors_params += "»»»";
                    vendors_params += dt_part_vendors.Rows[j]["VenOrder"].ToString();
                    vendors_params += "»»»";
                    vendors_params += dt_part_vendors.Rows[j]["Alloc"].ToString();
                    vendors_params += "»»»";
                    vendors_params += dt_part_vendors.Rows[j]["UnitPrice"].ToString();
                    vendors_params += "»»»";
                    vendors_params += dt_part_vendors.Rows[j]["VendorLT"].ToString();

                    if (j + 1 != total_vendors)
                        vendors_params += "┘";
                }
                dt_parts_tmp.Rows[i]["Vendors"] = vendors_params;
            }

            lst_parts = cCommon.ConvertDtToArrayList(dt_parts_tmp);
        }
        public List<Dictionary<string, object>> ConvertDtToList(DataTable dt)
        {
            List<Dictionary<string, object>>
            lstRows = new List<Dictionary<string, object>>();
            Dictionary<string, object> dictRow = null;

            foreach (DataRow dr in dt.Rows)
            {
                dictRow = new Dictionary<string, object>();
                foreach (DataColumn col in dt.Columns)
                {
                    dictRow.Add(col.ColumnName, dr[col]);
                }
                lstRows.Add(dictRow);
            }
            return lstRows;
        }

        public bool Save_Part_Vendors(string vendors)
        {
            cDAL portal_db = new cDAL(cDAL.ConnectionType.CWDB);
            string query = "";

            int employee_id = Convert.ToInt32(HttpContext.Current.Session["EmpId"]);

            decimal total_alloc = 0;
            if (vendors.Length > 0 && vendors != string.Empty)
            {
                string[] rows = vendors.Split('┘');
                for (int i = 0; i < rows.Length; i++)
                {

                    string[] columns = rows[i].Split(new string[] { "»»»" }, StringSplitOptions.None);
                    string vendor_id = columns[0];
                    int order = Convert.ToInt32(columns[1]);
                    decimal alloc = Convert.ToDecimal(columns[2]);
                    decimal UP = Convert.ToDecimal(columns[3]);
                    int LT = Convert.ToInt32(columns[4]);
                    char row_type = Convert.ToChar(columns[5]);

                    if (row_type == 'I')
                    {
                        query = "INSERT INTO zVendor_Attributes(Company, Plant, PartNum, Sourcing, VendorID, VenOrder, Alloc, UnitPrice, VendorLT, InsertedBy) ";
                        query += "VALUES('" + HttpContext.Current.Session["EpicorCompany"].ToString() + "',";
                        query += "'" + HttpContext.Current.Session["EpicorPlant"].ToString() + "',";
                        query += "'" + part_num + "',";
                        query += "'Y',";
                        query += "'" + vendor_id + "',";
                        query += "'" + order + "',";
                        query += "'" + alloc + "',";
                        query += "'" + UP + "',";
                        query += "'" + LT + "',";
                        query += "'" + employee_id + "'";
                        query += ")";
                        portal_db.AddQuery(query);
                    }
                    else if (row_type == 'U')
                    {
                        query = "UPDATE zVendor_Attributes ";
                        query += "SET Alloc = " + alloc + ", ";
                        query += "VenOrder = " + order + ", ";
                        query += "UnitPrice = " + UP + ", ";
                        query += "UpdatedBy = " + employee_id + ", ";
                        query += "UpdatedOn =GETDATE(), ";
                        query += "VendorLT = " + LT + " ";
                        query += "WHERE Company='" + HttpContext.Current.Session["EpicorCompany"].ToString() + "' ";
                        query += "AND Plant='" + HttpContext.Current.Session["EpicorPlant"].ToString() + "' ";
                        query += "AND PartNum = '" + part_num + "' AND VendorID = '" + vendor_id + "' ";
                        portal_db.AddQuery(query);
                    }
                    else if (row_type == 'D')
                    {
                        query = "DELETE FROM zVendor_Attributes ";
                        query += "WHERE Company='" + HttpContext.Current.Session["EpicorCompany"].ToString() + "' ";
                        query += "AND Plant='" + HttpContext.Current.Session["EpicorPlant"].ToString() + "' ";
                        query += "AND PartNum = '" + part_num + "' AND VendorID = '" + vendor_id + "' ";
                        portal_db.AddQuery(query);
                    }

                }

                portal_db.Commit();
                if (portal_db.HasErrors)
                    return false;

                return true;
            }
            return true;

        }

        public DataTable GetDtForExport()
        {
            cDAL oDAL = new cDAL(cDAL.ConnectionType.CWDB);
            string query = string.Empty;

            query = "SELECT PartNum, VenOrder, VendorID, Alloc, UnitPrice, VendorLT FROM zVendor_Attributes ";
            query += "WHERE Company = '" + HttpContext.Current.Session["EpicorCompany"].ToString() + "' ";
            query += "AND Plant = '" + HttpContext.Current.Session["EpicorPlant"].ToString() + "' ";
            DataTable dt_parts = oDAL.GetData(query);
            if (dt_parts == null)
                return null;
            DataTable dt_parts_tmp = dt_parts.DefaultView.ToTable(true, "PartNum");
            int max_vendors_for_x_part = (from p in dt_parts.AsEnumerable()
                                          let columns = new { c1 = p.Field<string>("PartNum") }
                                          group p by columns into g
                                          select g.Count()).Max();



            string[] column_names = { "Order", "VenID", "Allc", "UP", "LT" };
            int no_of_columns = max_vendors_for_x_part * 5; // per vendor columns are 5
            int index = 0;
            int vendor_seq = 1;
            for (int i = 0; i < no_of_columns; i++)
            {
                DataColumn col = new DataColumn();
                if (index > 4)
                {
                    index = 0;
                    vendor_seq += 1;
                }

                col.ColumnName = column_names[index] + "_" + vendor_seq;
                index += 1;
                dt_parts_tmp.Columns.Add(col);
            }
            string part_num = "";
            string vendors_params = string.Empty;
            for (int i = 0; i < dt_parts_tmp.Rows.Count; i++)
            {
                part_num = dt_parts_tmp.Rows[i]["PartNum"].ToString();
                dt_parts.DefaultView.RowFilter = "PartNum = '" + part_num + "'";
                DataTable dt_part_vendors = dt_parts.DefaultView.ToTable();

                for (int j = 0; j < dt_part_vendors.Rows.Count; j++)
                {
                    dt_parts_tmp.Rows[i]["Order_" + (j + 1)] = dt_part_vendors.Rows[j]["VenOrder"].ToString();
                    dt_parts_tmp.Rows[i]["VenID_" + (j + 1)] = dt_part_vendors.Rows[j]["VendorID"].ToString();
                    dt_parts_tmp.Rows[i]["Allc_" + (j + 1)] = dt_part_vendors.Rows[j]["Alloc"].ToString();
                    dt_parts_tmp.Rows[i]["UP_" + (j + 1)] = dt_part_vendors.Rows[j]["UnitPrice"].ToString();
                    dt_parts_tmp.Rows[i]["LT_" + (j + 1)] = dt_part_vendors.Rows[j]["VendorLT"].ToString();
                }
            }

            return dt_parts_tmp;


        }

        public DataTable ExcelToDataTable(string filePath, string sheetName)
        {
            try
            {
                // Open the Excel file using ClosedXML.
                // Keep in mind the Excel file cannot be open when trying to read it
                using (XLWorkbook workBook = new XLWorkbook(filePath))
                {
                    //Read the first Sheet from Excel file.
                    IXLWorksheet workSheet = workBook.Worksheet(1);

                    //Create a new DataTable.
                    DataTable dt = new DataTable();

                    //Loop through the Worksheet rows.
                    bool firstRow = true;
                    foreach (IXLRow row in workSheet.Rows())
                    {
                        //Use the first row to add columns to DataTable.
                        if (firstRow)
                        {
                            foreach (IXLCell cell in row.Cells())
                            {
                                dt.Columns.Add(cell.Value.ToString());
                            }
                            firstRow = false;
                        }
                        else
                        {
                            //Add rows to DataTable.
                            dt.Rows.Add();
                            int i = 0;
                            foreach (IXLCell cell in row.Cells())
                            {
                                dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                                i++;
                            }
                            //    foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                            //{
                            //    dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            //    i++;
                            //}
                        }
                    }

                    return dt;
                }

            }
            catch (Exception ex)
            {
                return null;
            }
        }


        public void Load_Excel_File(string excel_file_path)
        {
            error_messages = new List<string>();
            DataTable main_dt = ExcelToDataTable(excel_file_path, "list");
            if (main_dt == null)
            {
                error_messages.Add("Uploaded excel file is invalid.");
                return;

            }
            int total_columns = main_dt.Columns.Count;
            int vendor_columns_count = total_columns - 1; //substract part num column
            int columns_per_vendor = 5; //e.g. Order, LT, UP, Alloc, VenID

            if (vendor_columns_count % columns_per_vendor != 0)
            { 
                error_messages.Add("Invalid Columns found in uploaded file");
                return;
            }

            int total_vendors = vendor_columns_count / columns_per_vendor;

            string[] vendor_attributes_columns = { "PartNum", "VenOrder", "VendorID", "Alloc", "UnitPrice", "VendorLT" };
            DataTable dt_tmp = new DataTable();

            for (int i = 0; i < vendor_attributes_columns.Length; i++)
            {
                dt_tmp.Columns.Add(vendor_attributes_columns[i]);
            }

            DataRow dr;

            for (int i = 0; i < main_dt.Rows.Count; i++)
            {
                int count = 1; // 0 is part num, columns counting starting from 1
                for (int j = 1; j <= total_vendors; j++)
                {
                    dr = dt_tmp.NewRow();
                    dr["PartNum"] = main_dt.Rows[i][0].ToString();
                    for (int k = 1; k < (columns_per_vendor + 1); k++) //1 is k value
                    {
                        dr[k] = main_dt.Rows[i][count];
                        count += 1; //read next column
                    }

                    if (dr[1].ToString().Trim().Length > 0) // check if row has vendor then insert
                        dt_tmp.Rows.Add(dr);
                }
            }

            //excel to datatable conversion is completed, dt_tmp contains all vendors part wise

            // Now, Validating Each Part and inserting/updating in zVendor_Attributes.
            //Step 1 (Populating Records)
            cDAL oDAL = new cDAL(cDAL.ConnectionType.CWDB);
            int employee_id = Convert.ToInt32(HttpContext.Current.Session["EmpId"]);

            var excel_parts = (from p in dt_tmp.AsEnumerable()
                               group p by p.Field<string>("PartNum") into g
                               select g.Key).ToList();
            string excel_parts_1 = String.Join("','", excel_parts).Insert(0, "'").Insert(String.Join("','", excel_parts).Insert(0, "'").Length, "'");
            string query_1 = "SELECT PartNum from Mrp.Part_Attributes ";
            query_1 += "WHERE Company = '" + HttpContext.Current.Session["EpicorCompany"].ToString() + "' ";
            query_1 += "AND Plant = '" + HttpContext.Current.Session["EpicorPlant"].ToString() + "' ";
            query_1 += "AND PartNum IN (" + excel_parts_1 + ") ";
            DataTable dt_erp_parts = oDAL.GetData(query_1);

            var excel_vendors = (from p in dt_tmp.AsEnumerable()
                                 group p by p.Field<string>("VendorID") into g
                                 select g.Key).ToList();
            string excel_vendors_1 = String.Join("','", excel_vendors).Insert(0, "'").Insert(String.Join("','", excel_vendors).Insert(0, "'").Length, "'");
            string query_2 = "SELECT VendorID from Erp.Vendor ";
            query_2 += "WHERE Company = '" + HttpContext.Current.Session["EpicorCompany"].ToString() + "' ";
            query_2 += "AND InActive = 0 AND VendorID IN (" + excel_vendors_1 + ") ";
            DataTable dt_erp_vendors = oDAL.GetData(query_2);

            string query_3 = "SELECT PartNum, VendorID from zVendor_Attributes ";
            query_3 += "WHERE Company = '" + HttpContext.Current.Session["EpicorCompany"].ToString() + "' ";
            query_3 += "AND Plant = '" + HttpContext.Current.Session["EpicorPlant"].ToString() + "' ";
            DataTable dt_vendor_attributes_x = oDAL.GetData(query_3);

            //Step 2) Processing Records
            StringBuilder queries = new StringBuilder();
            invalid_vendors = new List<string>();
            invalid_alloc_values = new List<string>();
            invalid_LT_values = new List<string>();
            invalid_order_values = new List<string>();
            invalid_parts = new List<string>();
            invalid_part_alloc = new List<string>();
            invalid_UP_values = new List<string>();
            invalid_vendors = new List<string>();
            duplicate_vendors = new List<string>();


            //REMOVE PARTS FROM Vendor_Attributes, Doesn't Exist in Excel File
            string query_4 = "DELETE FROM zVendor_Attributes ";
            query_4 += "WHERE Company = '" + HttpContext.Current.Session["EpicorCompany"].ToString() + "' ";
            query_4 += "AND Plant = '" + HttpContext.Current.Session["EpicorPlant"].ToString() + "' ";
            query_4 += "AND PartNum NOT IN(" + excel_parts_1 + ")";
            queries.AppendLine(query_4);

           DataTable dt_tmp_parts = dt_tmp.DefaultView.ToTable(true, "PartNum");
            //DataView dv_tmp_parts = dt_tmp.DefaultView;
            //dv_tmp_parts.Sort = "PartNum Asc";
            //DataTable dt_tmp_parts = dv_tmp_parts.ToTable(true, "PartNum");
            for (int i = 0; i < dt_tmp_parts.Rows.Count; i++)
            {
                string part_num_1 = dt_tmp_parts.Rows[i]["PartNum"].ToString();

                //check part existance in erp
                DataRow[] part = dt_erp_parts.Select("PartNum = '" + part_num_1 + "'");
                if (part.Length == 0)
                {
                    invalid_parts.Add(part_num_1);
                    continue;
                }

                dt_tmp.DefaultView.RowFilter = "PartNum = '" + part_num_1 + "'";
                DataTable dt_tmp_vendors = dt_tmp.DefaultView.ToTable();

                StringBuilder tmp_queries = new StringBuilder();
                decimal alloc_sum = 0;
                ArrayList vendors = new ArrayList();
                for (int j = 0; j < dt_tmp_vendors.Rows.Count; j++)
                {
                    if (is_row_empty(dt_tmp_vendors.Rows[j]))
                        continue;
                    string vendor_id = dt_tmp_vendors.Rows[j]["VendorID"].ToString();

                    bool correct_row = true;



                    //check vendor existance in ERP
                    DataRow[] vendor = dt_erp_vendors.Select("VendorID = '" + vendor_id + "'");
                    if (vendor.Length == 0)
                    {
                        invalid_vendors.Add(vendor_id);
                        continue;
                    }
                    //CHECK IF DUPLICATE VENDORS
                    if (vendors.IndexOf(vendor_id) >= 0)
                    {
                        duplicate_vendors.Add(vendor_id);
                        continue;
                    }
                    vendors.Add(vendor_id);

                    string alloc = dt_tmp_vendors.Rows[j]["Alloc"].ToString();
                    string order = dt_tmp_vendors.Rows[j]["VenOrder"].ToString();
                    string LT = dt_tmp_vendors.Rows[j]["VendorLT"].ToString();
                    string UP = dt_tmp_vendors.Rows[j]["UnitPrice"].ToString();

                    //validate numeric and decimal values
                    if (!cCommon.IsDecimal(alloc) || Convert.ToDecimal(alloc) <= 0)
                    {
                        invalid_alloc_values.Add(part_num_1 + " > " + vendor_id + ">" + alloc);
                        correct_row = false;
                    }

                    if (!cCommon.IsNumber(LT) || Convert.ToInt32(LT) <= 0)
                    {
                        invalid_LT_values.Add(part_num_1 + " > " + vendor_id + ">" + LT);
                        correct_row = false;
                    }
                    if (!cCommon.IsNumber(order) || Convert.ToInt32(order) <= 0)
                    {
                        invalid_order_values.Add(part_num_1 + " > " + vendor_id + ">" + order);
                        correct_row = false;
                    }
                    if (!cCommon.IsDecimal(UP) || Convert.ToDecimal(UP) <= 0)
                    {
                        invalid_UP_values.Add(part_num_1 + " > " + vendor_id + " > " + UP);
                        correct_row = false;
                    }

                    if (correct_row)
                    {

                        string query = "";
                        //check if part has already vendor
                        DataRow[] exist_in_ven_attr = dt_vendor_attributes_x.Select("PartNum = '" + part_num_1 + "' AND VendorID = '" + vendor_id + "'");
                        if (exist_in_ven_attr.Length > 0)
                        {
                            query = "UPDATE zVendor_Attributes ";
                            query += "SET Alloc = " + alloc + ", ";
                            query += "VenOrder = " + order + ", ";
                            query += "UnitPrice = " + UP + ", ";
                            query += "UpdatedBy = " + employee_id + ", ";
                            query += "UpdatedOn =GETDATE(), ";
                            query += "VendorLT = " + LT + ", ";
                            query += "Sourcing = 'Y' ";
                            query += "WHERE Company='" + HttpContext.Current.Session["EpicorCompany"].ToString() + "' ";
                            query += "AND Plant='" + HttpContext.Current.Session["EpicorPlant"].ToString() + "' ";
                            query += "AND PartNum = '" + part_num_1 + "' AND VendorID = '" + vendor_id + "' ";
                        }
                        else
                        {
                            query = "INSERT INTO zVendor_Attributes(Company, Plant, PartNum, Sourcing, VendorID, VenOrder, Alloc, UnitPrice, VendorLT, InsertedBy) ";
                            query += "VALUES('" + HttpContext.Current.Session["EpicorCompany"].ToString() + "',";
                            query += "'" + HttpContext.Current.Session["EpicorPlant"].ToString() + "',";
                            query += "'" + part_num_1 + "',";
                            query += "'Y',";
                            query += "'" + vendor_id + "',";
                            query += "'" + order + "',";
                            query += "'" + alloc + "',";
                            query += "'" + UP + "',";
                            query += "'" + LT + "',";
                            query += "'" + employee_id + "'";
                            query += ")";
                        }
                        tmp_queries.AppendLine(query);
                        alloc_sum += Convert.ToDecimal(alloc);
                    }
                }

                //REMOVE Vendors FROM Vendor_Attributes, Doesn't Exist in Excel File
                List<string> lst_1 = vendors.Cast<string>().ToList();
                string excel_part_vendors = String.Join("','", lst_1).Insert(0, "'").Insert(String.Join("','", lst_1).Insert(0, "'").Length, "'");
                string query_5 = "DELETE FROM zVendor_Attributes ";
                query_5 += "WHERE Company = '" + HttpContext.Current.Session["EpicorCompany"].ToString() + "' ";
                query_5 += "AND Plant = '" + HttpContext.Current.Session["EpicorPlant"].ToString() + "' ";
                query_5 += "AND PartNum = '" + part_num_1 + "' AND VendorID NOT IN(" + excel_part_vendors + ")";
                queries.AppendLine(query_5);
                if (alloc_sum != 1)
                {
                    invalid_part_alloc.Add(part_num_1 + " > " + alloc_sum);
                    continue;
                }
                //FINALY MODIFY IN SYSTEM
                queries.AppendLine(tmp_queries.ToString());
            }



            oDAL.AddQuery(queries.ToString());
            oDAL.Commit();

            //EMAILING AND ALERTS

            string email_body = "";
            string email_subject = "";
            if (invalid_parts.Count > 0)
            {
                error_messages.Add("<p><strong>Following Parts not found in ERP.</strong></br>" + String.Join(", ", invalid_parts) + "</p>");
                email_subject = "Vendor Matrix > Invalid Parts";
                email_body = "\"" + String.Join(", ", invalid_parts) + "\"";
                cLog.SendEmail("VENDOR_ATTRIBUTES", email_subject, email_body);
            }
            if (invalid_vendors.Count > 0)
            {
                error_messages.Add("<p><strong>Following Vendors not found in ERP.</strong></br>" + String.Join(", ", invalid_vendors) + "</p>");
                email_subject = "Vendor Matrix > Invalid Vendors";
                email_body = "\"" + String.Join(", ", invalid_vendors) + "\"";
                cLog.SendEmail("VENDOR_ATTRIBUTES", email_subject, email_body);
            }
            if (duplicate_vendors.Count > 0)
            {
                error_messages.Add("<p><strong>Following vendors are found twice in excel file.</strong></br>" + String.Join(", ", duplicate_vendors) + "</p>");
                email_subject = "Vendor Matrix > Duplicate Vendors found";
                email_body = "\"" + String.Join(", ", invalid_vendors) + "\"";
                cLog.SendEmail("VENDOR_ATTRIBUTES", email_subject, email_body);
            }

            if (invalid_order_values.Count > 0)
            {
                error_messages.Add("<p><strong>Following Vendor order values are incorrect:</strong></br>" + String.Join(", ", invalid_order_values) + "</p>");
                email_subject = "Vendor Matrix > Invalid Vendor Order(s)";
                email_body = "\"" + String.Join(", ", invalid_order_values) + "\"";
                cLog.SendEmail("VENDOR_ATTRIBUTES", email_subject, email_body);
            }
            if (invalid_alloc_values.Count > 0)
            {
                error_messages.Add("<p><strong>Following Vendor order alloc. are incorrect:</strong></br>" + String.Join(", ", invalid_alloc_values) + "</p>");
                email_subject = "Vendor Matrix > Invalid Alloc% Values";
                email_body = "\"" + String.Join(", ", invalid_alloc_values) + "\"";
                cLog.SendEmail("VENDOR_ATTRIBUTES", email_subject, email_body);
            }
            if (invalid_LT_values.Count > 0)
            {
                error_messages.Add("<p><strong>Following Vendor LT values are incorrect:</strong></br>" + String.Join(", ", invalid_LT_values) + "</p>");
                email_subject = "Vendor Matrix > Invalid LT Values";
                email_body = "\"" + String.Join(", ", invalid_LT_values) + "\"";
                cLog.SendEmail("VENDOR_ATTRIBUTES", email_subject, email_body);
            }
            if (invalid_UP_values.Count > 0)
            {
                error_messages.Add("<p><strong>Following Vendor UP values are incorrect:</strong></br>" + String.Join(", ", invalid_UP_values) + "</p>");
                email_subject = "Vendor Matrix > Invalid UP Values";
                email_body = "\"" + String.Join(", ", invalid_UP_values) + "\"";
                cLog.SendEmail("VENDOR_ATTRIBUTES", email_subject, email_body);
            }
            if (invalid_part_alloc.Count > 0)
            {
                error_messages.Add("<p><strong>Sum of Alloc% for the following parts are incorrect:</strong></br>" + String.Join(", ", invalid_part_alloc) + "</p>");
                email_subject = "Vendor Matrix > Invalid Sum of Alloc%";
                email_body = "\"" + String.Join(", ", invalid_part_alloc) + "\"";
                cLog.SendEmail("VENDOR_ATTRIBUTES", email_subject, email_body);
            }



        }
        private bool is_row_empty(DataRow dr)
        {
            bool is_row_empty = true;
            for (int i = 0; i < dr.ItemArray.Length; i++)
            {
                if (dr.ItemArray[i].ToString().Trim().Length > 0)
                    return false;

            }
            return is_row_empty;
        }
        #endregion
    }
}