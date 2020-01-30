using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Web;

namespace WebApplication3
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        //============================================================
        // Added by Parth for generating list of error files
        DataTable dtErrorFiles = new DataTable();
        //============================================================

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!this.IsPostBack)
            {
                gvErrorFiles.DataSource = dtErrorFiles;
                gvErrorFiles.DataBind();
            }
        }

        protected void btnConvert_Click(object sender, EventArgs e)
        {
            try
            {
                string filepath = "";
                string filename = "";
                Boolean errorFlag ;
                AddColumnsInDatatableForErrorHandling(dtErrorFiles);

                if (fileUpload.HasFiles)
                {
                    DataSet ds = new DataSet();
                    foreach (HttpPostedFile postedFile in fileUpload.PostedFiles)
                    {
                        filename = System.IO.Path.GetFileName(postedFile.FileName);
                        filepath = Server.MapPath("~/uploadedFiles/") + filename;
                        postedFile.SaveAs(Server.MapPath("~/uploadedFiles/") + filename);

                        GetTextFromPDF(filepath, filename);
                        
                        DataTable dt = new DataTable();
                        
                        errorFlag = false;
                        AddColumnsInDatatable(dt);
                        errorFlag = GetDataFromTextFile(dt, Server.MapPath("~/TempFiles/temp.txt"));

                        //============================================================
                        // Added by Parth for storing file name
                        foreach (DataRow dr in dt.Rows)
                        {
                            dr["FILENAME"] = filename.Trim();
                            dr.AcceptChanges();
                        }
                        //============================================================

                        //============================================================
                        // Added by Parth for Catching filenames with error
                        if (!errorFlag)
                        {
                            DataRow dr = dtErrorFiles.NewRow();
                            dr["FILENAME"] = filename.Trim();
                            dr["ERROR"] = "Error in Reading the file";
                            dtErrorFiles.Rows.Add(dr);
                        }
                        //============================================================

                        ds.Tables.Add(dt);
                    }

                    if(dtErrorFiles.Rows.Count > 0)
                    {
                        gvErrorFiles.DataSource = dtErrorFiles;
                        gvErrorFiles.DataBind();
                        //lbl.Visible = false;
                    }

                    ExportDataToExcel(ds, "ExportedFile");
                    
                }
            }
            catch (Exception ex)
            {

            }
        }

        /// <summary>
        /// Reading text from PDF
        /// </summary>
        /// <returns></returns>
        private void GetTextFromPDF(string filePath, string filename)
        {
            StringBuilder text = new StringBuilder();
            using (PdfReader reader = new PdfReader(filePath))
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                }
            }

            //Delete pdf file from temp folder
            System.IO.File.Delete(filePath);

            //Save all text in text files
            System.IO.File.WriteAllText(Server.MapPath("~/TempFiles/temp.txt"), text.ToString());
        }
        
        /// <summary>
        /// This function is used add header in datatable
        /// </summary>
        /// <param name="dt"></param>
        public void AddColumnsInDatatable(DataTable dt)
        {
            //============================================================
            // Added by Parth for storing file name
            dt.Columns.Add("FILENAME");
            //============================================================

            dt.Columns.Add("BUYER");
            dt.Columns.Add("Date");
            dt.Columns.Add("SHIPTERMS");
            dt.Columns.Add("PO");
            dt.Columns.Add("REFMASTERPO");
            dt.Columns.Add("REPRINT");
            dt.Columns.Add("SHIPDATE");
            dt.Columns.Add("VENDOR");
            dt.Columns.Add("SHIPTONUMBER");
            dt.Columns.Add("SHIPTO");
            dt.Columns.Add("BILLTO");
            dt.Columns.Add("DPT");
            dt.Columns.Add("SKU");
            dt.Columns.Add("UPC");
            dt.Columns.Add("VENDORPART");
            dt.Columns.Add("DESCRIPTION");
            dt.Columns.Add("RETAIL");
            dt.Columns.Add("COST");
            dt.Columns.Add("EXTCOST");
            dt.Columns.Add("CTNS");
            dt.Columns.Add("CSPK");
            dt.Columns.Add("EXTQTY");
            dt.Columns.Add("CUBE");
            dt.Columns.Add("KILOGRAMS");
            dt.Columns.Add("TOTAL");
        }

        ///============================================================
        /// Added by Parth for Generating Error File Names
        /// <summary>
        /// This function is used add header in datatable for Error files
        /// </summary>
        /// <param name="dt"></param>
        public void AddColumnsInDatatableForErrorHandling(DataTable dt)
        {
            dt.Columns.Add("FILENAME");
            dt.Columns.Add("ERROR");
        }


        /// <summary>
        /// This function is used to get data from Text file
        /// Added On: 01/29/2020
        /// Added By: Parth
        /// Edit: Changing Return type to Boolean to catch file with error
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="filepath"></param>
        public Boolean GetDataFromTextFile(DataTable dt, string filepath)
        {
            try
            {
                string Buyer = "", itemDate = "", ShipTerms = "", PO = "", RefMasterPO = "", ShipDate = "", Vendor = "",ShipToNumber="", ShipTo = "", billTo = "", total = "";

                int counter = 0, lineItem = 0, vendorLineNumber = 0, shipToLineNumber = 0, billToLineNumber = 0, BillTotIndex = 0, UPCNumberLine = 0;
                string line, retValue = "", sku = "", vendorpart = "";

                List<string> UPCNcode = new List<string>();

                int skuStart = 0, skuLength = 0, Vendorstart = 0, vendorLength = 0;

                bool IsItemSearchInNextPage = false, isRePrint = false; ;
                using (System.IO.StreamReader file = new System.IO.StreamReader(filepath))
                {
                    while ((line = file.ReadLine()) != null)
                    {
                        counter++;

                        if (IsItemSearchInNextPage)
                        {
                            if (line.Contains("DPT  SKU/UPC   VENDOR PART#    DESCRIPTION                       RETAIL     COST        EXT COST     CTNS  CSPK  EXT QTY        CUBE  KILOGRAMS"))
                            {
                                IsItemSearchInNextPage = false;
                            }
                            else
                            {
                                goto ItemToNextPage;
                            }
                        }

                        System.Console.WriteLine(line);

                        if (line.Contains("VENDOR") && Vendor == "")
                        {
                            vendorLineNumber = counter + 2;
                        }

                        if (line.Contains("SHIP TO") && shipToLineNumber == 0)
                        {
                            shipToLineNumber = counter + 2;
                        }

                        if (line.Contains("BILL TO") && billToLineNumber == 0)
                        {
                            billToLineNumber = counter + 2;
                            BillTotIndex = line.IndexOf("BILL TO");
                        }

                        if (CheckAndReturnValueFromStart(line, "BUYER:", ref retValue) && Buyer == "") { Buyer = retValue; }
                        retValue = "";

                        if (line.Contains("BUYER:") && itemDate == "")
                        {
                            string[] dateStringArray = line.Split(new string[] { "   " }, StringSplitOptions.RemoveEmptyEntries);
                            itemDate = dateStringArray[dateStringArray.Length - 1].ToString();
                        }

                        if (CheckAndReturnValueFromStart(line, "SHIP TERMS:", ref retValue) && ShipTerms == "") { ShipTerms = retValue; }
                        retValue = "";

                        if (CheckAndReturnValueFromLast(line, "PO#:", ref retValue) && PO == "") { PO = retValue; }
                        retValue = "";

                        if (CheckAndReturnValueFromLast(line, "REF MASTER PO#:", ref retValue) && RefMasterPO == "") {

                            RefMasterPO = retValue;
                            if(RefMasterPO.Contains("* * * R E P R I N T * * *"))
                            {
                                RefMasterPO = RefMasterPO.Replace("* * * R E P R I N T * * *", "");
                                RefMasterPO = RefMasterPO.TrimEnd().TrimStart();
                                isRePrint = true;
                            }
                        }
                        retValue = "";

                        if (CheckAndReturnValueFromLast(line, "SHIP DATE:", ref retValue) && ShipDate == "") { ShipDate = retValue; }
                        retValue = "";
                        
                        if (counter == vendorLineNumber)
                        {
                            Vendor = line.Trim(); ;
                        }

                        if (counter == shipToLineNumber)
                        {
                            ShipTo = line.Substring(0, line.IndexOf("   ")).Trim();
                            ShipToNumber = ShipTo.Split(' ')[0];
                        }

                        if (counter == billToLineNumber)
                        {
                            billTo = line.Substring(BillTotIndex);
                        }


                        if(line.Contains("** CONTINUED ON NEXT PAGE **"))
                        {
                            lineItem = 0;
                            IsItemSearchInNextPage = true;
                            goto ItemToNextPage;
                        }

                        if (line.Contains("DPT  SKU/UPC   VENDOR PART#    DESCRIPTION                       RETAIL     COST        EXT COST     CTNS  CSPK  EXT QTY        CUBE  KILOGRAMS"))
                        {
                            lineItem = counter + 2;
                            UPCNumberLine = counter + 3;

                            skuStart = line.IndexOf("SKU/UPC");
                            skuLength = line.IndexOf("VENDOR PART") - skuStart;


                            Vendorstart = line.IndexOf("VENDOR PART");
                            vendorLength = line.IndexOf("DESCRIPTION") - Vendorstart;
                        }

                        if (line.Contains("TOTALS:"))
                        {
                            lineItem = 0;
                            UPCNumberLine = 0;

                            string[] lineItemdata;
                            lineItemdata = line.Split(new string[] { "   " }, StringSplitOptions.RemoveEmptyEntries);
                            total = lineItemdata[1];
                        }

                        if (lineItem == counter)
                        {
                            string[] lineItemdata;
                            lineItemdata = line.Split(new string[] { "   " }, StringSplitOptions.RemoveEmptyEntries);

                            sku = line.Substring(skuStart, skuLength);
                            vendorpart = line.Substring(Vendorstart, vendorLength);


                            DataRow dr = dt.NewRow();

                            dr["BUYER"] = Buyer.Trim();
                            dr["Date"] = itemDate.Trim();
                            dr["SHIPTERMS"] = ShipTerms.Trim();
                            dr["PO"] = PO.Trim();
                            dr["REFMASTERPO"] = RefMasterPO.Trim();
                            dr["REPRINT"] = isRePrint ? "Yes" : "No";
                            dr["SHIPDATE"] = ShipDate.Trim();
                            dr["VENDOR"] = Vendor.Trim();
                            dr["SHIPTONUMBER"] = ShipToNumber.Trim();
                            dr["SHIPTO"] = ShipTo.Trim();
                            dr["BILLTO"] = billTo.Trim();
                            dr["DPT"] = lineItemdata[0].Trim();
                            dr["SKU"] = sku.Trim().Trim();
                            dr["VENDORPART"] = vendorpart.Trim();
                            dr["DESCRIPTION"] = lineItemdata[2].Trim();
                            dr["RETAIL"] = lineItemdata[3].Trim();
                            dr["COST"] = lineItemdata[4].Trim();
                            dr["EXTCOST"] = lineItemdata[5].Trim();
                            dr["CTNS"] = lineItemdata[6].Trim();
                            dr["CSPK"] = lineItemdata[7].Trim();
                            dr["EXTQTY"] = lineItemdata[8].Trim();
                            dr["CUBE"] = lineItemdata[9].Trim();
                            dr["KILOGRAMS"] = lineItemdata[10].Trim();

                            dt.Rows.Add(dr);
                            lineItem = counter + 2;
                        }

                        if(UPCNumberLine == counter)
                        {
                            UPCNcode.Add(line.Trim());
                            UPCNumberLine = counter + 2;
                        }

                        ItemToNextPage:;
                    }
                    file.Close();
                    file.Dispose();
                }


                int index = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    dr["TOTAL"] = total.Trim();
                    dr["UPC"] = UPCNcode[index];
                    dr.AcceptChanges();
                    index++;


                }


                if (dt.Rows.Count > 0)
                {
                    Console.WriteLine("Operation completed.");
                }


                System.IO.File.Delete(filepath);
                

            }
            catch (Exception ex)
            {
                return false;    
            }
            
            return true;
        }

        /// <summary>
        /// This function is used to test value from starting of file and return if matched value found
        /// </summary>
        /// <param name="line"></param>
        /// <param name="checkVal"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool CheckAndReturnValueFromStart(string line, string checkVal, ref string value)
        {

            if (line.Contains(checkVal))
            {
                value = line.Substring(line.IndexOf(checkVal) + checkVal.Length);
                value = value.Substring(0, value.IndexOf("   "));
                return true;
            }
            return false;
        }

        /// <summary>
        /// This function is used to test value from ending of file and return if matched value found
        /// </summary>
        /// <param name="line"></param>
        /// <param name="checkVal"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool CheckAndReturnValueFromLast(string line, string checkVal, ref string value)
        {

            if (line.Contains(checkVal))
            {
                value = line.Substring(line.IndexOf(checkVal) + checkVal.Length);
                //value = value.Substring(0, value.IndexOf("   "));
                return true;
            }
            return false;
        }
        
        /// <summary>
        /// This function is used to export data in excel format
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="filename"></param>
        public void ExportDataToExcel(DataSet ds, string filename)
        {
            try
            {
                string attachment = "attachment; filename=" + filename + ".xls";
                Response.ClearContent();
                Response.AddHeader("content-disposition", attachment);
                Response.ContentType = "application/vnd.ms-excel";
                string tab = "";

                DataTable dtHeader = new DataTable();
                AddColumnsInDatatable(dtHeader);

                foreach (DataColumn dc in dtHeader.Columns)
                {
                    Response.Write(tab + dc.ColumnName);
                    tab = "\t";
                }
                Response.Write("\n");
                int i;

                foreach(DataTable dt in ds.Tables)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        tab = "";
                        for (i = 0; i < dt.Columns.Count; i++)
                        {
                            Response.Write(tab + dr[i].ToString());
                            tab = "\t";
                        }
                        Response.Write("\n");
                    }
                }

                //Response.End();
                
                //============================================================
                // Added by Parth to Solve error while downloading file
                HttpContext.Current.Response.Flush(); // Sends all currently buffered output to the client.
                HttpContext.Current.Response.SuppressContent = true;  // Gets or sets a value indicating whether to send HTTP content to the client.
                HttpContext.Current.ApplicationInstance.CompleteRequest(); // Causes ASP.NET to bypass all events and filtering in the HTTP pipeline chain of execution and directly execute the EndRequest event.
                //============================================================
            }
            catch (Exception e)
            {

            }
        }
    }
}