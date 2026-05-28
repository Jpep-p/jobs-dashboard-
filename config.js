// ============================================================
//  Dashboard configuration — sales / pipeline view
//  Fill in the tenantId + clientId, then leave the rest alone
//  unless your SharePoint column names change.
// ============================================================

window.DASHBOARD_CONFIG = {
  // ----- From your Entra (Azure AD) app registration --------
  tenantId: "YOUR_TENANT_ID_HERE",       // Directory (tenant) ID
  clientId: "YOUR_CLIENT_ID_HERE",       // Application (client) ID

  // ----- SharePoint site + list -----------------------------
  sharePointHostname: "mpepperjoinery",        // bit before .sharepoint.com
  sitePath:           "/sites/MPJDev",         // path after .sharepoint.com
  listName:           "Project Management List",

  // ----- Column mapping (SharePoint internal names) ---------
  //  NOTE: SharePoint internal names are not the same as
  //  display names. These were pulled from your list via the
  //  REST /fields endpoint. Don't change unless your list does.
  columns: {
    // Identifiers / labels
    customer:                "Title",                            // "Customer Name" — internal name is Title
    jobNo:                   "Job_x0020_No",
    jobName:                 "Job_x0020_Name",
    projectStatus:           "Enquiry_x0020_Status",             // display = "Project Status"
    projectType:             "Project_x0020_Type",

    // Pipeline dates (drive the funnel + KPIs)
    enquiryDate:             "Enquiry_x0020_Received_x0020_Dat", // "Enquiry Received Date" (truncated at 32 chars)
    quoteDate:               "Quote_x0020_Issued_x0020_Date",
    orderDate:               "Order_x0020_Placed_x0020_Date",
    orderDueDate:            "Order_x0020_Due_x0020_Date",
    productionCompleteDate:  "Production_x0020_Complete_x0020_",  // truncated a