// ============================================================
//  Dashboard configuration — sales / pipeline view
// ============================================================

window.DASHBOARD_CONFIG = {
  // ----- From your Entra (Azure AD) app registration --------
  tenantId: "fe828f6e-fc28-440e-9e5a-697eeaec168d",
  clientId: "35dd19a0-a2c2-4d9b-b144-2556efd3f5f5",

  // ----- SharePoint site + list -----------------------------
  sharePointHostname: "mpepperjoinery",
  sitePath:           "/sites/MPJDev",
  listName:           "Project Management List",

  // ----- Column mapping (SharePoint internal names) ---------
  columns: {
    customer:                "Title",
    jobNo:                   "Job_x0020_No",
    jobName:                 "Job_x0020_Name",
    projectStatus:           "Enquiry_x0020_Status",
    projectType:             "Project_x0020_Type",
    enquiryDate:             "Enquiry_x0020_Received_x0020_Dat",
    quoteDate:               "Quote_x0020_Issued_x0020_Date",
    orderDate:               "Order_x0020_Placed_x0020_Date",
    orderDueDate:            "Order_x0020_Due_x0020_Date",
    productionCompleteDate:  "Production_x0020_Complete_x0020_",
    declinedDate:            "Quote_x0020_Declined_x0020_Date",
    declinedReason:          "Declined_x0020_Reason",
    netValue:                "Order_x0020_Net_x0020_Value",
    liveProject:             "Live_x0020_Project",
    pricedBy:                "Priced_x0020_by"
  },

  // ----- Display behaviour ----------------------------------
  refreshMinutes:    5,
  rowsInTable:       10,
  rotateTable:       true,
  pageRotateSeconds: 15,
  currencySymbol:    "£",
  locale:            "en-GB"
};
