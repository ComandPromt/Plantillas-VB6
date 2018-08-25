Attribute VB_Name = "modData"
Option Explicit

Public InvoiceDetails As Details

Public Type Details
    InvoiceNumber As String
    PurchaseOrder As String
    InvoiceDate As String
    SaleAmount As String
    GST As String
    Total As String
    Paid As String
    Balance As String
    DueDate As String
End Type
