Attribute VB_Name = "MainEntryPoint"
Option Explicit

Private Const errOffset = 512

Public errReadOnlyRun As AddressError
Public errPropDesignOnly As AddressError
Public errUKUseOnly As AddressError
Public errUSUseOnly As AddressError
Public errEmailNotAvailable As AddressError
Public errCountryNotAvailable As AddressError

Type AddressError
  Number As Long
  Description As String
End Type

Sub Main()

  errReadOnlyRun.Number = vbObjectError + errOffset + 1
  errReadOnlyRun.Description = "Property is read-only at run time."         'Use in Prop Let.
  
  errPropDesignOnly.Number = vbObjectError + errOffset + 2
  errPropDesignOnly.Description = "Property is not available at run time."  'Use in Prop Get.
  
  errUKUseOnly.Number = vbObjectError + errOffset + 3
  errUKUseOnly.Description = "Property only available when Country = 1 - UK."
  
  errUSUseOnly.Number = vbObjectError + errOffset + 4
  errUSUseOnly.Description = "Property only available when Country = 0 - US."
  
  errEmailNotAvailable.Number = vbObjectError + errOffset + 5
  errEmailNotAvailable.Description = "Property only available when View_Email = 0 - Show_Email."

  errCountryNotAvailable.Number = vbObjectError + errOffset + 6
  errCountryNotAvailable.Description = "Property only available when View_Country = 0 - Show_Country."
  
End Sub


