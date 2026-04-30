Attribute VB_Name = "å_ñÒé©ìÆçXêVóp"
Option Compare Database
Option Explicit

Public Function End_Day(INPUTDATE As Variant, INPUTCNT As Long) As Variant
  
    Dim WKYY    As Long
    Dim WKMM    As Long
    Dim WKDD    As Long
    Dim WKDAY   As Date
    Dim OUTDAY  As Date
    
    End_Day = Null
    
    If Trim(Nz(INPUTDATE, "")) = "" Then Exit Function
    
    WKDAY = INPUTDATE
    
    WKYY = Val(Format(WKDAY, "YYYY"))
    WKMM = Val(Format(WKDAY, "MM"))
    WKDD = Val(Format(WKDAY, "DD"))
    
    OUTDAY = DateSerial(WKYY, WKMM + INPUTCNT + 1, 0)
    
    End_Day = CDate(OUTDAY)
  
End Function
Public Function Next_Year(INPUTDATE As Variant, INPUTCNT As Long) As Variant
  
    Dim WKYY    As Long
    Dim WKMM    As Long
    Dim WKDD    As Long
    Dim WKDAY   As Date
    Dim OUTDAY  As Date
    
    Next_Year = Null
    
    If Trim(Nz(INPUTDATE, "")) = "" Then Exit Function
    
    WKDAY = INPUTDATE
    
    WKYY = Val(Format(WKDAY, "YYYY"))
    WKMM = Val(Format(WKDAY, "MM"))
    WKDD = Val(Format(WKDAY, "DD"))
    
'''''    OUTDAY = DateSerial(WKYY, WKMM + INPUTCNT, WKDD)
    OUTDAY = DateSerial(WKYY, WKMM + INPUTCNT, 1)
    
    Next_Year = CDate(OUTDAY)
  
End Function
Public Function Matu_Day(INPUTDATE As Variant) As Variant
  
    Dim WKYY    As Long
    Dim WKMM    As Long
    Dim WKDD    As Long
    Dim WKDAY   As Date
    Dim OUTDAY  As Date
    
    Matu_Day = Null
    
    If Trim(Nz(INPUTDATE, "")) = "" Then Exit Function
    
    WKDAY = INPUTDATE
    
    WKYY = Val(Format(WKDAY, "YYYY"))
    WKMM = Val(Format(WKDAY, "MM"))
    WKDD = Val(Format(WKDAY, "DD"))
    
    OUTDAY = DateSerial(WKYY, WKMM + 1, 0)
    
    Matu_Day = CDate(OUTDAY)
  
End Function
Public Function Matu_Year(INPUTDATE As Variant, INPUTCNT As Long) As Variant
  
    Dim WKYY    As Long
    Dim WKMM    As Long
    Dim WKDD    As Long
    Dim WKDAY   As Date
    Dim OUTDAY  As Date
    
    Matu_Year = Null
    
    If Trim(Nz(INPUTDATE, "")) = "" Then Exit Function
    
    WKDAY = INPUTDATE
    
    WKYY = Val(Format(WKDAY, "YYYY"))
    WKMM = Val(Format(WKDAY, "MM"))
    WKDD = Val(Format(WKDAY, "DD"))
    
    OUTDAY = DateSerial(WKYY, WKMM + INPUTCNT, 0)
    
    Matu_Year = CDate(OUTDAY)
  
End Function

