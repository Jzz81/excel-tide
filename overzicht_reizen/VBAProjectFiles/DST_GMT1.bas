Attribute VB_Name = "DST_GMT1"
Option Explicit
Option Private Module

 ' For more on DST/Summer Time: http://webexhibits.org/daylightsaving/
 
 ' ISO country codes: http://www.iso.org/iso/en/prods-services/iso3166ma/02iso-3166-code-lists/list-en1.html
 
 ' The function ConvertToGMT relies on the included function NthWeekday to help with
 ' Daylight Saving/Summer Time calculation.  The version listed here is set up to work in
 ' Excel and not Access.  The comments to NthWeekday explain the easy steps required
 ' to make this function compatible with Access
 
Public Function convert_dt_to_xml_string(dt As Date) As String
'2015-02-02T08:22:00Z
convert_dt_to_xml_string = Format(dt, "yyyy-mm-ddThh:nn:ssZ")
End Function
Public Function convert_xml_string_to_dt(dt As String) As Date
'converts an xml formatted string to a usable date format
convert_xml_string_to_dt = ConvertToLT(CDate(Left(dt, 10)) + CDate(Mid(dt, 12, 8)))
End Function
 
Function ConvertToGMT(LocalTime As Date) As Date
     
' LocalTime is datetime in local
' GTM_Adjust is the normal number of hours you add to or subtract from GMT to get local standard time
'     Function turns hours into minutes in DateAdd operations to accommodate half-hour GMT adjustments
     
Dim StartDST As Date
Dim EndDST As Date
Dim GMT_Adjust As Long

GMT_Adjust = 1
 
StartDST = DateAdd("h", 1 + GMT_Adjust, NthWeekday("L", 1, 3, Year(LocalTime)))
EndDST = DateAdd("h", 1 + GMT_Adjust, NthWeekday("L", 1, 10, Year(LocalTime)))
 
If LocalTime >= StartDST And LocalTime <= EndDST Then
    ConvertToGMT = DateAdd("h", -(GMT_Adjust + 1), LocalTime)
Else
    ConvertToGMT = DateAdd("h", -GMT_Adjust, LocalTime)
End If
     
End Function
Function ConvertToLT(GMT As Date) As Date
     
' GMT is datetime
' GTM_Adjust is the normal number of hours you add to or subtract from GMT to get local standard time
'     Function turns hours into minutes in DateAdd operations to accommodate half-hour GMT adjustments
     
Dim StartDST As Date
Dim EndDST As Date
Dim GMT_Adjust As Long
Dim LTZ As Date

GMT_Adjust = 1

'first start by converting to the local time zone
LTZ = DateAdd("h", GMT_Adjust, GMT)

StartDST = DateAdd("h", 1 + GMT_Adjust, NthWeekday("L", 1, 3, Year(GMT)))
EndDST = DateAdd("h", 1 + GMT_Adjust, NthWeekday("L", 1, 10, Year(GMT)))
 
If LTZ >= StartDST And LTZ <= EndDST Then
    ConvertToLT = DateAdd("h", 1, LTZ)
Else
    ConvertToLT = LTZ
End If
     
End Function
 
 
Public Function NthWeekday(Position As Variant, DayIndex As Long, TargetMonth As Long, Optional TargetYear As Long)
     
     ' Returns any arbitrary weekday (the "Nth" weekday) of a given month
     ' Position is the weekday's position in the month.  Must be a number 1-5, or the letter L (last)
     ' DayIndex is weekday: 1=Sunday, 2=Monday, ..., 7=Saturday
     ' TargetMonth is the month the date is in: 1=Jan, 2=Feb, ..., 12=Dec
     ' If TargetYear is omitted, year for current system date/time is used
     
     ' This function as written supports Excel.  To support Access, replace instances of
     ' CVErr(xlErrValue) with Null.  To use with other VBA-supported applications or with VB,
     ' substitute a similar value
     
    Dim FirstDate As Date
     
     ' Validate DayIndex
    If DayIndex < 1 Or DayIndex > 7 Then
        NthWeekday = CVErr(xlErrValue)
        Exit Function
    End If
     
    If TargetYear = 0 Then TargetYear = Year(Now)
     
    Select Case Position
         
         'Validate Position
    Case 1, 2, 3, 4, 5, "L", "l"
         
         ' Determine date for first of month
        FirstDate = DateSerial(TargetYear, TargetMonth, 1)
         
         ' Find first instance of our targeted weekday in the month
        If Weekday(FirstDate, vbSunday) < DayIndex Then
            FirstDate = FirstDate + (DayIndex - Weekday(FirstDate, vbSunday))
        ElseIf Weekday(FirstDate, vbSunday) > DayIndex Then
            FirstDate = FirstDate + (DayIndex + 7 - Weekday(FirstDate, vbSunday))
        End If
         
         ' Find the Nth instance.  If Position is not numeric, then it must be "L" for last.
         ' In that case, loop to find last instance of the month (could be the 4th or the 5th)
        If IsNumeric(Position) Then
            NthWeekday = FirstDate + (Position - 1) * 7
            If Month(NthWeekday) <> Month(FirstDate) Then NthWeekday = CVErr(xlErrValue)
        Else
            NthWeekday = FirstDate
            Do Until Month(NthWeekday) <> Month(NthWeekday + 7)
                NthWeekday = NthWeekday + 7
            Loop
        End If
         
         ' This only comes into play if the user supplied an invalid Position argument
    Case Else
        NthWeekday = CVErr(xlErrValue)
    End Select
     
End Function



