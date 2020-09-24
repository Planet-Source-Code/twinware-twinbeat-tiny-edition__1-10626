Attribute VB_Name = "Module1"

'Declare function for timezone
Public Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation _
    As TIME_ZONE_INFORMATION) As Long
    
'Declare Systemtime (16 or 32 Bytes)

Public Type SYSTEMTIME ' 16 Bytes
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
    End Type


Public Type TIME_ZONE_INFORMATION  'That's the TimeZoneInformation...
    Bias As Long
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
    End Type


