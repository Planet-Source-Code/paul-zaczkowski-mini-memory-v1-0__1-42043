Attribute VB_Name = "modGetMemUsage"
' Copyright©2002-2004 CP_You Software
'

' All variables MUST be declared
Option Explicit

' Used to hold the percentage of usage
Dim Percent As Integer

' Used to hold the total amount of memory
Dim Total As Single

' Used to hold the available amount of memory
Dim Avail As Single

' Used to hold the Used amount of memory
Dim Used As Single

' Used to find the current memory status
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

' Used to find the current memory status
Public Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

' Declare MS as a MemoryStatus type
Dim MS As MEMORYSTATUS

Public Function GetMemUsage() As Integer

' Set the buffer size of ms
MS.dwLength = Len(MS)

' Get the current memory status and store it in MS
GlobalMemoryStatus MS

' Find the total amount of memory
Total = MS.dwTotalPhys

' Find the available amount of memory
Avail = MS.dwAvailPhys

' Calculate the usage
Used = (Total - Avail)

' Find the percent of usage
Percent = (Used * 100) / Total

' Return the percentage
GetMemUsage = Percent

End Function




