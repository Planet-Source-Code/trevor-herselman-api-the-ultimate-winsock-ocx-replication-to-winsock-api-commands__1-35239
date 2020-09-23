Attribute VB_Name = "modSox"
Option Explicit

Public SoxControl As Sox ' Our Public reference to Sox, This will allow us to call Sox commands from anywhere in the project like Sox.SendData instead of frmMain.Sox.SendData

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Let WindowProc = SoxControl.WndProc(hWnd, uMsg, wParam, lParam)
End Function
    
