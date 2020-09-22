Option Explicit On
Option Strict On
Option Infer Off
Imports System
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
#Region "Win32Interop"
Public Class Win32Interop
    Public Const WM_USER As Int32 = 1024
    Public Const PBM_SETMARQUEE As Int32 = (WM_USER + 10)
    Public Const GWL_STYLE As Int32 = -16
    Public Const PBS_MARQUEE As Int32 = 8

    <DllImport("user32.dll", EntryPoint:="GetWindowLong")> _
    Public Shared Function GetWindowLong(ByVal hWnd As IntPtr, ByVal nIndex As Int32) As Int32
    End Function

    <DllImport("user32.dll", EntryPoint:="SetWindowLong")> _
    Public Shared Function SetWindowLong(ByVal hWnd As IntPtr, ByVal nIndex As Int32, ByVal dwNewLong As Int32) As Int32
    End Function

    <DllImport("User32", SetLastError:=True)> _
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Int32, ByVal wParam As Int32, ByVal lParam As Int32) As Int32
    End Function
End Class
#End Region
Public Class ProgressBarMoving
    Private _pb As Windows.Forms.ProgressBar
    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
        Application.EnableVisualStyles()
        Application.DoEvents()
        _pb = New Windows.Forms.ProgressBar
        _pb.Dock = Windows.Forms.DockStyle.Fill
        Width = 500
        Height = 50
        Controls.Add(_pb)
        Dim i As Int32 = Win32Interop.GetWindowLong(_pb.Handle, Win32Interop.GWL_STYLE)
        Win32Interop.SetWindowLong(_pb.Handle, Win32Interop.GWL_STYLE, i Or Win32Interop.PBS_MARQUEE)
        Win32Interop.SendMessage(_pb.Handle, Win32Interop.PBM_SETMARQUEE, 1, 100)
    End Sub
End Class
