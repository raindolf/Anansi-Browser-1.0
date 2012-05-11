'Author Raindolf Owusu
'raindolf@oasiswebsoft.com and www.oasiswebsoft.com  


Option Explicit On
Option Strict On

Public Class iCam
#Region "Api/constants"

    Private Const WS_CHILD As Integer = &H40000000
    Private Const WS_VISIBLE As Integer = &H10000000
    Private Const SWP_NOMOVE As Short = &H2S
    Private Const SWP_NOZORDER As Short = &H4S
    Private Const WM_USER As Short = &H400S
    Private Const WM_CAP_DRIVER_CONNECT As Integer = WM_USER + 10
    Private Const WM_CAP_DRIVER_DISCONNECT As Integer = WM_USER + 11
    Private Const WM_CAP_SET_VIDEOFORMAT As Integer = WM_USER + 45
    Private Const WM_CAP_SET_PREVIEW As Integer = WM_USER + 50
    Private Const WM_CAP_SET_PREVIEWRATE As Integer = WM_USER + 52
    Private Const WM_CAP_GET_FRAME As Long = 1084
    Private Const WM_CAP_COPY As Long = 1054
    Private Const WM_CAP_START As Long = WM_USER
    Private Const WM_CAP_STOP As Long = (WM_CAP_START + 68)
    Private Const WM_CAP_SEQUENCE As Long = (WM_CAP_START + 62)
    Private Const WM_CAP_SET_SEQUENCE_SETUP As Long = (WM_CAP_START + 64)
    Private Const WM_CAP_FILE_SET_CAPTURE_FILEA As Long = (WM_CAP_START + 20)

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Short, ByVal lParam As String) As Integer
    Private Declare Function capCreateCaptureWindowA Lib "avicap32.dll" (ByVal lpszWindowName As String, ByVal dwStyle As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Short, ByVal hWndParent As Integer, ByVal nID As Integer) As Integer
    Private Declare Function capGetDriverDescriptionA Lib "avicap32.dll" (ByVal wDriver As Short, ByVal lpszName As String, ByVal cbName As Integer, ByVal lpszVer As String, ByVal cbVer As Integer) As Boolean
    Private Declare Function BitBlt Lib "GDI32.DLL" (ByVal hdcDest As IntPtr, ByVal nXDest As Integer, ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As Int32) As Boolean

#End Region

    Private iDevice As String
    Private hHwnd As Integer
    Private lwndC As Integer

    Public iRunning As Boolean

    Private CamFrameRate As Integer = 15
    Private OutputHeight As Integer = 240
    Private OutputWidth As Integer = 360

    Public Sub resetCam()
        'resets the camera after setting change
        If iRunning Then
            closeCam()
            Application.DoEvents()

            If setCam() = False Then
                MessageBox.Show("Errror Setting/Re-Setting Camera")
            End If
        End If

    End Sub

    Public Sub initCam(ByVal parentH As Integer)
        'Gets the handle and initiates camera setup
        If Me.iRunning = True Then
            MessageBox.Show("Camera Is Already Running")
            Exit Sub
        Else

            hHwnd = capCreateCaptureWindowA(iDevice, WS_VISIBLE Or WS_CHILD, 0, 0, OutputWidth, CShort(OutputHeight), parentH, 0)


            If setCam() = False Then
                MessageBox.Show("Error setting Up Camera")
            End If
        End If
    End Sub

    Public Sub setFrameRate(ByVal iRate As Long)
        'sets the frame rate of the camera
        CamFrameRate = CInt(1000 / iRate)

        resetCam()

    End Sub

    Private Function setCam() As Boolean
        'Sets all the camera up
        If SendMessage(hHwnd, WM_CAP_DRIVER_CONNECT, CShort(iDevice), CType(0, String)) = 1 Then
            SendMessage(hHwnd, WM_CAP_SET_PREVIEWRATE, CShort(CamFrameRate), CType(0, String))
            SendMessage(hHwnd, WM_CAP_SET_PREVIEW, 1, CType(0, String))
            Me.iRunning = True
            Return True
        Else
            Me.iRunning = False
            Return False
        End If
    End Function

    Public Function closeCam() As Boolean
        'Closes the camera
        If Me.iRunning Then
            closeCam = CBool(SendMessage(hHwnd, WM_CAP_DRIVER_DISCONNECT, 0, CType(0, String)))
            Me.iRunning = False
        End If
    End Function

    Public Function copyFrame(ByVal src As PictureBox, ByVal rect As RectangleF) As Bitmap
        If iRunning Then
            Dim srcPic As Graphics = src.CreateGraphics
            Dim srcBmp As New Bitmap(src.Width, src.Height, srcPic)
            Dim srcMem As Graphics = Graphics.FromImage(srcBmp)


            Dim HDC1 As IntPtr = srcPic.GetHdc
            Dim HDC2 As IntPtr = srcMem.GetHdc

            BitBlt(HDC2, 0, 0, CInt(rect.Width), _
              CInt(rect.Height), HDC1, CInt(rect.X), CInt(rect.Y), 13369376)

            copyFrame = CType(srcBmp.Clone(), Bitmap)

            'Clean Up 
            srcPic.ReleaseHdc(HDC1)
            srcMem.ReleaseHdc(HDC2)
            srcPic.Dispose()
            srcMem.Dispose()
        Else
            MessageBox.Show("Camera Is Not Running!")
        End If
    End Function

    Public Function FPS() As Integer
        Return CInt(1000 / (CamFrameRate))
    End Function

End Class
