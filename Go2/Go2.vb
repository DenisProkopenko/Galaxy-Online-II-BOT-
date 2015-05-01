
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.Drawing.Printing
Imports System.Text

Public Class Go2

    Dim StartStop As Integer = 0


    'Define some old windows API functions
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function GetWindowText(ByVal hwnd As IntPtr, ByVal lpString As StringBuilder, ByVal cch As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function GetClassName(ByVal hwnd As IntPtr, ByVal lpString As StringBuilder, ByVal cch As Integer) As Integer
    End Function

    Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll")> Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Int32) As IntPtr
    End Function

    <DllImport("user32.dll")> Shared Function PostMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Int32) As IntPtr
    End Function

    <DllImport("user32.dll")> Shared Function WindowFromPoint(ByVal pnt As Point) As IntPtr
    End Function

    Structure RECT
        Dim Left As Long
        Dim Top As Long
        Dim Right As Long
        Dim Bottom As Long
    End Structure

    Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

    Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
       ByVal nHeight As Long, ByVal bRepaint As Long) As Long

    Declare Function GetDesktopWindow Lib "user32" () As Long

    Declare Function GetDoubleClickTime Lib "user32" () As Long

    'The capture bitmaps
    Dim capture1 As System.Drawing.Bitmap
    Dim capture2 As System.Drawing.Bitmap

    Dim strWindowText As String = "none"
    Dim nRecapt As Integer = 0

    Dim StrGameState = "unknown"
    Dim strBotState = "nothing"

    Function MINOF(ByVal X As Integer, ByVal Y As Integer)
        If X < Y Then
            MINOF = X
        Else
            MINOF = Y
        End If
    End Function

    'Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
    'Start the game
    'LOAD
    'Dim H As Integer = 815
    'note igg html says the game is 1024 x 820.

    'Thread.Sleep(10000)

    ' Label7.Text = "Form3 Hwnd:" & Me.Handle.ToString
    ' Label8.Text = "WebBr Hwnd:" & Me.WebBrowser.Handle.ToString

    'End Sub

    Function WindowFromXY(ByVal x, ByVal y) As IntPtr
        Dim pnt As Point = New Point(x, y)
        Return WindowFromPoint(pnt)
    End Function

    Function bobsclick(ByVal x, ByVal y, ByVal Hwnd)
        Dim pnt As Point = New Point(x, y) ' Specify the location where you want to click.
        'Dim hWnd As IntPtr = WindowFromPoint(pnt)
        Label4.Text = "BobsClick hWnd = " & Hwnd.ToString
        If Hwnd <> IntPtr.Zero Then
            PostMessage(Hwnd, WM_LBUTTONDOWN, MK_LBUTTON, Covertxytolparam(x, y))
            Thread.Sleep(100)
            PostMessage(Hwnd, WM_LBUTTONUP, MK_LBUTTON, Covertxytolparam(x, y))
        End If

        Return 0
    End Function

    'Function bobsgetwindow()
    '    Dim pnt As Point = New Point(x, y) ' Specify the location where you want to click.
    '    Dim hWnd As IntPtr = WindowFromPoint(pnt)
    '    hWndBrowserWindow = hWnd
    'End Function

    Function Covertxytolparam(ByVal x As Integer, ByVal y As Integer) As Int32
        Return y << 16 Or x
    End Function



    Const BM_CLICK As Integer = &HF5
    Const WM_LBUTTONDOWN As Integer = &H201
    Const WM_LBUTTONUP As Integer = &H202
    Const MK_LBUTTON As Integer = &H1

    Private Function LoadABitMap(ByVal strFileName As String) As Bitmap
        ' Compose the picture's file name.
        'Dim file_name As String = "C:\go2\harvest.bmp"

        ' Load the picture into a Bitmap.
        Dim bm As New Bitmap(strFileName)

        ' Display the results.
        Return bm
    End Function

    Private Sub Go2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Enabled = True
        Timer1.Interval = 100
    End Sub


    Function AbetterWait(ByVal delaytime As Integer)
        Dim divisor = 100
        For X = 0 To delaytime / divisor
            Application.DoEvents()
            Thread.Sleep(divisor)
        Next
        Return 0
    End Function

    'Function BobsCount(ByVal delaytime As Integer)
    'Dim divisor = 100
    'For X = 0 To delaytime / divisor
    '  Application.DoEvents()
    '  Thread.Sleep(divisor)
    ' Next
    'Return 0
    'End Function

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim pnt As Point = New Point(0, 0)
        'ncount is number of retries before reloading
        Dim ncount As Integer = 100

        Dim Hwnd As IntPtr

        Hwnd = WindowFromXY(Cursor.Position.X, Cursor.Position.Y)

        lblHwind.Text = "HWND @(" & Cursor.Position.X & "," & Cursor.Position.Y & ") =" & Hwnd.ToString

    End Sub

    Function GetTheChildHandle(ParentHwnd As IntPtr) As IntPtr
        Dim List As IntPtr()
        Dim RetPtr As IntPtr = 0
        Dim strClass As String
        List = NativeMethods.GetChildWindows(ParentHwnd)
        Label9.Text = "Child Handles on the Form:" & vbCrLf
        If List.Length > 1 Then
            RetPtr = List(List.Length - 1)
            For X = 0 To List.Length - 1
                strClass = GetClass(List(X))
                If strClass = "MacromediaFlashPlayerActiveX" Then
                    Label10.Text = "Macromedia Found: " & List(X).ToString
                    RetPtr = List(X)
                End If
                Label9.Text = Label9.Text & CStr(X) & ":" & List(X).ToString & " Class= " & strClass & vbCrLf
            Next
        End If
        Return RetPtr
    End Function

    Public Function GetText(ByVal hWnd As IntPtr) As String
        Dim strTemp As String

        If hWnd.ToInt32 = 0 Then
            Return Nothing
        End If

        Dim sb As New System.Text.StringBuilder("", 255)

        GetWindowText(hWnd, sb, 255)
        strTemp = StripNulls(sb.ToString())
        Return strTemp
    End Function

    Public Function GetClass(ByVal hWnd As IntPtr) As String
        Dim strTemp As String

        If hWnd.ToInt32 = 0 Then
            Return Nothing
        End If

        Dim sb As New System.Text.StringBuilder("", 255)

        GetClassName(hWnd, sb, 255)
        strTemp = StripNulls(sb.ToString())
        Return strTemp
    End Function

    Public Function StripNulls(OriginalStr As String) As String
        ' This removes the extra Nulls so String comparisons will work
        If (InStr(OriginalStr, Chr(0)) > 0) Then
            OriginalStr = Mid(OriginalStr, 1, InStr(OriginalStr, Chr(0)) - 1)
        End If
        StripNulls = OriginalStr
    End Function

    Private Function dblclicktime()
        'This function gets the double click time for the mouse clicks.
        Dim s
        s = GetDoubleClickTime()
        'MsgBox(s)
        Return s
    End Function



    'Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
    'Click the login box
    'bobsclick(339, 284, WindowFromXY(339, 284))
    '   bobsclick(339, 284, GetTheChildHandle(Me.Handle))
    'End Sub
    ' Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
    'Click on the cancel after login
    'bobsclick(734, 550, WindowFromXY(734, 550))
    '    bobsclick(734, 550, GetTheChildHandle(Me.Handle))
    'End Sub

    Private Sub btnHarvest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHarvest.Click

        Warehouse()

    End Sub


    'loading
    'cancel after loading
    'warehouse
    'warehouse harvest
    'space base
    'space station
    'space station instances
    'instances
    'restricted
    'constellations
    'instance start
    'instance start 2
    'check mail
    'mail item filter
    'instance rewards mail
    'take all mail items
    'delete mail
    'military
    'military 2


    'play

    'loading
    Sub Loading()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\loading.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            Loading()
        Else
            LoadingCancel()
        End If
    End Sub
    'Cancel after loading
    Sub LoadingCancel()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\cancel.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\cancel2.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\cancel3.bmp")
                If pnt.X = 0 And pnt.Y = 0 Then
                    pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\cancel3.bmp")
                    If pnt.X = 0 And pnt.Y = 0 Then
                        MsgBox("Cancel after window load is not found", MsgBoxStyle.Exclamation, "Loading Cancel")
                    End If
                End If
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
            Warehouse()
        End If

    End Sub

    'warehouse
    Sub Warehouse()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\warehouse.bmp")
        lblDispay.Text = 1
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\warehouse2.bmp")
            lblDispay.Text = 2
            If pnt.X = 0 And pnt.Y = 0 Then
                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\warehouse3.bmp")
                lblDispay.Text = 3
                If pnt.X = 0 And pnt.Y = 0 Then
                    pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\warehouse4.bmp")
                    lblDispay.Text = 4
                    If pnt.X = 0 And pnt.Y = 0 Then
                        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\warehouse5.bmp")
                        lblDispay.Text = 5
                        If pnt.X = 0 And pnt.Y = 0 Then
                            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\warehouse6.bmp")
                            lblDispay.Text = 6
                            If pnt.X = 0 And pnt.Y = 0 Then
                                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\warehouse7.bmp")
                                lblDispay.Text = 7
                                If pnt.X = 0 And pnt.Y = 0 Then
                                    pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\warehouse8.bmp")
                                    lblDispay.Text = 8
                                    If pnt.X = 0 And pnt.Y = 0 Then
                                        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\warehouse9.bmp")
                                        lblDispay.Text = 9
                                        If pnt.X = 0 And pnt.Y = 0 Then
                                            MsgBox("Warehouse not found, please make sure warehouse is visable", MsgBoxStyle.Exclamation, "Warehouse not found")
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            lblWarehouse.Text = "(" & pnt.X & "," & pnt.Y & ")"
            Application.DoEvents()
            Thread.Sleep(100)
            Harvest()
        End If

    End Sub
    'warehouse harvest

    Sub Harvest()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\harvest.bmp")
        lblDisplay2.Text = 1
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\harvest2.bmp")
        lblDisplay2.Text = 2
        If pnt.X = 0 And pnt.Y = 0 Then
                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\harvest3.bmp")
        lblDisplay2.Text = 3
                If pnt.X = 0 And pnt.Y = 0 Then
                    MsgBox("Harvest not found", MsgBoxStyle.Exclamation, "Harvest not found")
                End If
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
            lblHarvest.Text = "(" & pnt.X & "," & pnt.Y & ")"
            SpaceBase()
        End If

    End Sub
    'ground base

    'space base button
    Sub SpaceBase()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\spacebase.bmp")
        lblSpace.Text = 1
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\spacebase2.bmp")
            lblSpace.Text = 2
            If pnt.X = 0 And pnt.Y = 0 Then
                MsgBox("Space Base button not visable... oops", MsgBoxStyle.Exclamation, "Space base button error")
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
            lblSpaceBase.Text = "(" & pnt.X & "," & pnt.Y & ")"
            'SpaceStation()
        End If

    End Sub

    'space station
    Sub SpaceStation()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\station.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\station2.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\station3.bmp")
                If pnt.X = 0 And pnt.Y = 0 Then
                    MsgBox("Space station not visable", MsgBoxStyle.Exclamation, "Space Station")
                End If
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub

    'space station instance
    Sub SpaceStationInstance()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instance.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instance3.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instance4.bmp")
                If pnt.X = 0 And pnt.Y = 0 Then
                    MsgBox("Space station instance not visable", MsgBoxStyle.Exclamation, "Space Station Instance")
                End If
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
            InstanceChoices()
            Thread.Sleep(100)
        End If

    End Sub

    'Instance choices
    Sub InstanceChoices()

        If radInstance30.Checked = True Then
            Instance30()
        ElseIf radInstance29.Checked = True Then
            Instance29()
        ElseIf radInstance28.Checked = True Then
            Instance28()
        ElseIf radInstance27.Checked = True Then
            Instance27()
        ElseIf radInstance26.Checked = True Then
            Instance26()
        ElseIf radInstance25.Checked = True Then
            Instance25()
        ElseIf radInstance24.Checked = True Then
            Instance24()
        ElseIf radInstance23.Checked = True Then
            Instance23()
        ElseIf radInstance22.Checked = True Then
            Instance22()
        ElseIf radInstance21.Checked = True Then
            Instance21()
        ElseIf radInstance20.Checked = True Then
            Instance20()
        ElseIf radInstance19.Checked = True Then
            Instance19()
        ElseIf radInstance18.Checked = True Then
            Instance18()
        ElseIf radInstance17.Checked = True Then
            Instance17()
        ElseIf radInstance16.Checked = True Then
            Instance16()
        ElseIf radInstance15.Checked = True Then
            Instance15()
        ElseIf radInstance14.Checked = True Then
            Instance14()
        ElseIf radInstance13.Checked = True Then
            Instance13()
        ElseIf radInstance12.Checked = True Then
            Instance12()
        ElseIf radInstance11.Checked = True Then
            Instance11()
        ElseIf radInstance10.Checked = True Then
            Instance10()
        ElseIf radInstance09.Checked = True Then
            Instance9()
        ElseIf radInstance08.Checked = True Then
            Instance8()
        ElseIf radInstance07.Checked = True Then
            Instance7()
        ElseIf radInstance06.Checked = True Then
            Instance6()
        ElseIf radInstance05.Checked = True Then
            Instance5()
        ElseIf radInstance04.Checked = True Then
            Instance4()
        ElseIf radInstance03.Checked = True Then
            Instance3()
        ElseIf radInstance02.Checked = True Then
            Instance2()
        ElseIf radInstance01.Checked = True Then
            Instance1()
        Else
            MsgBox("No instances are selected", MsgBoxStyle.Information, "Instance Options")
        End If

    End Sub

    'Instances
    'instance 1
    Sub Instance1()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i1i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 1 not found", MsgBoxStyle.Information, "Instance 1")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 2
    Sub Instance2()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i2i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 2 not found", MsgBoxStyle.Information, "Instance 2")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 3
    Sub Instance3()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i3i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 3 not found", MsgBoxStyle.Information, "Instance 3")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 4
    Sub Instance4()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i4i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 4 not found", MsgBoxStyle.Information, "Instance 4")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 5
    Sub Instance5()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i5i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 5 not found", MsgBoxStyle.Information, "Instance 5")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 6
    Sub Instance6()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i6i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 6 not found", MsgBoxStyle.Information, "Instance 6")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 7
    Sub Instance7()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i7i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 7 not found", MsgBoxStyle.Information, "Instance 7")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 8
    Sub Instance8()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i8i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 8 not found", MsgBoxStyle.Information, "Instance 8")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 9
    Sub Instance9()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i9i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 9 not found", MsgBoxStyle.Information, "Instance 9")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 10
    Sub Instance10()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i10i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 10 not found", MsgBoxStyle.Information, "Instance 10")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 11
    Sub Instance11()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i11i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 11 not found", MsgBoxStyle.Information, "Instance 11")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 12
    Sub Instance12()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i12i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 12 not found", MsgBoxStyle.Information, "Instance 12")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 13
    Sub Instance13()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i13i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 13 not found", MsgBoxStyle.Information, "Instance 13")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 14
    Sub Instance14()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i14i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 14 not found", MsgBoxStyle.Information, "Instance 14")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 15
    Sub Instance15()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i15i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 15 not found", MsgBoxStyle.Information, "Instance 15")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 16
    Sub Instance16()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i16i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 16 not found", MsgBoxStyle.Information, "Instance 16")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 17
    Sub Instance17()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i17i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 17 not found", MsgBoxStyle.Information, "Instance 17")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 18
    Sub Instance18()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i18i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i18i2.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                MsgBox("Instance 18 not found", MsgBoxStyle.Information, "Instance 18")
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 19
    Sub Instance19()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i19i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 19 not found", MsgBoxStyle.Information, "Instance 19")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 20
    Sub Instance20()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i20i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 20 not found", MsgBoxStyle.Information, "Instance 20")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 21
    Sub Instance21()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i21i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 21 not found", MsgBoxStyle.Information, "Instance 21")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 22
    Sub Instance22()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i22i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i22i2.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                MsgBox("Instance 22 not found", MsgBoxStyle.Information, "Instance 22")
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 23
    Sub Instance23()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i23i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 23 not found", MsgBoxStyle.Information, "Instance 23")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 24
    Sub Instance24()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i24i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 24 not found", MsgBoxStyle.Information, "Instance 24")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 25
    Sub Instance25()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i25i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 25 not found", MsgBoxStyle.Information, "Instance 25")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 26
    Sub Instance26()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i26i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 26 not found", MsgBoxStyle.Information, "Instance 26")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 27
    Sub Instance27()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i27i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 27 not found", MsgBoxStyle.Information, "Instance 27")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 28
    Sub Instance28()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i28i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 28 not found", MsgBoxStyle.Information, "Instance 28")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 29
    Sub Instance29()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i29i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 29 not found", MsgBoxStyle.Information, "Instance 29")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance 30
    Sub Instance30()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\i30i.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Instance 30 not found", MsgBoxStyle.Information, "Instance 30")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub

    'Restricted Instances
    'restricted 1
    Sub Restricted1()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r1r.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r1r1.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                MsgBox("Restricted 1 not found", MsgBoxStyle.Information, "Restricted 1")
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'restricted 2
    Sub Restricted2()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r2r.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r2r1.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                MsgBox("Restricted 2 not found", MsgBoxStyle.Information, "Restricted 2")
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'restricted 3
    Sub Restricted3()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r3r.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r3r1.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                MsgBox("Restricted 3 not found", MsgBoxStyle.Information, "Restricted 3")
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'restricted 4
    Sub Restricted4()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r4r.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r4r1.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                MsgBox("Restricted 4 not found", MsgBoxStyle.Information, "Restricted 4")
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'restricted 5
    Sub Restricted5()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r5r.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r5r1.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                MsgBox("Restricted 5 not found", MsgBoxStyle.Information, "Restricted 5")
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'restricted 6
    Sub Restricted6()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r6r.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r6r1.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                MsgBox("Restricted 6 not found", MsgBoxStyle.Information, "Restricted 6")
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'restricted 7
    Sub Restricted7()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r7r.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r7r1.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                MsgBox("Restricted 7 not found", MsgBoxStyle.Information, "Restricted 7")
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'restricted 8
    Sub Restricted8()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r8r.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r8r1.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                MsgBox("Restricted 8 not found", MsgBoxStyle.Information, "Restricted 8")
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'restricted 9
    Sub Restricted9()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r9r.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r9r1.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                MsgBox("Restricted 9 not found", MsgBoxStyle.Information, "Restricted 9")
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'restricted 10
    Sub Restricted10()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r10r.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\r10r1.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                MsgBox("Restricted 10 not found", MsgBoxStyle.Information, "Restricted 10")
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub

    'constallations
    Sub Constellations()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\constellations.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\constellations2.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\constellations3.bmp")
                If pnt.X = 0 And pnt.Y = 0 Then
                    MsgBox("Constellations tab not found", MsgBoxStyle.Information, "Constellations Tab")
                End If
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If
    End Sub
    'Aquarius
    Sub ConstAquarius()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\ConstAquarius.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Constellation Aquarius not found", MsgBoxStyle.Information, "Constellations Aquarius")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'Capricorn
    Sub ConstCapricorn()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\ConstCapricorn.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Constellation Capricorn not found", MsgBoxStyle.Information, "Constellations Capricorn")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'Leo
    Sub ConstLeo()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\ConstLeo.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\ConstLeo1.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                MsgBox("Constellation Leo not found", MsgBoxStyle.Information, "Constellations Leo")
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If
    End Sub
    'Libra
    Sub ConstLibra()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\ConstLibra.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Constellation Libra not found", MsgBoxStyle.Information, "Constellations Libra")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If
    End Sub
    'Picses
    Sub ConstPicses()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\ConstPicses.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Constellation Picses not found", MsgBoxStyle.Information, "Constellations Picses")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If
    End Sub
    'Sigitarius
    Sub ConstSigitarius()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\ConstSigitarius.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Constellation Sigitarius not found", MsgBoxStyle.Information, "Constellations Sigitarius")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If
    End Sub
    'Taurus
    Sub ConstTaurus()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\ConstTaurus.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Constellation Taurus not found", MsgBoxStyle.Information, "Constellations Taurus")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If
    End Sub
    'Virgo
    Sub ConstVirgo()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\ConstVirgo.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\ConstVirgo1.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                MsgBox("Constellation Virgo not found", MsgBoxStyle.Information, "Constellations Virgo")
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If
    End Sub

    'trials
    'trial 1-10
    'selecting instance
    Dim InstanceSelected As Integer = 0
    Private Function btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        If radInstance30.Checked = True Then
            InstanceSelected = 30
        ElseIf radInstance29.Checked = True Then
            InstanceSelected = 29
        ElseIf radInstance28.Checked = True Then
            InstanceSelected = 28
        ElseIf radInstance27.Checked = True Then
            InstanceSelected = 27
        ElseIf radInstance26.Checked = True Then
            InstanceSelected = 26
        ElseIf radInstance25.Checked = True Then
            InstanceSelected = 25
        ElseIf radInstance24.Checked = True Then
            InstanceSelected = 24
        ElseIf radInstance23.Checked = True Then
            InstanceSelected = 23
        ElseIf radInstance22.Checked = True Then
            InstanceSelected = 22
        ElseIf radInstance21.Checked = True Then
            InstanceSelected = 21
        ElseIf radInstance20.Checked = True Then
            InstanceSelected = 20
        ElseIf radInstance19.Checked = True Then
            InstanceSelected = 19
        ElseIf radInstance18.Checked = True Then
            InstanceSelected = 18
        ElseIf radInstance17.Checked = True Then
            InstanceSelected = 17
        ElseIf radInstance16.Checked = True Then
            InstanceSelected = 16
        ElseIf radInstance15.Checked = True Then
            InstanceSelected = 15
        ElseIf radInstance14.Checked = True Then
            InstanceSelected = 14
        ElseIf radInstance13.Checked = True Then
            InstanceSelected = 13
        ElseIf radInstance12.Checked = True Then
            InstanceSelected = 12
        ElseIf radInstance11.Checked = True Then
            InstanceSelected = 11
        ElseIf radInstance10.Checked = True Then
            InstanceSelected = 10
        ElseIf radInstance09.Checked = True Then
            InstanceSelected = 9
        ElseIf radInstance08.Checked = True Then
            InstanceSelected = 8
        ElseIf radInstance07.Checked = True Then
            InstanceSelected = 7
        ElseIf radInstance06.Checked = True Then
            InstanceSelected = 6
        ElseIf radInstance05.Checked = True Then
            InstanceSelected = 5
        ElseIf radInstance04.Checked = True Then
            InstanceSelected = 4
        ElseIf radInstance03.Checked = True Then
            InstanceSelected = 3
        ElseIf radInstance02.Checked = True Then
            InstanceSelected = 2
        ElseIf radInstance01.Checked = True Then
            InstanceSelected = 1
        Else
            MsgBox("No instances are selected", MsgBoxStyle.Information, "Instance Options")
        End If
        Return InstanceSelected
    End Function
    'selecting instance to run
    Sub InstanceNumber()
        If InstanceSelected = 30 Then
            Instance30()
        ElseIf InstanceSelected = 29 Then
            Instance29()
        ElseIf InstanceSelected = 28 Then
            Instance28()
        ElseIf InstanceSelected = 27 Then
            Instance27()
        ElseIf InstanceSelected = 26 Then
            Instance26()
        ElseIf InstanceSelected = 25 Then
            Instance25()
        ElseIf InstanceSelected = 24 Then
            Instance24()
        ElseIf InstanceSelected = 23 Then
            Instance23()
        ElseIf InstanceSelected = 22 Then
            Instance22()
        ElseIf InstanceSelected = 21 Then
            Instance21()
        ElseIf InstanceSelected = 20 Then
            Instance20()
        ElseIf InstanceSelected = 19 Then
            Instance19()
        ElseIf InstanceSelected = 18 Then
            Instance18()
        ElseIf InstanceSelected = 17 Then
            Instance17()
        ElseIf InstanceSelected = 16 Then
            Instance16()
        ElseIf InstanceSelected = 15 Then
            Instance15()
        ElseIf InstanceSelected = 14 Then
            Instance14()
        ElseIf InstanceSelected = 13 Then
            Instance13()
        ElseIf InstanceSelected = 12 Then
            Instance12()
        ElseIf InstanceSelected = 11 Then
            Instance11()
        ElseIf InstanceSelected = 10 Then
            Instance10()
        ElseIf InstanceSelected = 9 Then
            Instance9()
        ElseIf InstanceSelected = 8 Then
            Instance8()
        ElseIf InstanceSelected = 7 Then
            Instance7()
        ElseIf InstanceSelected = 6 Then
            Instance6()
        ElseIf InstanceSelected = 5 Then
            Instance5()
        ElseIf InstanceSelected = 4 Then
            Instance4()
        ElseIf InstanceSelected = 3 Then
            Instance3()
        ElseIf InstanceSelected = 2 Then
            Instance2()
        ElseIf InstanceSelected = 1 Then
            Instance1()
        End If
    End Sub
    'instance start
    'instance start
    Sub InstanceStart()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancestart3.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancestart4.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancestart5.bmp")
                If pnt.X = 0 And pnt.Y = 0 Then
                    pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancestart7.bmp")
                    If pnt.X = 0 And pnt.Y = 0 Then
                        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancestart9.bmp")
                        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancestart11.bmp")
                        If pnt.X = 0 And pnt.Y = 0 Then
                            MsgBox("Instance start button not visable", MsgBoxStyle.Exclamation, "Instance Start")
                        End If
                    End If
                End If
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance start 2
    Sub InstanceStart2()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancestart6.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancestart8.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancestart10.bmp")
                If pnt.X = 0 And pnt.Y = 0 Then
                    MsgBox("Instance start 2 button not visable", MsgBoxStyle.Exclamation, "Instance Start 2")
                End If
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub

    'mail icon
    'mail
    Sub CheckMail()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\mail.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            MsgBox("Mail not found", MsgBoxStyle.Information, "Mail")
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'mail Item filter
    Sub ItemMailFilter()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\mailitemfilter.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\mailitemfilter2.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\mailitemfilter3.bmp")
                If pnt.X = 0 And pnt.Y = 0 Then
                    MsgBox("Item mail filter not found", MsgBoxStyle.Information, "Item Mail Filter")
                End If
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'instance rewards mail
    Sub InstanceRewardsMail()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancereward.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancereward2.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancereward3.bmp")
                If pnt.X = 0 And pnt.Y = 0 Then
                    pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancereward4.bmp")
                    If pnt.X = 0 And pnt.Y = 0 Then
                        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancereward5.bmp")
                        If pnt.X = 0 And pnt.Y = 0 Then
                            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancereward6.bmp")
                            If pnt.X = 0 And pnt.Y = 0 Then
                                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancereward7.bmp")
                                If pnt.X = 0 And pnt.Y = 0 Then
                                    pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancereward8.bmp")
                                    If pnt.X = 0 And pnt.Y = 0 Then
                                        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancereward9.bmp")
                                        If pnt.X = 0 And pnt.Y = 0 Then
                                            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancereward10.bmp")
                                            If pnt.X = 0 And pnt.Y = 0 Then
                                                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancereward11.bmp")
                                                If pnt.X = 0 And pnt.Y = 0 Then
                                                    pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\instancereward12.bmp")
                                                    If pnt.X = 0 And pnt.Y = 0 Then
                                                        MsgBox("Mail instance rewards not found", MsgBoxStyle.Information, "Mail Instance Rewards")
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'take all items
    Sub TakeAllItems()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\takeitems.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\takeitems3.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\takeitems4.bmp")
                If pnt.X = 0 And pnt.Y = 0 Then
                    MsgBox("Take all items button not visable", MsgBoxStyle.Information, "Take All Mail Items")
                End If
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)

            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'delete mail
    Sub Mail()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\deletemail.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\deletemail2.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\deletemail3.bmp")
                If pnt.X = 0 And pnt.Y = 0 Then
                    MsgBox("Delete mail button not found", MsgBoxStyle.Information, "Delete Mail")
                End If
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub

    'military/ships
    'military
    Sub Military()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\military.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\military2.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\military7.bmp")
                If pnt.X = 0 And pnt.Y = 0 Then
                    MsgBox("Military not found", MsgBoxStyle.Information, "Military")
                End If
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub
    'military 2
    Sub Military2()
        Dim pnt As Point

        btnScreenGrab.PerformClick()
        Thread.Sleep(100)
        pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\military.bmp")
        If pnt.X = 0 And pnt.Y = 0 Then
            pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\military2.bmp")
            If pnt.X = 0 And pnt.Y = 0 Then
                pnt = GetCoordsOf("C:\Users\Denis\Desktop\Go2\Go2\Resources\military7.bmp")
                If pnt.X = 0 And pnt.Y = 0 Then
                    MsgBox("Military 2 not found", MsgBoxStyle.Information, "Military 2")
                End If
            End If
        Else
            bobsclick(pnt.X, pnt.Y, GetTheChildHandle(Me.Handle))
            Application.DoEvents()
            Thread.Sleep(100)
        End If

    End Sub


    ' Return a Bitmap holding an image of the control.
    Private Function GetControlImage(ByVal ctl As Control) As Bitmap
        Dim bm As New Bitmap(ctl.Width, ctl.Height)
        ctl.DrawToBitmap(bm, New Rectangle(0, 0, ctl.Width, ctl.Height))
        'ctl.CopyFromScreen(bm, New Rectangle(0, 0, ctl.Width, ctl.Height))

        Return bm
    End Function


    ' Return the form's image without its borders and
    ' decorations.
    'Private Function GetFormImageWithoutBorders(ByVal frm As  _
    '    System.Windows.Forms.WebBrowser) As Bitmap
    Private Function GetFormImageWithoutBorders(ByVal frm As Object) As Bitmap
        ' Get the form's whole image.
        Using whole_form As Bitmap = GetControlImage(frm)
            ' See how far the form's upper left corner is
            ' from the upper left corner of its client area.
            Dim origin As Point = frm.PointToScreen(New Point(0, 0))
            Dim dx As Integer = origin.X - frm.Left
            Dim dy As Integer = origin.Y - frm.Top

            ' Copy the client area into a new Bitmap.
            Dim wid As Integer = frm.ClientSize.Width
            Dim hgt As Integer = frm.ClientSize.Height
            Dim bm As New Bitmap(wid, hgt)
            Using gr As Graphics = Graphics.FromImage(bm)
                gr.DrawImage(whole_form, 0, 0, New Rectangle(dx, dy, wid, hgt), GraphicsUnit.Pixel)
            End Using
            Return bm
        End Using
    End Function

    Function GetCoordsOf(ByVal strFileName As String) As Point
        Dim bm As Bitmap
        Dim pnt As Point = New Point(0, 0)
        bm = LoadABitMap(strFileName)

        pnt = ImageCompare(bm)
        If pnt.Equals(New Point(0, 0)) Then
            Label12.Text = "no Match: "
        Else
            Label12.Text = strFileName & " image match = (" & pnt.ToString & ")"
        End If
        Return pnt
    End Function

    Function comparepixels(ByVal A As Color, ByVal B As Color)

        Dim M As Integer = 1
        Dim X As Integer
        Dim Y As Integer

        X = A.R
        Y = B.R
        If Math.Abs(X - Y) > M Then
            Return False
        End If
        X = A.G
        Y = B.G
        If Math.Abs(X - Y) > M Then
            Return False
        End If
        X = A.B
        Y = B.B
        If Math.Abs(X - Y) > M Then
            Return False
        End If
        Return True

    End Function

    Function ImageCompare(bm As Bitmap) As Point
        Dim base As Bitmap = PictureBox3.Image
        Dim pnt As Point = New Point(0, 0)
        Dim ncount As Integer = 0
        Label13.Text = ""
        'Find first pixle match
        For X = 0 To base.Width - 1
            For Y = 0 To base.Height - 1
                'If bm.GetPixel(0, 0) = base.GetPixel(X, Y) Then
                'If bm.GetPixel(0, 0).Equals(Color.FromArgb(bm.GetPixel(0, 0).A, base.GetPixel(X, Y).R, base.GetPixel(X, Y).G, base.GetPixel(X, Y).B)) Then
                If comparepixels(bm.GetPixel(0, 0), base.GetPixel(X, Y)) Then
                    If ncount = 0 Then
                        Label8.Text = "1st pnt@:(" & X & "," & Y & ")"
                    End If
                    If ImageCheck(bm, X, Y) Then
                        pnt = New Point(X + bm.Width / 2, Y + bm.Height / 2)
                        Return pnt
                    End If
                    ncount = ncount + 1

                End If
            Next
        Next
        Label7.Text = "# 1st point matches: " & ncount
        Return pnt
    End Function

    Function ImageCheck(bm As Bitmap, X As Integer, Y As Integer)
        Dim base As Bitmap = PictureBox3.Image
        For XX = 0 To bm.Width - 1
            For YY = 0 To bm.Height - 1
                'If bm.GetPixel(XX, YY) <> base.GetPixel(X + XX, Y + YY) Then
                'If bm.GetPixel(XX, YY).Equals(Color.FromArgb(bm.GetPixel(XX, YY).A, base.GetPixel(X + XX, Y + YY).R, base.GetPixel(X + XX, Y + YY).G, base.GetPixel(X + XX, Y + YY).B)) Then
                If Not comparepixels(bm.GetPixel(XX, YY), base.GetPixel(X + XX, Y + YY)) Then
                    'It doesn't match
                    Label13.Text = Label13.Text & X & "," & Y & bm.GetPixel(XX, YY).ToString & "::" & base.GetPixel(X + XX, Y + YY).ToString & vbCrLf
                    Return False
                End If
            Next
        Next
        'It does match
        Label11.Text = "Image match at: (" & CStr(X) & "," & CStr(Y) & ")"
        Return True
    End Function

    Function isbmblank()
        'This function checks to see if capture 1 is blank 
        Dim pixel As System.Drawing.Color
        Dim black As System.Drawing.Color = Color.FromArgb(0, 0, 0, 0)
        Dim white As System.Drawing.Color = Color.FromArgb(255, 255, 255, 255)
        Dim matchcount As Integer = 200 'If 200 pixels are black/white or the same as the one at 10,10 then it's a blank image.
        pixel = capture1.GetPixel(10, 10)
        For X = 10 To capture1.Width - 10 Step 5 'I'm only testing every 5th pixel to speed things up.
            For Y = 10 To capture1.Height - 10 Step 5
                If (capture1.GetPixel(X, Y).Equals(pixel)) Or (capture1.GetPixel(X, Y).Equals(black)) Or (capture1.GetPixel(X, Y).Equals(white)) Then
                    'It matches, so it might be a blank
                    matchcount = matchcount - 1
                    If matchcount = 0 Then
                        Label6.Text = "Blank"
                        Return True
                    End If
                Else
                    'It doesn't match so it's not blank
                    'MsgBox("Pix1: " & pixel.ToString & ", Pix2: " & capture1.GetPixel(X, Y).ToString, MsgBoxStyle.Information, "IsBmBlank")
                    Label6.Text = "Not Blank"
                    Return False
                End If
            Next
        Next

        Label6.Text = "Blank"
        Return True
    End Function

    Sub DoScreenGrab()

        'First try the alternate image capture routine
        PictureBox1.Image = GetFormImageWithoutBorders(Me)

        'Now try to capture the form (this usually works for the regular web stuff, once the game starts the other option works
        capture1 = CaptureForm.CreateScreenshot(TryCast(Me, Form), Me.Handle)
        PictureBox1.Image = capture1

        'Now capture the form via the flash player handle. This usually works after the flash player starts.
        capture2 = CaptureForm.CreateScreenshot(TryCast(Me, Form), GetTheChildHandle(Me.Handle))
        PictureBox3.Image = capture2

        If isbmblank() Then 'If capture 1 is blank use capture 2. 
            PictureBox2.Image = capture2
        Else
            PictureBox2.Image = capture1
        End If

        PictureBox2.Image = CaptureForm.CreateScreenshot(TryCast(Me, Form), Me.WebBrowser.Handle)

    End Sub

    Private Sub btnScreenGrab_Click(sender As Object, e As EventArgs) Handles btnScreenGrab.Click
        'capture button clicked 
        DoScreenGrab()
    End Sub

    Private Function btnStartStop_Click(sender As Object, e As EventArgs) Handles btnStartStop.Click
        StartStop = StartStop + 1

        If StartStop = 1 Then

        ElseIf StartStop = 2 Then

        Else
            StartStop = 0
        End If

        Return StartStop
    End Function

    'reset
    Private Sub btnClearAll_Click(sender As Object, e As EventArgs) Handles btnClearAll.Click
        'reset instance
        radInstance01.Checked = False
        radInstance02.Checked = False
        radInstance03.Checked = False
        radInstance04.Checked = False
        radInstance05.Checked = False
        radInstance06.Checked = False
        radInstance07.Checked = False
        radInstance08.Checked = False
        radInstance09.Checked = False
        radInstance10.Checked = False
        radInstance11.Checked = False
        radInstance12.Checked = False
        radInstance13.Checked = False
        radInstance14.Checked = False
        radInstance15.Checked = False
        radInstance16.Checked = False
        radInstance17.Checked = False
        radInstance18.Checked = False
        radInstance19.Checked = False
        radInstance20.Checked = False
        radInstance21.Checked = False
        radInstance22.Checked = False
        radInstance23.Checked = False
        radInstance24.Checked = False
        radInstance25.Checked = False
        radInstance26.Checked = False
        radInstance27.Checked = False
        radInstance28.Checked = False
        radInstance29.Checked = False
        radInstance30.Checked = False
        'reset trials
        radTrials.Checked = False
        'reset restricted
        radRestricted01.Checked = False
        radRestricted02.Checked = False
        radRestricted03.Checked = False
        radRestricted04.Checked = False
        radRestricted05.Checked = False
        radRestricted06.Checked = False
        radRestricted07.Checked = False
        radRestricted08.Checked = False
        radRestricted09.Checked = False
        radRestricted10.Checked = False
        'reset check boxes
        chkInstanceConstalations.Checked = False
        chkRestricted.Checked = False
        chkTrials.Checked = False

    End Sub

    'Tenth Second Timer
    Private Sub btnTenthSecondStart_Click(sender As Object, e As EventArgs) Handles btnTenthSecondStart.Click
        lblHalfSecondSeconds.Text = 0
        lblHalfSecondsMiliseconds.Text = 0
        tmrTenthSecond.Enabled = True
    End Sub

    Private Sub btnTenthSecondStop_Click(sender As Object, e As EventArgs) Handles btnTenthSecondStop.Click
        tmrTenthSecond.Enabled = False
    End Sub

    Private Sub btnTenthSecondReset_Click(sender As Object, e As EventArgs) Handles btnTenthSecondReset.Click
        tmrTenthSecond.Enabled = False
        lblHalfSecondsMiliseconds.Text = 0
        lblHalfSecondSeconds.Text = 0
    End Sub

    Private Sub tmrTenthSecond_Tick(sender As Object, e As EventArgs) Handles tmrTenthSecond.Tick
        lblHalfSecondsMiliseconds.Text = lblHalfSecondsMiliseconds.Text + 1
        If lblHalfSecondsMiliseconds.Text > 9 Then
            lblHalfSecondSeconds.Text = lblHalfSecondSeconds.Text + 1
            lblHalfSecondsMiliseconds.Text = 0
            If lblHalfSecondSeconds.Text = 1 Then
                tmrTenthSecond.Enabled = False
            End If
        End If
    End Sub

    'One Second Timer
    Private Sub btnOneSecondStart_Click(sender As Object, e As EventArgs) Handles btnOneSecondStart.Click
        lblOneSecondSeconds.Text = 0
        lblFiveSecondsMiliseconds.Text = 0
        tmrOneSecond.Enabled = True
        If lblOneSecondMiliseconds.Text > 9 Then
            tmrOneSecond.Enabled = False
        End If
    End Sub

    Private Sub btnOneSecondStop_Click(sender As Object, e As EventArgs) Handles btnOneSecondStop.Click
        tmrOneSecond.Enabled = False
    End Sub

    Private Sub btnOneSecondReset_Click(sender As Object, e As EventArgs) Handles btnOneSecondReset.Click
        lblOneSecondMiliseconds.Text = 0
        lblOneSecondSeconds.Text = 0
    End Sub

    Private Sub tmrOneSecond_Tick(sender As Object, e As EventArgs) Handles tmrOneSecond.Tick
        lblOneSecondMiliseconds.Text = lblOneSecondMiliseconds.Text + 1
        If lblOneSecondMiliseconds.Text > 9 Then
            lblOneSecondSeconds.Text = lblOneSecondSeconds.Text + 1
            lblOneSecondMiliseconds.Text = 0
            tmrOneSecond.Enabled = False
        End If
    End Sub

    'Five Seconds Timer
    Private Sub btnFiveSecondsStart_Click(sender As Object, e As EventArgs) Handles btnFiveSecondsStart.Click
        lblFiveSecondsSeconds.Text = 0
        lblFiveSecondsMiliseconds.Text = 0
        tmrFiveSeconds.Enabled = True
        If lblFiveSecondsSeconds.Text = 5 Then
            tmrFiveSeconds.Enabled = False

        End If
    End Sub

    Private Sub btnFiveSecondsStop_Click(sender As Object, e As EventArgs) Handles btnFiveSecondsStop.Click
        tmrFiveSeconds.Enabled = False
    End Sub

    Private Sub btnFiveSecondsReset_Click(sender As Object, e As EventArgs) Handles btnFiveSecondsReset.Click
        tmrFiveSeconds.Enabled = False
        lblFiveSecondsSeconds.Text = 0
        lblFiveSecondsMiliseconds.Text = 0
    End Sub

    Private Sub tmrFiveSeconds_Tick(sender As Object, e As EventArgs) Handles tmrFiveSeconds.Tick
        lblFiveSecondsMiliseconds.Text = lblFiveSecondsMiliseconds.Text + 1
        If lblFiveSecondsMiliseconds.Text > 9 Then
            lblFiveSecondsSeconds.Text = lblFiveSecondsSeconds.Text + 1
            lblFiveSecondsMiliseconds.Text = 0
            If lblFiveSecondsSeconds.Text = 5 Then
                tmrFiveSeconds.Enabled = False
            End If
        End If
    End Sub

    'Ten Seconds Timer
    Private Sub btnTenSecondsStart_Click(sender As Object, e As EventArgs) Handles btnTenSecondsStart.Click
        lblTenSecondsMiliseconds.Text = 0
        lblTenSecondsSeconds.Text = 0
        tmrTenSeconds.Enabled = True
        If lblTenSecondsSeconds.Text = 10 Then
            tmrTenSeconds.Enabled = False
        End If
    End Sub

    Private Sub btnTenSecondsStop_Click(sender As Object, e As EventArgs) Handles btnTenSecondsStop.Click
        tmrTenSeconds.Enabled = False
    End Sub

    Private Sub btnTenSecondsReset_Click(sender As Object, e As EventArgs) Handles btnTenSecondsReset.Click
        lblTenSecondsMiliseconds.Text = 0
        lblTenSecondsSeconds.Text = 0
        tmrTenSeconds.Enabled = False
    End Sub

    Private Sub tmrTenSeconds_Tick(sender As Object, e As EventArgs) Handles tmrTenSeconds.Tick
        lblTenSecondsMiliseconds.Text = lblTenSecondsMiliseconds.Text + 1
        If lblTenSecondsMiliseconds.Text > 9 Then
            lblTenSecondsSeconds.Text = lblTenSecondsSeconds.Text + 1
            lblTenSecondsMiliseconds.Text = 0
            If lblTenSecondsSeconds.Text = 10 Then
                tmrTenSeconds.Enabled = False
            End If
        End If
    End Sub

    'One Minute Timer
    Private Sub btnOneMinuteStart_Click(sender As Object, e As EventArgs) Handles btnOneMinuteStart.Click
        lblOneMinuteSeconds.Text = 0
        lblOneMinuteMinutes.Text = 0
        tmrOneMinute.Enabled = True
        If lblOneMinuteMinutes.Text = 1 Then
            tmrOneMinute.Enabled = False
        End If
    End Sub

    Private Sub btnOneMinuteStop_Click(sender As Object, e As EventArgs) Handles btnOneMinuteStop.Click
        tmrOneMinute.Enabled = False
    End Sub

    Private Sub btnOneMinuteReset_Click(sender As Object, e As EventArgs) Handles btnOneMinuteReset.Click
        lblOneMinuteSeconds.Text = 0
        lblOneMinuteMinutes.Text = 0
        tmrOneMinute.Enabled = False
    End Sub

    Private Sub tmrOneMinute_Tick(sender As Object, e As EventArgs) Handles tmrOneMinute.Tick
        lblOneMinuteSeconds.Text = lblOneMinuteSeconds.Text + 1
        If lblOneMinuteSeconds.Text > 59 Then
            lblOneMinuteMinutes.Text = lblOneMinuteMinutes.Text + 1
            lblOneMinuteSeconds.Text = 0
            If lblOneMinuteMinutes.Text = 1 Then
                tmrOneMinute.Enabled = False
            End If
        End If
    End Sub

    'Five Minutes
    Private Sub btnFiveMinutesStart_Click(sender As Object, e As EventArgs) Handles btnFiveMinutesStart.Click
        lblFiveMinutesSeconds.Text = 0
        lblFiveMinutesMinutes.Text = 0
        tmrFiveMinutes.Enabled = True
        If lblFiveMinutesMinutes.Text = 5 Then
            tmrFiveMinutes.Enabled = False
        End If
    End Sub

    Private Sub btnFiveMinutesStop_Click(sender As Object, e As EventArgs) Handles btnFiveMinutesStop.Click
        tmrFiveMinutes.Enabled = False
    End Sub

    Private Sub btnFiveMinutesReset_Click(sender As Object, e As EventArgs) Handles btnFiveMinutesReset.Click
        lblFiveMinutesSeconds.Text = 0
        lblFiveMinutesMinutes.Text = 0
        tmrFiveMinutes.Enabled = False
    End Sub

    Private Sub tmrFiveMinutes_Tick(sender As Object, e As EventArgs) Handles tmrFiveMinutes.Tick
        lblFiveMinutesSeconds.Text = lblFiveMinutesSeconds.Text + 1
        If lblFiveMinutesSeconds.Text > 59 Then
            lblFiveMinutesMinutes.Text = lblFiveMinutesMinutes.Text + 1
            lblFiveMinutesSeconds.Text = 0
            If lblFiveMinutesMinutes.Text = 5 Then
                tmrFiveMinutes.Enabled = False
            End If
        End If

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        btnTenthSecondStart_Click(e, Nothing)
        btnOneSecondStart_Click(e, Nothing)
        btnFiveSecondsStart_Click(e, Nothing)
        btnTenSecondsStart_Click(e, Nothing)
        btnOneMinuteStart_Click(e, Nothing)
        btnFiveMinutesStart_Click(e, Nothing)
    End Sub
End Class



Public Class NativeMethods
    <DllImport("User32.dll")> Private Shared Function EnumChildWindows(ByVal WindowHandle As IntPtr, ByVal Callback As EnumWindowProcess, ByVal lParam As IntPtr) As Boolean
    End Function

    Public Delegate Function EnumWindowProcess(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean

    Public Shared Function GetChildWindows(ByVal ParentHandle As IntPtr) As IntPtr()
        Dim ChildrenList As New List(Of IntPtr)
        Dim ListHandle As GCHandle = GCHandle.Alloc(ChildrenList)
        Try
            EnumChildWindows(ParentHandle, AddressOf EnumWindow, GCHandle.ToIntPtr(ListHandle))
        Finally
            If ListHandle.IsAllocated Then ListHandle.Free()
        End Try
        Return ChildrenList.ToArray
    End Function

    Private Shared Function EnumWindow(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean
        Dim ChildrenList As List(Of IntPtr) = GCHandle.FromIntPtr(Parameter).Target
        If ChildrenList Is Nothing Then Throw New Exception("GCHandle Target could not be cast as List(Of IntPtr)")
        ChildrenList.Add(Handle)
        Return True
    End Function

End Class

Public Class CaptureForm
    Private Declare Function BitBlt Lib "gdi32" Alias "BitBlt" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Integer) As Integer

    Private Declare Function GetDC Lib "user32" Alias "GetDC" (ByVal hwnd As Integer) As Integer

    Private Declare Function ReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hwnd As Integer, ByVal hdc As Integer) As Integer

    <DllImport("user32.dll")> Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Int32) As IntPtr
    End Function

    Public Shared Function CreateScreenshot(ByVal Capture, ByVal hWnd) As Bitmap
        Dim SRCCOPY = 13369376
        Dim WM_PRINT = &H317
        Dim WM_PRINTCLIENT = &H318
        Dim WM_PAINT = &HF
        Dim PRF_OWNED = &H20
        Dim PRF_CLIENT = &H4
        Dim PRF_CHILDREN = &H10

        Dim Rect As Rectangle
        Rect = New Rectangle(Capture.Location.X, Capture.Location.Y, Capture.Location.X + 1024, Capture.Location.Y + 1024)

        Dim gDest As Graphics
        Dim hdcDest As IntPtr
        Dim hdcSrc As Integer

        Dim screenBMP As New Bitmap(Rect.Right, Rect.Bottom)
        gDest = Graphics.FromImage(screenBMP)

        hdcSrc = GetDC(0)
        hdcDest = gDest.GetHdc

        SendMessage(hWnd, WM_PAINT, hdcDest, 0) ' PRF_CHILDREN Or PRF_OWNED Or PRF_CLIENT)
        SendMessage(hWnd, WM_PRINT, hdcDest, PRF_CHILDREN Or PRF_OWNED Or PRF_CLIENT)

        'Not the BitBlt is a screen grab and only works on visible windows on top of the desktop. It's here for any testing needs only. 
        '   Uncomment it to use it for testing.
        'BitBlt(hdcDest.ToInt32, 0, 0, Rect.Width, Rect.Height, hdcSrc, Capture.Location.X, Capture.Location.Y, SRCCOPY)

        gDest.ReleaseHdc(hdcDest)
        ReleaseDC(0, hdcSrc)
        Return screenBMP
    End Function
End Class

