﻿Imports System.Linq
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Net
Imports System.Net.Sockets
Imports System.IO
Imports Bt.CommLib
Imports Bt
Imports System.Data.SqlClient


Public Class scan_location
    Public Const INTERNET_DEFAULT_FTP_PORT As Int32 = 21
    Public Const INTERNET_OPEN_TYPE_PRECONFIG As Int32 = 0
    Public Const INTERNET_OPEN_TYPE_DIRECT As Int32 = 1
    Public Const INTERNET_OPEN_TYPE_PROXY As Int32 = 3
    Public Const INTERNET_INVALID_PORT_NUMBER As Int32 = 0
    Public Const INTERNET_SERVICE_FTP As Int32 = 1
    Public Const INTERNET_SERVICE_GOPHER As Int32 = 2
    Public Const INTERNET_SERVICE_HTTP As Int32 = 3
    Public Const FTP_TRANSFER_TYPE_BINARY As Int64 = &H2
    Public Const FTP_TRANSFER_TYPE_ASCII As Int64 = &H1
    Public Const INTERNET_FLAG_NO_CACHE_WRITE As Int64 = &H4000000
    Public Const INTERNET_FLAG_RELOAD As Int64 = &H80000000UI
    Public Const INTERNET_FLAG_KEEP_CONNECTION As Int64 = &H400000
    Public Const INTERNET_FLAG_MULTIPART As Int64 = &H200000
    Public Const INTERNET_FLAG_PASSIVE As Int64 = &H8000000
    Public Const FILE_ATTRIBUTE_READONLY As Int64 = &H1
    Public Const FILE_ATTRIBUTE_HIDDEN As Int64 = &H2
    Public Const FILE_ATTRIBUTE_SYSTEM As Int64 = &H4
    Public Const FILE_ATTRIBUTE_DIRECTORY As Int64 = &H10
    Public Const FILE_ATTRIBUTE_ARCHIVE As Int64 = &H20
    Public Const FILE_ATTRIBUTE_NORMAL As Int64 = &H80
    Public Const FILE_ATTRIBUTE_TEMPORARY As Int64 = &H100
    Public Const FILE_ATTRIBUTE_COMPRESSED As Int64 = &H800
    Public Const FILE_ATTRIBUTE_OFFLINE As Int64 = &H1000
    ' Constant to use with coredll.dll
    Public Const WAIT_OBJECT_0 As Int32 = &H0
    ' The constant that is used with printing data
    Public Const STX As [Byte] = &H2
    Public Const ETX As [Byte] = &H3
    Public Const DLE As [Byte] = &H10
    Public Const SYN As [Byte] = &H16
    Public Const ENQ As [Byte] = &H5
    Public Const ACK As [Byte] = &H6
    Public Const NAK As [Byte] = &H15
    Public Const ESC As [Byte] = &H1B
    Public Const LF As [Byte] = &HA

    Dim myConn As SqlConnection
    Public PD4 As Select_Line
    Dim path As String
    Dim imagefile As String
    Dim a As Boolean
    Public getwi As String = "wi_null"
    Dim reader As SqlDataReader
    Dim dat As String = String.Empty
    Dim check_scan As String = "NOOO"
    Private Sub scan_location_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            getwi = PD4.get_wi()
            'myConn = New SqlConnection("Data Source=192.168.43.42\SQLEXPRESS2017,1433;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=sa;Password=p@sswd;")
            myConn = New SqlConnection("Data Source=192.168.161.101;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=pcs_admin;Password=P@ss!fa")
            myConn.Open()
        Catch ex As Exception
            MsgBox("Connect Database Fail" & vbNewLine & ex.Message, 16, "Status scan")
        Finally

            Button1.Hide()
            PictureBox3.Visible = False
            path = Me.GetType().Assembly.GetModules()(0).FullyQualifiedName
            Dim en As Int32 = path.LastIndexOf("\")
            path = path.Substring(0, en)
            Me.text_box_location.Focus()
            lb_code_user.Text = main.show_code_id_user()
            lb_code_pd.Text = Module1.CODE_PD
            lb_code_line.Text = PD4.given_code_line()
            Part_Selected.Text = PD4.get_Part_Selected()
            Part_Name.Text = PD4.get_Part_Name()
            Location.Text = PD4.get_Location()
            get_data()
            Dim QTY_STOCK As Integer = Module1.M_QTY_LOT_ALL
            Dim QTY_System As Integer = Module1.check_QTY

            'MsgBox("QTY_STOCK = " & QTY_STOCK)
            'MsgBox("QTY_System = " & QTY_System)
            If QTY_STOCK < QTY_System Then
                STOCK_QTY.Text = "STOCK QTY : " & Module1.M_QTY_LOT_ALL & " ( Not enough ). "
            Else
                STOCK_QTY.Text = "STOCK QTY : " & Module1.M_QTY_LOT_ALL
            End If
            QTY.Text = PD4.get_QTY()
            Lot_No.Text = PD4.get_Lot_No
            Lot_No.Hide()
            ' alert()
            'MsgBox("LOAD SCAN LOCATION")
        End Try
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ok.Click
        Module1.M_CHECK_LOT_COUNT_FW = New ArrayList()
        Module1.M_QTY_LOT_ALL = 0
        Module1.arr_pick_detail_lot = New ArrayList()
        Module1.arr_pick_detail_po = New ArrayList()
        Module1.arr_pick_detail_qty = New ArrayList()
        Dim Line As Select_Line = New Select_Line()
        PD4.Show()
        Me.Close()
    End Sub

    Private Sub lb_code_user_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lb_code_user.ParentChanged

    End Sub


    Private Sub Label1_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lb_code_line.ParentChanged

    End Sub

    Private Sub lb_code_pd_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lb_code_pd.ParentChanged

    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Part_Selected_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Part_Selected.ParentChanged

    End Sub

    Private Sub QTY_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QTY.ParentChanged

    End Sub

    Private Sub text_box_location_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles text_box_location.TextChanged
m:
        System.Windows.Forms.Application.DoEvents()
        Try

            Dim sub_box = text_box_location.Text
            Dim loc_val As String = Location.Text.Substring(11)
            Dim ps = Part_Selected.Text.Substring(16)
            Module1.past_numer = ps
            Dim length_of_text_box As Integer
            length_of_text_box = (ps.Length() + loc_val.Length()) + 6
            'MsgBox(length_of_text_box)
            'MsgBox(sub_box.Length())
            Dim sub_box_length As Integer
            sub_box_length = sub_box.Length() + 6
            a = (length_of_text_box = sub_box_length)
            Dim val_box = sub_box.Split(" ")
            Dim part = val_box(0)
            Dim loca = val_box(6)
            Dim default_length As Integer = Len(ps & "      " & loc_val)
            Dim L_location As Integer = Len(text_box_location.Text)

            If default_length >= L_location Or default_length <= L_location Then

                If part = ps Then
                    If loca = loc_val Then
                        Button1.Show()
                        check_scan = "OK"
                    Else
                        text_box_location.Text = ""
                        Button1.Hide()
                        check_scan = "NO_OK"
                    End If
                Else
                    text_box_location.Text = ""
                    Button1.Hide()
                    check_scan = "NO_OK"
                End If

                If check_scan = "NO_OK" Then
                    fo.Focus()
                    PictureBox3.Visible = True
                    Button1.Visible = False
                    btn_ok.Visible = False
                    text_box_location.Visible = False
                End If
            End If
            'End If
        Catch ex As Exception
            ' MsgBox("scan faill")
        End Try


    End Sub
    Public Function check_alert()

        If check_scan = "NO_OK" Then
            Try
                alert()
            Catch ex As Exception
                MsgBox("Connect Database Fail" & vbNewLine & ex.Message, 16, "Status naja")
            End Try
            'GoTo M
        End If
        Return 0
    End Function
    Private Sub fo_key_down(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles fo.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Enter
                alert()
                PictureBox3.Visible = False
                Button1.Visible = False
                btn_ok.Visible = True
                text_box_location.Visible = True
                text_box_location.Focus()
        End Select



    End Sub


    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim SearchWithinThis As String = text_box_location.Text
        check()

    End Sub

    Private Sub Location_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Location.ParentChanged

    End Sub
    Public Sub show_fifo()
        'MsgBox("READY")
        Dim page_PO_NO As PO_NO = New PO_NO()
        page_PO_NO.Show()
    End Sub
    Public Sub check()
        ' MsgBox("In check")
        Dim part_de As part_detail = New part_detail()
        part_de.PD5 = Me
        part_de.Show()
        Me.Hide()
        show_fifo()
        'Dim show_image_part As show_image_part = New show_image_part()
        ' show_image_part.Show()
        'part_de.scan_qty.Focus()wq
        'show_image_part.Enabled = True
    End Sub

    Private Sub Panel1_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel1.GotFocus

    End Sub

    Private Sub Label1_ParentChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.ParentChanged

    End Sub

    Private Sub Part_Name_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Part_Name.ParentChanged

    End Sub

    Private Sub Lot_No_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Lot_No.ParentChanged

    End Sub

    Private Sub STOCK_QTY_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles STOCK_QTY.ParentChanged

    End Sub
    Public Sub get_data()
        Try
            Dim sum_strock_lot As String = "select sum(qty) as S_QTY from sup_frith_in_out where item_cd = '" & Module1.M_Part_Selected & "' and com_flg ='0' GROUP BY item_cd "
            'MsgBox("sum_strock_lot = " & sum_strock_lot)
            Dim command3 As SqlCommand = New SqlCommand(sum_strock_lot, myConn)
            reader = command3.ExecuteReader()
            Do While reader.Read()
                Module1.M_QTY_LOT_ALL = reader("S_QTY").ToString
            Loop
            reader.Close()
        Catch ex As Exception
            reader.Close()
        End Try
    End Sub

 
    Public Sub alert()
        Dim stBuz As New Bt.LibDef.BT_BUZZER_PARAM()
        Dim stVib As New Bt.LibDef.BT_VIBRATOR_PARAM()
        Dim stLed As New Bt.LibDef.BT_LED_PARAM()
        stBuz.dwOn = 200
        stBuz.dwOff = 100
        stBuz.dwCount = 2
        stBuz.bVolume = 3
        stBuz.bTone = 1
        stVib.dwOn = 200
        stVib.dwOff = 100
        stVib.dwCount = 2
        stLed.dwOn = 200
        stLed.dwOff = 100
        stLed.dwCount = 2
        stLed.bColor = Bt.LibDef.BT_LED_MAGENTA
        ' MsgBox("READY")
        Bt.SysLib.Device.btBuzzer(1, stBuz)
        Bt.SysLib.Device.btVibrator(1, stVib)
        Bt.SysLib.Device.btLED(1, stLed)
    End Sub


End Class