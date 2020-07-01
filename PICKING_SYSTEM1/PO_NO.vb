Imports System.Linq
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
Public Class PO_NO
    Public myConn As SqlConnection
    Dim path As String
    Dim reader As SqlDataReader
    Dim x As ListViewItem
    Dim arr_qty_total As ArrayList = New ArrayList()
    Dim g_index As Integer = 0
    Private Sub PO_NO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            path = Me.GetType().Assembly.GetModules()(0).FullyQualifiedName
            Dim en As Int32 = path.LastIndexOf("\")
            path = path.Substring(0, en)
            path = Me.GetType().Assembly.GetModules()(0).FullyQualifiedName
            myConn = New SqlConnection("Data Source=192.168.161.101;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=pcs_admin;Password=P@ss!fa")
            'myConn = New SqlConnection("Data Source=192.168.43.42\SQLEXPRESS2017,1433;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=sa;Password=p@sswd;")
            myConn.Open()

        Catch ex As Exception
            MsgBox("Connect Database Fail" & vbNewLine & ex.Message, 16, "Status")
        Finally
            get_data_tetail()
        End Try

    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub
    Public Sub get_data_tetail()
        Try

            If Module1.M_CHECK_TYPE = "WEB_POST" Then
                Dim num As Integer = 0
                For Each key In Module1.arr_pick_detail_po
                    Dim lot As String = Module1.arr_pick_detail_lot(num).ToString
                    Dim po As String = Module1.arr_pick_detail_po(num).ToString
                    Dim QTY As String = Module1.arr_pick_detail_qty(num).ToString
                    x = New ListViewItem(Module1.arr_pick_detail_po(num).ToString)
                    x.SubItems.Add(Module1.arr_pick_detail_lot(num).ToString)
                    x.SubItems.Add(Module1.arr_pick_detail_qty(num).ToString)
                    x.SubItems.Add("0")
                    ListView2.Items.Add(x)
                    If Module1.check_pick_detail <> 0 Then
                        Dim count_scan As Integer = 0
                        For Each key2 In Module1.arr_check_PO_scan
                            If lot = Module1.arr_check_lot_scan(count_scan).ToString And po = Module1.arr_check_PO_scan(count_scan).ToString Then
                                Dim i As Integer = 0
                                Dim total_qty As Integer = 0
                                For Each key3 In Module1.arr_check_QTY_scan
                                    If lot = Module1.arr_check_lot_scan(i).ToString And po = Module1.arr_check_PO_scan(i).ToString Then
                                        total_qty = total_qty + Module1.arr_check_QTY_scan(i)
                                        If QTY <= total_qty Then
                                            ListView2.Items(num).BackColor = Color.FromArgb(103, 255, 103)
                                            ' MsgBox(num & " ," & count_scan & "===SET OK")
                                        ElseIf QTY > total_qty Then
                                            ListView2.Items(num).BackColor = Color.Yellow
                                            'MsgBox(num & " ," & count_scan & "===SETing")
                                        End If
                                    End If
                                    ListView2.Items(num).SubItems(3).Text = total_qty
                                    i = i + 1
                                Next

                            End If
                            count_scan = count_scan + 1
                        Next
                    End If
                    num = num + 1
                Next

            ElseIf M_CHECK_TYPE = "FW" Then
                Dim check As Integer = 0
                Dim num As Integer = 0
                For Each key In Module1.arr_pick_detail_lot
                    Dim lot As String = Module1.arr_pick_detail_lot(num).ToString
                    ' Dim po As String = Module1.arr_pick_detail_po(num).ToString
                    Dim QTY As String = Module1.arr_pick_detail_qty(num).ToString
                    x = New ListViewItem(" - ")
                    x.SubItems.Add(Module1.arr_pick_detail_lot(num).ToString)
                    x.SubItems.Add(Module1.arr_pick_detail_qty(num).ToString)
                    x.SubItems.Add("0")
                    ListView2.Items.Add(x)
                    If Module1.check_pick_detail <> 0 Then
                        Dim count_scan As Integer = 0
                        For Each key2 In Module1.arr_check_lot_scan
                            If lot = Module1.arr_check_lot_scan(count_scan).ToString Then
                                Dim i As Integer = 0
                                Dim total_qty As Integer = 0
                                For Each key3 In Module1.arr_check_QTY_scan
                                    If lot = Module1.arr_check_lot_scan(i).ToString Then
                                        total_qty = total_qty + Module1.arr_check_QTY_scan(i)
                                        If QTY <= total_qty Then
                                            ListView2.Items(num).BackColor = Color.FromArgb(103, 255, 103)
                                            Dim val = Module1.M_CHECK_LOT_COUNT_FW(num)
                                            ' If val = "2" Then
                                            'GoTo Exit_count2
                                            ' End If
                                            ' MsgBox(num & " ," & count_scan & "===SET OK")
                                        ElseIf QTY > total_qty Then
                                            ListView2.Items(num).BackColor = Color.Yellow
                                            Dim val = Module1.M_CHECK_LOT_COUNT_FW(num)

                                            ' If val = "2" Then
                                            'GoTo Exit_count2
                                            'End If

                                            'MsgBox(num & " ," & count_scan & "===SETing")
                                        End If
                                    End If
                                    ListView2.Items(num).SubItems(3).Text = total_qty
exit_loop:
                                    i = i + 1
                                Next

                            End If
Exit_count2:
                            count_scan = count_scan + 1
                        Next
                    End If
                    num = num + 1
                Next
            End If
        Catch ex As Exception
            MsgBox("ERROR LOT FAILL FROM CODE SUPPIER " & vbNewLine & ex.Message, 16, "Status in")
        End Try
      
    End Sub

    Private Sub ListView2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView2.SelectedIndexChanged

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
        part_detail.scan_qty.Focus()
    End Sub
    Public Function check_scan()

        Return 0
    End Function
    Public Function del_parts()
        Try
            Dim lot As String = "NOOOO"
            Dim PO As String = "NO"
            Dim QTY_REMOVE As String = "0"
            If IsNothing(Me.ListView2.FocusedItem) Then

            ElseIf ListView2.FocusedItem.Index >= 0 Then
                If ListView2.Items.Count > 0 Then
                    Dim index As Integer = ListView2.FocusedItem.Index
                    g_index = index
                    '   MsgBox(index)
                    MsgBox("g_index = " & g_index)
                End If
            End If

            If M_CHECK_TYPE = "FW" Then
                lot = Module1.arr_pick_detail_lot(g_index).ToString
                QTY_REMOVE = arr_pick_detail_qty(g_index).ToString
                Module1.arr_pick_detail_lot.Remove(g_index)
                Module1.arr_pick_detail_qty.Remove(g_index)

            ElseIf M_CHECK_TYPE = "WEB_POST" Then
                lot = Module1.arr_pick_detail_po(g_index).ToString
                QTY_REMOVE = arr_pick_detail_qty(g_index).ToString
                Module1.arr_pick_detail_po.Remove(g_index)
                Module1.arr_pick_detail_lot.Remove(g_index)
                Module1.arr_pick_detail_qty.Remove(g_index)
            End If
            Dim S_number As String = main.scan_terminal_id
            Dim strCommand2 As String = "delete from check_qr_part where item_cd = '" & Module1.past_numer & "' and scan_lot = '" & lot & "'  and S_number = '" & S_number & "'"

            'MsgBox(strCommand2)
            Dim command2 As SqlCommand = New SqlCommand(strCommand2, myConn)
            reader = command2.ExecuteReader()
            reader.Close()
            ListView2.Items(g_index).BackColor = Color.White
            part_detail.text_tmp.Text = part_detail.text_tmp.Text - QTY_REMOVE
            MsgBox("part_detail.text_tmp.Text = " & part_detail.text_tmp.Text)
            ' If part_detail.text_tmp.Text < "0" Or part_detail.text_tmp.Text < 0 Then
            'part_detail.text_tmp.Text = "0"
            ' End If
        Catch ex As Exception
            reader.Close()
            MsgBox("SORRY Delete ERROR Function del_parts" & vbNewLine & ex.Message, 16, "ALERT")
        End Try
        Return 0
    End Function

    ' Private Sub DEL_PARTS_SCAN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '   del_parts()
    '  End Sub
End Class