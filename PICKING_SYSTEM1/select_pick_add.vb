Imports System.Runtime.InteropServices
Imports System.Net
Imports System.Data
'Imports System.Data.SqlServerCe
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Windows.Forms.Form
Imports System
Public Class select_pick_add
    Public myConn = "NOO"
    Dim reader As SqlDataReader
    Private Sub select_pick_add_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'myConn = New SqlConnection("Data Source= 192.168.10.19\SQLEXPRESS2017,1433;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=sa;Password=p@sswd;")
            'myConn = New SqlConnection("Data Source=192.168.161.101;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=pcs_admin;Password=P@ss!fa")
            'myConn.Open()
            Dim connect_db = New connect()
            myConn = connect_db.conn()
        Finally
            TextBox2.Enabled = False
        End Try

    End Sub
    Private Sub Label2_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        main.ml = 0
        main.Timer1.Enabled = False
        main.Panel2.Visible = False
        Me.Close()
        main.Show()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker1.CustomFormat = "yyyy-MM-dd"
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        Dim get_date As String = DateTimePicker1.Format
        load_pd()
        load_line()
        load_wi()
        load_part()
        load_stock()
        TextBox1.Focus()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.TextChanged
        load_line()
        load_wi()
        load_part()
        load_stock()
        TextBox1.Focus()
    End Sub

    Private Sub Label8_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub DateTimePicker1_ValueChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Public Sub load_pd()
        ComboBox1.Items.Clear()
        Dim get_pd As String = "select sec_name,dep_id from sys_department order by dep_id"
        Dim command As SqlCommand = New SqlCommand(get_pd, myConn)
        reader = command.ExecuteReader()
        Do While reader.Read()
            ComboBox1.Items.Add(reader.Item(0))
        Loop
        reader.Close()
        ComboBox1.SelectedIndex = 0
    End Sub
    Public Sub load_line()
        ComboBox2.Items.Clear()
        Dim get_data_line As String = "SELECT DISTINCT LINE_CD from sup_work_plan_supply_dev where PD like '%" & ComboBox1.Text & "%' order by LINE_CD asc"
        Dim cmd_line As SqlCommand = New SqlCommand(get_data_line, myConn)
        reader = cmd_line.ExecuteReader()
        Dim num As Integer
        num = 0
        Dim status_line As Integer = 0
        Do While reader.Read()
            ComboBox2.Items.Add(reader.Item(0))
            num = num + 1
            status_line = 1
        Loop
        reader.Close()
        If status_line = 1 Then
            ComboBox2.SelectedIndex = 0
        End If
    End Sub
    Public Sub load_wi()
        ComboBox3.Items.Clear()
        Module1.date_now_database = DateTimePicker1.Format
        Dim get_wi As String = "SELECT AA.* FROM ( SELECT sw.PICK_QTY , sw.WORK_ODR_DLV_DATE AS d, sw.LVL AS LVL, sw.PICK_FLG AS PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2, pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE = '" & DateTimePicker1.Text & "' AND AA.LINE_CD = '" & ComboBox2.Text & "' ORDER BY AA.wi1 ASC" 'แบบ list ออกมา ในวัน พน ในการpick แต่จะมีปัญหาคือ ถ้าถึงวันศุกร์ แล้วต้องจัดแผนวันจัน จะมองไม่เห็นข้อมูล จะมองเห็นแค่ วันเสาร์'
        'Dim strCommand As String = "SELECT AA.* FROM ( SELECT sw.PICK_QTY, sw.WORK_ODR_DLV_DATE AS d, sw.LVL AS LVL, sw.PICK_FLG AS PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2, pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE BETWEEN '2020-07-07' AND '2020-07-07' AND AA.LINE_CD = '" & sel_where & "' ORDER BY AA.wi1 ASC"
        'MsgBox("get_wi = " & get_wi)
        Dim command_get_wi As SqlCommand = New SqlCommand(get_wi, myConn)
        reader = command_get_wi.ExecuteReader()
        Dim totala_scan_qty As Double = 0.0
        Dim Status_wi As String = 0
        Do While reader.Read = True
            If reader.Item(2) = "1" Then
                Status_wi = 1
                ComboBox3.Items.Add(reader.Item(6))
            End If
        Loop
        reader.Close()
        If Status_wi = 1 Then
            ComboBox3.SelectedIndex = 0
        End If
    End Sub
    Public Sub load_part()
        ComboBox4.Items.Clear()
        Module1.date_now_database = DateTimePicker1.Format
        Dim get_wi As String = "SELECT AA.* FROM ( SELECT sw.PICK_QTY , sw.WORK_ODR_DLV_DATE AS d, sw.LVL AS LVL, sw.PICK_FLG AS PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2, pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE = '" & DateTimePicker1.Text & "' AND AA.LINE_CD = '" & ComboBox2.Text & "' and AA.wi1 ='" & ComboBox3.Text & "'ORDER BY AA.wi1 ASC" 'แบบ list ออกมา ในวัน พน ในการpick แต่จะมีปัญหาคือ ถ้าถึงวันศุกร์ แล้วต้องจัดแผนวันจัน จะมองไม่เห็นข้อมูล จะมองเห็นแค่ วันเสาร์'
        'Dim strCommand As String = "SELECT AA.* FROM ( SELECT sw.PICK_QTY, sw.WORK_ODR_DLV_DATE AS d, sw.LVL AS LVL, sw.PICK_FLG AS PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2, pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE BETWEEN '2020-07-07' AND '2020-07-07' AND AA.LINE_CD = '" & sel_where & "' ORDER BY AA.wi1 ASC"
        'MsgBox("get_wi = " & get_wi)
        Dim command_get_wi As SqlCommand = New SqlCommand(get_wi, myConn)
        reader = command_get_wi.ExecuteReader()
        Dim totala_scan_qty As Double = 0.0
        Dim Status_part As String = 0
        Do While reader.Read = True
            If reader.Item(2) = "2" Then
                ComboBox4.Items.Add(reader.Item(4))
                Status_part = 1
            End If
        Loop
        reader.Close()
        If Status_part = 1 Then
            ComboBox4.SelectedIndex = 0
        End If
    End Sub
    Public Sub load_stock()
        Dim check_val As String = "SELECT count(id) as C_id from sup_frith_in_out where item_cd = '" & ComboBox4.Text & "' and com_flg <> 1"
        Dim c_check_val As SqlCommand = New SqlCommand(check_val, myConn)
        Dim C_val As Integer = 0
        reader = c_check_val.ExecuteReader()
        Do While reader.Read()
            C_val = reader("C_id").ToString
        Loop
        reader.Close()
        If C_val > 0 Then
            GET_DATA_WEB_POST()
        Else
            GET_DATA_FW()
        End If

    End Sub
    Private Sub ComboBox2_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        load_wi()
        load_part()
        load_stock()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        load_part()
        load_stock()
        TextBox1.Focus()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        load_stock()
        TextBox1.Focus()
    End Sub
    Public Function GET_DATA_WEB_POST()

        Dim strCommand As String = "select sum(qty) as stock from sup_frith_in_out where item_cd='" & ComboBox4.Text & "' and com_flg = '0'" 'ใช้ที่บ้านแบบใหม่
        Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
        reader = command.ExecuteReader()
        Dim Model As String = Nothing
        Dim totala_scan_qty As Double = 0.0
        Do While reader.Read()
            totala_scan_qty = CDbl(Val(reader("stock").ToString))
            ' Part_Selected = reader("ITEM_CD").ToString
            'Part_Name = reader("ITEM_NAME").ToString
            'lo = reader("LOCATION_PART").ToString
            QTY = totala_scan_qty
            ' wi = reader("WI").ToString
            'Lot_No = reader("LT").ToString
            ' Model = reader("MODEL").ToString
            'If Model <> Nothing Or Model <> "data_null" Then
            'Module1.M_Model = reader("MODEL").ToString
            'Else
            'Module1.M_Model = " - "
            'End If
            ' Module1.M_WI_STOP_SCAN = wi
            ' Module1.M_LINE_CD = code_line
            ' Module1.check_QTY = totala_scan_qty
            ' Module1.past_name = Part_Name
            'Module1.locations = lo
            ' Module1.M_LOT = reader("LT").ToString
        Loop
        reader.Close()
        TextBox2.Text = QTY
        Return 0
    End Function
    Public Sub GET_DATA_FW()
        Dim strCommand As String = "select sum(fa_total) -sum(fa_use) as stock from sup_frith_in_out_fa where fa_item_cd='" & ComboBox4.Text & "' and fa_com_flg = '0'"
        ' Dim strCommand As String = "SELECT item_cd, item_name , qty , location_part , wi ,MODEL , PICK_QTY  FROM sup_work_plan_supply_dev WHERE line_cd  = '" & code_line & "' AND item_cd = '" & code_part & "'AND wi  = '" & code_wi & "' AND (ps_unit_numerator <> '' AND location_part <> '') AND pick_flg != 1 AND WORK_ODR_DLV_DATE BETWEEN '2020-06-24' AND '2020-06-24' ORDER BY wi ASC"

        Dim Model As String = Nothing
        Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
        'MsgBox(strCommand)
        reader = command.ExecuteReader()
        Dim totala_scan_qty As Double = 0.0

        Do While reader.Read()
            totala_scan_qty = CDbl(Val(reader("stock").ToString))
            '(Part_Selected = reader("ITEM_CD").ToString)
            '(Part_Name = reader("ITEM_NAME").ToString)
            '(lo = reader("LOCATION_PART").ToString)
            QTY = totala_scan_qty
            'wi = reader("WI").ToString
            'Model = reader("MODEL").ToString
            Module1.check_QTY = QTY
            'Module1.past_name = Part_Name
            'Module1.locations = lo
            'Module1.M_LOT = "-"
            'If Model <> Nothing Or Model <> "data_null" Then
            ' Module1.M_Model = reader("MODEL").ToString
            ' Else
            ' Module1.M_Model = " - "
            'End If
            'Module1.M_WI_STOP_SCAN = wi
            'Module1.M_LINE_CD = code_line
        Loop
        ' MsgBox("LOT = " & Module1.M_LOT)
        reader.Close()
        TextBox2.Text = totala_scan_qty
    End Sub

End Class