﻿Imports System.Runtime.InteropServices
Imports System.Data
'Imports System.Data.SqlServerCe
Imports System.Data.SqlClient
Imports System.Configuration
Module Module1
    Public check_query As Integer = 0
    Public user_id = "No data"
    Public hand_number = "No data"
    Public CODE_PD As String = "No data"
    Public data_combo As String = "No data"
    Public combo_pd As String = "No data"
    Public wi As String = "No data"
    Public line As String = "no data"
    Public tag_remain_qty As String = "no data"
    Public locations As String = "no data"
    Public past_numer As String = "no data"
    Public past_name As String = "no data"
    Public check_QTY As Integer = 0
    Public bool_check_scan As String = "no data"
    Public scan_qr_part_detail As String = "no data"
    Public ck As Integer = 0
    Public SCAN_QTY_TOTAL As Integer = 0
    Public check_count As Integer = 0
    Public check_count2 As Integer = 0
    Public QTY_OF_SCAN As Integer = 0
    Public M_QTY_STOP_SCAN As String = "0"
    Public M_WI_STOP_SCAN As String = "NNNNNNNOOOOOO"
    Public M_LINE_CD As String = "NNNNNNNOOOOOO"
    Public QTY As Integer = 0
    Public total_database As Integer = 0
    Public total_qty As Integer = 0
    Public select_pd As String = "no_data"
    Public time_scan As String = "no_data"
    Public Fullname As String = "no_data"
    Public M_Part_Selected As String = "no_data"
    Public M_LOT As String = "no_data"
    Public M_QTY_LOT_ALL As Integer = 0
    Public M_check_remain As String = "DATA_NULLLLLLL"
    Public M_Model As String = "data_null"
    Public M_SEQ_PRINT As String = "data_null"
    Public M_CHECK_TYPE As String = "data_null"
    Public M_update_stock As Integer = 0
    Public M_check_id As String = "NOOOOOOOOOO"
    Public M_CHECK_LOT_COUNT_FW As ArrayList = New ArrayList()
    Public check_pick_detail As Integer = 0
    Dim myConn As SqlConnection
    Dim reader As SqlDataReader
    Public user_reprint As String = "NODATA"
    Public M_reprint As String = "NODATA"
    Public arr_LVL As ArrayList = New ArrayList()
    Public arr_com_flg As ArrayList = New ArrayList()


    Public arr_check_qr_remain_lot As ArrayList = New ArrayList()
    Public arr_check_qr_remain_seq As ArrayList = New ArrayList()

    Public arr_check_qr_detail_remain_lot As ArrayList = New ArrayList()
    Public arr_check_qr_detail_remain_seq As ArrayList = New ArrayList()

    Public arr_pick_detail_po As ArrayList = New ArrayList()
    Public arr_pick_detail_qty As ArrayList = New ArrayList()
    Public arr_pick_detail_lot As ArrayList = New ArrayList()
    '----------------------check อันที่ scan ไปแล้ว-----------------------------'
    Public arr_check_lot_scan As ArrayList = New ArrayList()
    Public arr_check_PO_scan As ArrayList = New ArrayList()
    Public arr_check_QTY_scan As ArrayList = New ArrayList()
End Module
