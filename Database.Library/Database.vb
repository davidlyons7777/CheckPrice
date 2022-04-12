Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.IO
'Imports Microsoft.Data.Odbc
'Imports System

Public Class Database

    Public Shared Sub CopyDatabaseToLocal()
        If Directory.Exists("C:\in_house_files\local_database") Then
            Try
                Directory.Delete("C:\in_house_files\local_database", True)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

        Try
            Directory.CreateDirectory("C:\in_house_files\local_database")
            File.Copy("t:\Database3\orders_info.dbf", "C:\in_house_files\local_database\orders_info.dbf")
            File.Copy("t:\Database3\faxback_info.dbf", "C:\in_house_files\local_database\faxback_info.dbf")
            File.Copy("t:\Database3\beta.dbc", "C:\in_house_files\local_database\beta.dbc")
            File.Copy("t:\Database3\beta.DCT", "C:\in_house_files\local_database\beta.DCT")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Shared Function GetInfoFromDatabase(strSelectColumn As String,
                                               strTableName As String,
                                               strWhereColumn As String,
                                               strEquals As String,
                                               strDatabasePath As String) As String
        Dim multifields As Boolean = False
        If strSelectColumn = "MULTI" Then multifields = True

        Dim strOutput As String = "0"
        Dim commandText As String

        If multifields = True Then
            strSelectColumn = " panel_s_q_x,panel_s_q_y,panel_edge "
        End If

        commandText = "SELECT " + strSelectColumn + " FROM " + strTableName + " WHERE " + strWhereColumn + " = '" + strEquals + "' order by time_stamp DESC"
        Dim ConnString As String = "Provider=VFPOLEDB.1;Data Source= " + strDatabasePath
        Dim Connection As New OleDbConnection(ConnString)
        Dim CommandResult As New OleDbCommand(commandText, Connection)

        Debug.WriteLine("START: " + commandText)
        Try
            Connection.Open()
            Dim reader As OleDbDataReader = CommandResult.ExecuteReader(CommandBehavior.CloseConnection)


            If multifields = False Then
                While reader.Read()
                    strOutput = Trim(reader.GetString(0))
                    Exit While
                End While
            Else
                While reader.Read()
                    strOutput = Trim(reader.GetString(0))
                    If CDbl(strOutput) > 1 Then
                        strOutput = "MULTI"
                        Exit While
                    End If

                    strOutput = Trim(reader.GetString(1))
                    If CDbl(strOutput) > 1 Then
                        strOutput = "MULTI"
                        Exit While
                    End If

                    strOutput = Trim(reader.GetString(2))
                    If CDbl(strOutput) > 0 Then
                        strOutput = "MULTI"
                        Exit While
                    End If

                    Exit While
                End While
            End If

            Connection.Close()
        Catch ex As Exception
            Connection.Close()
        End Try

        'Debug.WriteLine("END: " + commandText)

        Return strOutput

    End Function

    Public Shared Function GetActualArea(strAU As String) As String
        strAU = strAU.Trim
        Dim dblRunningArea As Double = 0.0
        Dim strFaxbackInfoDatabasePath As String = "C:\in_house_files\local_database\faxback_info.dbf"
        Dim strX, strY, strQty, strArea As String
        Dim boolZeroAreaFound As Boolean = False
        Dim intDashPartNumber As Integer = 0
        Dim arrLetterExtensions As New ArrayList From {"", "_E", "_F", "_G", "_H", "_I", "_J", "_K", "_L", "_M", "_N", "_O", "_P", "_Q", "_R", "_S", "_T", "_U", "_V", "_W", "_X", "_Y", "_Z"}
        Dim intLetterIndex As Integer = 0


        While boolZeroAreaFound = False
            Dim strDashNumber As String = ""
            For Each strExtension As String In arrLetterExtensions
                If intDashPartNumber > 0 Then
                    strDashNumber = "-" + intDashPartNumber.ToString
                End If
                strX = GetInfoFromDatabase("x_size", "faxback_info", "au_num", strAU + strDashNumber + strExtension, strFaxbackInfoDatabasePath)
                strY = GetInfoFromDatabase("y_size", "faxback_info", "au_num", strAU + strDashNumber + strExtension, strFaxbackInfoDatabasePath)
                strQty = GetInfoFromDatabase("qty", "faxback_info", "au_num", strAU + strDashNumber + strExtension, strFaxbackInfoDatabasePath)
                strArea = ((CDbl(strX) / 100) * (CDbl(strY) / 100)) * CDbl(strQty)
                If strArea = "0" And strExtension = "" Then
                    boolZeroAreaFound = True
                Else
                    dblRunningArea += CDbl(strArea)
                End If
            Next
            intDashPartNumber += 1
        End While

        Return dblRunningArea.ToString

    End Function

    Public Shared Function CheckIfDashOneExists(strAU As String) As Boolean
        Dim boolDashFound As Boolean
        Dim strQty As String
        Dim strOrdersInfoDatabasePath As String = "C:\in_house_files\local_database\orders_info.dbf"

        boolDashFound = False
        strQty = GetInfoFromDatabase("order_qty", "orders_info", "au_num", strAU + "-1", strOrdersInfoDatabasePath)

        If strQty <> "0" Then
            boolDashFound = True
        Else
            boolDashFound = False
        End If

        Return boolDashFound
    End Function

    Public Shared Function GetMhtArea(strAU As String) As String
        strAU = strAU.Trim
        Dim dblRunningArea As Double = 0.0
        Dim strOrdersInfoDatabasePath As String = "C:\in_house_files\local_database\orders_info.dbf"
        Dim strX, strY, strQty, strArea, strPanelEdge, strScoring, strQtyX, strQtyY As String
        Dim boolZeroAreaFound As Boolean = False
        Dim intDashPartNumber As Integer = 0
        Dim strFrameAdded = ""
        Dim strCountry As String = ""

        While boolZeroAreaFound = False
            Dim strAUWithDash As String = ""
            If intDashPartNumber > 0 Then
                strAUWithDash = "-" + intDashPartNumber.ToString
            End If

            strX = GetInfoFromDatabase("order_x", "orders_info", "au_num", strAU + strAUWithDash, strOrdersInfoDatabasePath)
            strY = GetInfoFromDatabase("order_y", "orders_info", "au_num", strAU + strAUWithDash, strOrdersInfoDatabasePath)
            strQty = GetInfoFromDatabase("order_qty", "orders_info", "au_num", strAU + strAUWithDash, strOrdersInfoDatabasePath)
            strCountry = GetInfoFromDatabase("country", "orders_info", "au_num", strAU + strAUWithDash, strOrdersInfoDatabasePath)
            If strCountry = "USA" Then
                strX = strX * 25.4
                strY = strY * 25.4
            End If

            strPanelEdge = GetInfoFromDatabase("panel_edge", "orders_info", "au_num", strAU + strAUWithDash, strOrdersInfoDatabasePath)
            strScoring = GetInfoFromDatabase("scoring", "orders_info", "au_num", strAU + strAUWithDash, strOrdersInfoDatabasePath)

            If strPanelEdge = "1" Then
                strQtyX = GetInfoFromDatabase("panel_s_q_x", "orders_info", "au_num", strAU + strAUWithDash, strOrdersInfoDatabasePath)
                strQtyY = GetInfoFromDatabase("panel_s_q_y", "orders_info", "au_num", strAU + strAUWithDash, strOrdersInfoDatabasePath)
                Dim strXPanel, strYPanel As String
                Dim intPanelCount As Integer
                Dim dblGap As Double
                If strScoring = "1" Then
                    dblGap = 0.0
                Else
                    dblGap = 2.0
                End If

                intPanelCount = strQty / (strQtyX * strQtyY)
                strXPanel = strQtyX * strX + (strQtyX - 1) * dblGap + 20
                strYPanel = strQtyY * strY + (strQtyY - 1) * dblGap + 20
                strArea = ((CDbl(strXPanel) / 100) * (CDbl(strYPanel) / 100)) * CDbl(intPanelCount)
                    strFrameAdded = " (inc Frame)"
                Else
                    strArea = ((CDbl(strX) / 100) * (CDbl(strY) / 100)) * CDbl(strQty)
            End If


            If strArea = "0" Then
                boolZeroAreaFound = True
            Else
                dblRunningArea += CDbl(strArea)
            End If
            intDashPartNumber += 1
        End While

        dblRunningArea = Math.Round(dblRunningArea, 4)

        Return dblRunningArea.ToString + strFrameAdded

    End Function
    Public Shared Function GetOrderRef(strAU As String) As String
        strAU = strAU.Trim
        Dim strOrdersInfoDatabasePath As String = "C:\in_house_files\local_database\orders_info.dbf"
        Dim strRef As String

        strRef = GetInfoFromDatabase("base_ref_number", "orders_info", "order_num", strAU, strOrdersInfoDatabasePath)


        Return strRef
    End Function

    Public Shared Function GetOrderThickness(strAU As String) As String
        strAU = strAU.Trim
        Dim strOrdersInfoDatabasePath As String = "C:\in_house_files\local_database\orders_info.dbf"
        Dim strThickness As String
        Dim result As String = "1"

        strThickness = GetInfoFromDatabase("thickness", "orders_info", "order_num", strAU, strOrdersInfoDatabasePath)

        If InStr(strThickness, "1_00") Then result = "2"

        Return result
    End Function

    Public Shared Function GetOrderSurface(strAU As String) As String
        strAU = strAU.Trim
        Dim strOrdersInfoDatabasePath As String = "C:\in_house_files\local_database\orders_info.dbf"
        Dim strSurface As String

        strSurface = GetInfoFromDatabase("surface", "orders_info", "order_num", strAU, strOrdersInfoDatabasePath)

        If strSurface = "3" Then strSurface = "4"

        Return strSurface
    End Function

    Public Shared Function GetActualCct(strAU As String) As String
        strAU = strAU.Trim
        Dim intRunningCct As Integer = 0
        Dim strFaxbackInfoDatabasePath As String = "C:\in_house_files\local_database\faxback_info.dbf"
        Dim strCct As String
        Dim boolZeroCctFound As Boolean = False
        Dim intDashPartNumber As Integer = 0

        While boolZeroCctFound = False
            Dim strAUWithDash As String = ""
            If intDashPartNumber > 0 Then
                strAUWithDash = "-" + intDashPartNumber.ToString
            End If
            strCct = GetInfoFromDatabase("qty", "faxback_info", "au_num", strAU + strAUWithDash, strFaxbackInfoDatabasePath)
            If strCct = "0" Then
                boolZeroCctFound = True
            Else
                intRunningCct += CInt(strCct)
            End If
            intDashPartNumber += 1
        End While

        Return intRunningCct.ToString
    End Function

    Public Shared Function GetMhtCct(strAU As String) As String
        strAU = strAU.Trim
        Dim intRunningCct As Integer = 0
        Dim strOrdersInfoDatabasePath As String = "C:\in_house_files\local_database\orders_info.dbf"
        Dim strQty, strCctTotal As String
        Dim boolZeroCctFound As Boolean = False
        Dim intDashPartNumber As Integer = 0

        While boolZeroCctFound = False
            Dim strAUWithDash As String = ""
            If intDashPartNumber > 0 Then
                strAUWithDash = "-" + intDashPartNumber.ToString
            End If

            strQty = GetInfoFromDatabase("order_qty", "orders_info", "au_num", strAU + strAUWithDash, strOrdersInfoDatabasePath)
            strCctTotal = strQty

            If strQty = "0" Then
                boolZeroCctFound = True
            Else
                intRunningCct += CInt(strCctTotal)
            End If
            intDashPartNumber += 1
        End While

        Return intRunningCct.ToString
    End Function



    Public Shared Function GetHiSpecGapCharge(strAU As String) As String
        Dim strFaxbackInfoDatabasePath As String = "C:\in_house_files\local_database\faxback_info.dbf"
        Dim result As String = GetInfoFromDatabase("t_gaps_spec", "faxback_info", "au_num", strAU, strFaxbackInfoDatabasePath)
        Return result
    End Function

    Public Shared Function GetEdgePlatingCharge(strAU As String) As String
        Dim strFaxbackInfoDatabasePath As String = "C:\in_house_files\local_database\faxback_info.dbf"
        Dim result As String = GetInfoFromDatabase("ger", "faxback_info", "au_num", strAU, strFaxbackInfoDatabasePath)
        Return result
    End Function

    Public Shared Function GetHiSpecDrillCharge(strAU As String) As String
        Dim strFaxbackInfoDatabasePath As String = "C:\in_house_files\local_database\faxback_info.dbf"
        Dim result As String = GetInfoFromDatabase("holes_spec", "faxback_info", "au_num", strAU, strFaxbackInfoDatabasePath)
        Return result
    End Function


    Public Shared Function GetScoringCharge(strAU As String) As String
        Dim strFaxbackInfoDatabasePath As String = "C:\in_house_files\local_database\faxback_info.dbf"
        Dim result As String = GetInfoFromDatabase("score", "faxback_info", "au_num", strAU, strFaxbackInfoDatabasePath)
        Return result
    End Function


    Public Shared Function GetCutoutCharge(strAU As String) As String
        Dim strFaxbackInfoDatabasePath As String = "C:\in_house_files\local_database\faxback_info.dbf"
        Dim result As String = GetInfoFromDatabase("rout1", "faxback_info", "au_num", strAU, strFaxbackInfoDatabasePath)
        Return result
    End Function


    Public Shared Function GetMaskCount(strAU As String) As String
        Dim strFaxbackInfoDatabasePath As String = "C:\in_house_files\local_database\faxback_info.dbf"
        Dim result As String = GetInfoFromDatabase("mask", "faxback_info", "au_num", strAU, strFaxbackInfoDatabasePath)
        Return result
    End Function


    Public Shared Function GetSilkCount(strAU As String) As String
        Dim strFaxbackInfoDatabasePath As String = "C:\in_house_files\local_database\faxback_info.dbf"
        Dim result As String = GetInfoFromDatabase("silk", "faxback_info", "au_num", strAU, strFaxbackInfoDatabasePath)
        Return result
    End Function

    Public Shared Function GetPanelEdge(strAU As String) As String
        Dim strDatabasePath As String = "C:\in_house_files\local_database\orders_info.dbf"
        Dim result As String = GetInfoFromDatabase("panel_edge", "orders_info", "au_num", strAU, strDatabasePath)
        Return result
    End Function


    Public Shared Function GetPool(strAU As String) As String
        Dim strDatabasePath As String = "C:\in_house_files\local_database\orders_info.dbf"
        Dim result As String = GetInfoFromDatabase("pool", "faxback_info", "au_num", strAU, strDatabasePath)
        Dim thickness As String = GetInfoFromDatabase("thickness", "orders_info", "au_num", strAU, strDatabasePath)

        If thickness = "1_50MM_MATERIAL" Then
            result = "ALU"
        End If
        Return result
    End Function

    Public Shared Sub UpdateDannedToTrue(strAU As String)
        strAU = strAU.Trim
        Dim commandText As String
        commandText = "UPDATE tiff_status SET danned=.T. WHERE order_num = '" + strAU + "'"
        Dim ConnString As String = "Provider=VFPOLEDB.1;Data Source=t:\database3\tiff_status.dbf"
        Dim Connection As New OleDbConnection(ConnString)
        Dim Command As New OleDbCommand(commandText, Connection)

        Try
            Connection.Open()
            Command.ExecuteNonQuery()
            Connection.Close()
        Catch ex As Exception
            Connection.Close()
        End Try
    End Sub

    Public Shared Sub UpdateDannedToTrueFaxBack(strAU As String)
        strAU = strAU.Trim
        Dim commandText As String
        commandText = "UPDATE faxback_info SET danned=.T. WHERE order = '" + strAU + "'"
        Dim ConnString As String = "Provider=VFPOLEDB.1;Data Source=t:\database3\faxback_info.dbf"
        Dim Connection As New OleDbConnection(ConnString)
        Dim Command As New OleDbCommand(commandText, Connection)

        Try
            Connection.Open()
            Command.ExecuteNonQuery()
            Connection.Close()
        Catch ex As Exception
            Connection.Close()
        End Try
    End Sub

    Public Shared Sub UpdatePriceToInOrdersInfo(strAU As String, strNewPrice As String)
        strAU = strAU.Trim
        Dim commandText As String
        commandText = "UPDATE orders_info SET order_price='" + strNewPrice + "' WHERE order_num = '" + strAU + "'"
        Dim ConnString As String = "Provider=VFPOLEDB.1;Data Source=t:\database3\orders_info.dbf"
        Dim Connection As New OleDbConnection(ConnString)
        Dim Command As New OleDbCommand(commandText, Connection)

        Try
            Connection.Open()
            Command.ExecuteNonQuery()
            Connection.Close()
        Catch ex As Exception
            Connection.Close()
        End Try



    End Sub

    Public Shared Sub UpdatePriceToInOrdersInfoSQL(strAU As String, strNewPrice As String)
        strAU = strAU.Trim

        Dim sqlCommand As String
        sqlCommand = "UPDATE orders_info SET order_price='" + strNewPrice + "' WHERE order_num = '" + strAU + "'"

        Try
            Dim id As Integer
            id = Shell("S:\Job\in_house_software\OMS_Dev_SQL_update " & Chr(34) & sqlCommand & Chr(34))
        Catch
        End Try

    End Sub


    Public Shared Function GetCountry(strAU As String) As String
        Dim strDatabasePath As String = "C:\in_house_files\local_database\orders_info.dbf"
        Dim result As String = GetInfoFromDatabase("country", "orders_info", "au_num", strAU, strDatabasePath)
        Return result
    End Function
End Class
