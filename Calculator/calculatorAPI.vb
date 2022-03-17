Imports System.Net
Imports System.Text
Imports System.Data.OleDb
Imports System.Net.Mail
Imports System.IO
Imports System.Data.SqlClient
Imports System.Object


Public Class calculatorAPI


    Public Shared fp As Boolean = False
    Public Shared ref_number As String = ""
    Public Shared priceAPIResponse As String = ""

    Public Shared Function GetTotalPrice(ByVal order As String)
        Dim order_str As String

        Dim thick As String
        Dim material As Integer = 1  ' 3-ALU
        Dim charge_area As Double = 0
        Dim country As String
        Dim language As String = "de" 'de,uk,za,us,fr,it,nl
        Dim quantity As Integer = 0
        Dim layer_count As Integer = 2
        Dim soldermask As Integer = 1
        Dim silk As Integer = 0
        Dim silkscreen As Integer = 0
        Dim gap As Integer = 1
        Dim track_gap_size As Integer = 1
        Dim drill As Integer = 1
        Dim drill_diameter As Integer = 1
        Dim edge As Integer = 0
        Dim edge_metalization As Integer = 0
        Dim scoring As Integer = 0
        Dim cutout As Integer = 0
        Dim panel_processing As Integer = 3  ' 1-Scoring,2-Pips,3-SinglePieces,4-Cutout
        Dim lead_time As Integer = 6
        Dim layouts As Integer = 0
        Dim multi As Integer = 0
        Dim TheYear As String = CStr(Now.Year)
        Dim TheMonth As String = CStr(Now.Month)
        Dim au_number As String

        If Len(TheMonth) < 2 Then TheMonth = "0" & TheMonth

        Dim partslist As Object
        partslist = CreateObject("System.Collections.ArrayList")
        partslist.clear()

        Dim au_num As String

        Dim table As String
        table = "C:\in_house_files\local_database\faxback_info.dbf"

        Dim CommandText As String

        CommandText = "SELECT au_num,x_size,y_size,qty,silk,t_gaps_spec,holes_spec,ger,score,rout1 FROM " + table + " WHERE order = '" + order + "' order by time_stamp DESC"


        Dim ConnString As String = "Provider=VFPOLEDB.1;Data Source= " + table
        Dim Connection As New OleDbConnection(ConnString)
        Dim CommandResult As New OleDbCommand(CommandText, Connection)


        fp = False
        Try
            Connection.Open()
            Dim reader As OleDbDataReader = CommandResult.ExecuteReader(CommandBehavior.CloseConnection)
            While reader.Read()
                au_num = Trim(reader.GetString(0))
                If partslist.contains(au_num) Then GoTo done_already
                partslist.add(au_num)

                charge_area = charge_area + ((Trim(reader.GetString(1))) * (Trim(reader.GetString(2))) * (Trim(reader.GetString(3)))) / 10000
                If charge_area < CDbl(Trim(reader.GetString(3))) Then fp = True


                quantity = quantity + Trim(reader.GetString(3))

                If InStr(au_num, "-") > 0 Then
                    If (fp = True) Then multi = 2
                    layouts = layouts + 1
                End If


                silk = Trim(reader.GetString(4))
                If silk = 2 Then silk = 3
                If silk > silkscreen Then silkscreen = silk

                gap = Trim(reader.GetString(5)) + 1
                If gap > track_gap_size Then track_gap_size = gap

                drill = Trim(reader.GetString(6)) + 1
                If drill > drill_diameter Then drill_diameter = drill

                edge = Trim(reader.GetString(7))
                If edge > edge_metalization Then edge_metalization = edge

                scoring = Trim(reader.GetString(8))
                cutout = Trim(reader.GetString(9))
                panel_processing = 3
                If cutout > 0 Then panel_processing = 4
                If scoring > 0 Then panel_processing = 1

done_already:
            End While
            Connection.Close()
        Catch ex As Exception
            Connection.Close()
        End Try


        table = "C:\in_house_files\local_database\tiff_status.dbf"
        CommandText = "SELECT pool,thickness,work_days FROM " + table + " WHERE order_num = '" + order + "' order by time_stamp DESC"

        ConnString = "Provider=VFPOLEDB.1;Data Source= " + table
        Connection = New OleDbConnection(ConnString)
        CommandResult = New OleDbCommand(CommandText, Connection)

        Dim count As Integer = 0
        Try
            Connection.Open()
            Dim reader1 As OleDbDataReader = CommandResult.ExecuteReader(CommandBehavior.CloseConnection)
            While reader1.Read() And count < 1
                count = count + 1

                layer_count = Trim(reader1.GetString(0))
                Select Case layer_count
                    Case 1
                        layer_count = 2
                    Case 2
                        layer_count = 6
                    Case 3
                        layer_count = 2
                    Case 4
                        layer_count = 4
                    Case 5
                        layer_count = 6
                    Case 7
                        layer_count = 4
                    Case 8
                        layer_count = 1
                    Case 9
                        layer_count = 6
                    Case 10
                        layer_count = 2
                    Case 11
                        layer_count = 2
                    Case Else
                        layer_count = 2
                End Select


                thick = Trim(reader1.GetString(1))
                material = 1
                If InStr(thick, "1_50") > 0 Then material = 3

                lead_time = Trim(reader1.GetString(2))

            End While
            Connection.Close()
        Catch ex As Exception
            Connection.Close()
        End Try

        Dim panel As String

        table = "C:\in_house_files\local_database\orders_info.dbf"
        CommandText = "SELECT country,au_num,full_au,currency,Base_ref_number FROM " + table + " WHERE order_num = '" + order + "' order by time_stamp DESC"
        ConnString = "Provider=VFPOLEDB.1;Data Source= " + table
        Connection = New OleDbConnection(ConnString)
        CommandResult = New OleDbCommand(CommandText, Connection)

        partslist.clear()
        language = "de"
        au_number = "0"
        Try
            Connection.Open()
            Dim reader2 As OleDbDataReader = CommandResult.ExecuteReader(CommandBehavior.CloseConnection)
            While reader2.Read()
                au_number = Trim(reader2.GetString(2))
                ref_number = Trim(reader2.GetString(4))

                au_num = Trim(reader2.GetString(1))
                If partslist.contains(au_num) Then GoTo done_already2
                partslist.add(au_num)

                If InStr(au_num, "Z") Then
                    language = "za"
                    GoTo done_already2
                End If


                Dim currency As String
                country = Trim(reader2.GetString(0))
                currency = Trim(reader2.GetString(3))

                If InStr(au_num, "I") Then
                    If currency = "GBP" Then
                        language = "uk"
                    Else
                        language = "eu"
                    End If
                End If

                If InStr(country, "USA") > 0 Then language = "us"
                If InStr(country, "UNITED_STATE") > 0 Then language = "us"
                If InStr(country, "FRAN") > 0 Then language = "fr"
                If InStr(country, "ITA") > 0 Then language = "it"
                'If InStr(country, "UK") > 0 Then language = "uk"
                If InStr(country, "NETHER") > 0 Then language = "nl"

                If multi < 2 Then  ' Set to 1 if multipanel 
                    panel = Trim(reader2.GetString(3))
                    If CDbl(panel) > 1 Then multi = 1
                    panel = Trim(reader2.GetString(4))
                    If CDbl(panel) > 1 Then multi = 1
                    panel = Trim(reader2.GetString(5))
                    If CDbl(panel) > 0 Then multi = 1
                End If


done_already2:
            End While
            Connection.Close()
        Catch ex As Exception
            Connection.Close()
        End Try

        Dim precamNumberType, precamNumber As String

        'If InStr(au_num, "P") Then
        ' precamNumberType = "au_number"
        ' precamNumber = "AU-" & TheYear & TheMonth & "/" & order
        ' Else
        precamNumberType = "ref_number"
        precamNumber = ref_number
        '  End If

        order_str = "{  ""language"": """ & language & """, """ & precamNumberType & """: """ & precamNumber & """ , ""product"": { ""material_thickness"":" & material & ", ""layer_count"":" & layer_count & ", ""quantity"":" & quantity & ", ""surface"": 1, ""area"":" & charge_area & ", ""silkscreen"":" & silkscreen & ", ""track_gap_size"":" & track_gap_size & ", ""drill_diameter"":" & drill_diameter & ", ""edge_metallisation"":" & edge_metalization & ", ""add_layoutfiles"":" & layouts & ", ""multipanel"":" & multi & ", ""multipanel_processing"":" & panel_processing & ", ""lead_time"":" & lead_time & "  } } "

        Dim Jresult As String

        Try
            Jresult = JsonRequest(order_str)
        Catch ex As Exception
            MsgBox(ex.Message)
            Jresult = "0"
        End Try

        Return Jresult

    End Function


    Public Shared Function JsonRequest(ByVal order_str As String)
        Dim result As String

        Try
            Dim jsonString As String = order_str

            Dim Uri As New Uri(String.Format("https://api.beta-layout.com"))
            Dim data = Encoding.UTF8.GetBytes(jsonString)
            Dim header As New WebHeaderCollection
            header.Add("Api-Key", "48c97ca84593c93ca9043c9346e6ce00211e58bd11a3036bf9d9ed696b073939")
            result = SendRequest(Uri, data, "application/json", "POST", header)
            Return result
        Catch ex As Exception
            MsgBox(ex.Message, "Error")
            Return 0
        End Try

    End Function


    Public Shared Function SendRequest(ByVal uri As Uri, ByVal jsonDataBytes As Byte(), ByVal contentType As String, ByVal method As String, ByVal header As WebHeaderCollection) As String
        Try
            Dim req As WebRequest = WebRequest.Create(uri)
            req.Headers = header
            req.ContentType = contentType
            req.Method = method
            req.ContentLength = jsonDataBytes.Length


            Dim stream = req.GetRequestStream()
            stream.Write(jsonDataBytes, 0, jsonDataBytes.Length)
            stream.Close()

            Dim response = req.GetResponse().GetResponseStream()

            Dim reader As New StreamReader(response)
            Dim res = reader.ReadToEnd()
            reader.Close()
            response.Close()
            priceAPIResponse = res
            Return res
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return "0"
        End Try

    End Function



    Public Shared Function CalculatePrice(precamNumberType As String, precamNumber As String, language As String, material As String, layer_count As String, quantity As String, charge_area As String, silkscreen As String, track_gap_size As String, drill_diameter As String, edge_metalization As String, layouts As String, panel_processing As String, lead_time As String, multi As String)
        Dim order_str As String

        order_str = "{  ""language"": """ & language & """, """ & precamNumberType & """: """ & precamNumber & """ , ""product"": { ""material_thickness"":" & material & ", ""layer_count"":" & layer_count & ", ""quantity"":" & quantity & ", ""surface"": 1, ""area"":" & charge_area & ", ""silkscreen"":" & silkscreen & ", ""track_gap_size"":" & track_gap_size & ", ""drill_diameter"":" & drill_diameter & ", ""edge_metallisation"":" & edge_metalization & ", ""add_layoutfiles"":" & layouts & ", ""multipanel"":" & multi & ", ""multipanel_processing"":" & panel_processing & ", ""lead_time"":" & lead_time & "  } } "

        Dim Jresult As String
        Try
            Jresult = JsonRequest(order_str)
        Catch ex As Exception
            MsgBox(ex.Message)
            Jresult = "0"
        End Try
        Return Jresult
    End Function




End Class
