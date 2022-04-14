Imports System.Data.OleDb
Imports System.IO
Imports Calculator
Imports Database.Library

Public Class frmMain
    Dim boolAllFieldsFilled As Boolean
    Dim strAuNumber As String
    Dim dblTotalArea As Double
    Dim intWhatBase As Integer
    Dim intCcts As Integer
    Dim intSilkCount As Integer
    Dim intWD As Integer
    Dim strPool As String
    Dim dblPricePerDm As Double

    Dim dblAreaPrice As Double
    Dim dblMaskPrice As Double
    Dim dblSilkPrice As Double
    Dim dblRoutPrice As Double
    Dim dblScorePrice As Double
    Dim dblHiSpecDrillPrice As Double
    Dim dblHiSpecTrackGapPrice As Double
    Dim dblEdgePlatingPrice As Double
    Dim dblExtraLayoutPrice As Double

    Dim dblAreaPriceAPI As Double
    Dim dblMaskPriceAPI As Double
    Dim dblSilkPriceAPI As Double
    Dim dblRoutPriceAPI As Double
    Dim dblScorePriceAPI As Double
    Dim dblHiSpecDrillPriceAPI As Double
    Dim dblHiSpecTrackGapPriceAPI As Double
    Dim dblEdgePlatingPriceAPI As Double
    Dim dblExtraLayoutPriceAPI As Double

    Dim intExtraLayoutCcts As Integer
    Dim strTotalWithVAT As Double
    Dim dblOriginalPrice As Double
    Dim dblActualPrice As Double
    Dim strOperName As String
    Dim strCountry As String
    Dim strCurrencySymbol As String
    Dim boolDashExists As Boolean
    Dim boolAUTextFileExists As Boolean

    Dim strAPICurrency As String

    Dim strJSONResult As String

    Dim strOriginalPrice As String
    'Dim orderRefNumber As String
    Dim orderThickness As String
    Dim orderSurface As String


    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetOperName()
        Database.Library.Database.CopyDatabaseToLocal()
        EnableNoteButton()
        AddHandlersToControls()
        LoadAUNumberFromTextFile()
        GetAPIPricing()
    End Sub

    Private Sub GetOperName()
        If File.Exists("C:\in_house_files\oper_name.txt") Then
            Dim fileContents As String
            fileContents = My.Computer.FileSystem.ReadAllText("C:\in_house_files\oper_name.txt")
            If fileContents = "" Then
                MsgBox("C:\in_house_files\oper_name.txt is empty. Please add text to file!")
                End
            End If
            strOperName = fileContents.Replace(vbCr, "")
            strOperName = strOperName.Replace(vbLf, "")
            strOperName = Trim(strOperName)
        Else
            MsgBox("Please create file C:\in_house_files\oper_name.txt")
            End
        End If
    End Sub

    Private Sub EnableNoteButton()
        If File.Exists("T:\Database3\NOTES\" + strAuNumber + "_note.txt") Then
            btnOpenNote.Enabled = True
        Else
            btnOpenNote.Enabled = False
        End If
    End Sub

    Private Sub LoadAUNumberFromTextFile()
        If File.Exists("C:\in_house_files\price_check_au.txt") Then
            boolAUTextFileExists = True
            Dim auNum As String
            auNum = My.Computer.FileSystem.ReadAllText("C:\in_house_files\price_check_au.txt")
            auNum = auNum.Trim
            txtAuNumber.Text = auNum
            If txtAuNumber.Text <> "" Then
                ImportDataFromDatabase()
            End If
        Else
            boolAUTextFileExists = False
        End If
    End Sub

    Private Sub AddHandlersToControls()
        AddHandler txtTotalArea.TextChanged, AddressOf dataChanged
        AddHandler chkInches.CheckedChanged, AddressOf dataChanged
        AddHandler cmbWD.TextChanged, AddressOf dataChanged
        AddHandler cmbLanguage.TextChanged, AddressOf dataChanged
        AddHandler txtCctCount.TextChanged, AddressOf dataChanged
        AddHandler txtVAT.TextChanged, AddressOf dataChanged
        AddHandler txtExtraLayouts.TextChanged, AddressOf dataChanged
        AddHandler chkSoldermask.CheckedChanged, AddressOf dataChanged
        AddHandler chkCutout.CheckedChanged, AddressOf dataChanged
        AddHandler chkScoring.CheckedChanged, AddressOf dataChanged
        AddHandler chkHiSpecDrills.CheckedChanged, AddressOf dataChanged
        AddHandler chkHiSpecGaps.CheckedChanged, AddressOf dataChanged
        AddHandler chkEdgePlating.CheckedChanged, AddressOf dataChanged
        AddHandler chkDiscount.CheckedChanged, AddressOf dataChanged

        Dim panels = Me.Controls.OfType(Of Panel)()
        For Each panel In panels
            Dim rdos = panel.Controls.OfType(Of RadioButton)()
            For Each rdo In rdos
                AddHandler rdo.CheckedChanged, AddressOf dataChanged
            Next
        Next
    End Sub


    Private Sub btnCheckPrice_Click(sender As Object, e As EventArgs) Handles btnCheckPrice.Click
        GetAPIPricing()
    End Sub

    Private Sub GetAPIPricing()
        strAuNumber = Trim(txtAuNumber.Text)
        CalculatePriceFromUserInputAndAPI()
        btnCheckPrice.Text = "Price Up To Date"
        btnCheckPrice.BackColor = Color.LightGreen
        btnCheckPrice.Enabled = False
        lblTotalPrice.Text = "actual: " + Replace(lblTotalAPI.Text, "Total:", "")
        UpdateCctLabel()
        UpdateAreaLabel()
        UpdateBaseLabel()
        UpdateDifferenceLabel()
        EnableNoteButton()
    End Sub

    Private Sub UpdatePrice()
        boolAllFieldsFilled = CheckInputs()
        If boolAllFieldsFilled Then
            CheckInputs()
            SetVariables()
            GetInfoFromDatabase()
            UpdatePriceLabels()
            UpdateDifferenceLabel()
            UpdateCctLabel()
            UpdateAreaLabel()
            UpdateBaseLabel()
            btnCheckPrice.Enabled = False
            EnableNoteButton()
        End If
    End Sub

    Private Sub UpdateBaseLabel()
        Dim intBase As Integer = Math.Truncate(CDbl(txtTotalArea.Text))
        If intBase < 1 Then
            intBase = 1
        End If
        If intBase > 50 Then
            intBase = 50
        End If
        lblBaseActual.Text = "actual: " + intBase.ToString
    End Sub

    Private Sub UpdateAreaLabel()
        lblTotalActualArea.Text = "actual: " + txtTotalArea.Text
    End Sub

    Private Sub UpdateCctLabel()
        lblActualCct.Text = "actual: " + txtCctCount.Text
    End Sub

    Private Sub UpdateDifferenceLabel()
        Dim AllowedDif As Double = (dblOriginalPrice / 100) * 1
        If AllowedDif > 10 Then AllowedDif = 10
        If AllowedDif < 1 Then AllowedDif = 1

        Dim OutOfToll As Boolean = False
        If Math.Abs(dblActualPrice - dblOriginalPrice) > AllowedDif Then OutOfToll = True


        Dim dblDiff As Double = dblActualPrice - dblOriginalPrice
        Dim dblDiffPercent As Double = (dblDiff / dblOriginalPrice) * 100
        dblDiffPercent = Math.Round(dblDiffPercent, 2)
        If dblOriginalPrice = 0.0 Then
            dblDiffPercent = 100
        End If
        lblPriceDiff.Text = "Price diff: " + strCurrencySymbol + Math.Round(dblDiff, 2).ToString + " (" + dblDiffPercent.ToString + "%)"

        'If dblDiffPercent > 3.0 Or dblDiffPercent < -3.0 Then
        If OutOfToll = True Then
            lblPriceDiff.ForeColor = Color.Red
        Else
            lblPriceDiff.ForeColor = Color.Green
        End If
    End Sub

    Private Sub GetInfoFromDatabase()

        CalculateAreaPrice()
        CalculateSilkPrice()
        CalculateMaskPrice()
        CalculateRoutPrice()
        CalculateScorePrice()
        CalculateHiSpecPrices()
        CalculateEdgePlatingPrice()
        CalculateExtraLayoutPrice()

    End Sub

    Private Sub GetWD()
        Dim strWD As String = Database.Library.Database.GetInfoFromDatabase("work_days",
                                                                            "orders_info",
                                                                            "au_num",
                                                                            txtAuNumber.Text,
                                                                            "T:\Database3\orders_info.dbf"
                                                                            )
        cmbWD.Text = strWD
    End Sub
    Private Sub SetMulti()
        Dim strMulti As String = Database.Library.Database.GetInfoFromDatabase("MULTI",
                                                                            "orders_info",
                                                                            "au_num",
                                                                            txtAuNumber.Text,
                                                                            "T:\Database3\orders_info.dbf"
                                                                            )
        If strMulti = "MULTI" Then
            rdoMultiMixed.Checked = True
        Else
            rdoMultiSingle.Checked = True
        End If

    End Sub
    Private Sub GetOriginalPrice()
        strOriginalPrice = Database.Library.Database.GetInfoFromDatabase("order_price",
                                                                                       "orders_info",
                                                                                       "au_num",
                                                                                       txtAuNumber.Text,
                                                                                       "T:\Database3\orders_info.dbf"
                                                                                       )
        lblOriginalOrderPrice.Text = "mht:    " + strCurrencySymbol + strOriginalPrice
        dblOriginalPrice = CDbl(strOriginalPrice)
    End Sub

    Private Function CheckAuFieldIsFilled() As Boolean
        Dim strAuNumberTemp As String = txtAuNumber.Text
        If strAuNumberTemp = "" Then
            MsgBox("Please Enter AU Number")
            Return True
        End If
        Return False
    End Function

    Private Function CheckInputs()
        If txtCctCount.Text = "" Or txtExtraLayouts.Text = "" Or txtTotalArea.Text = "" Or cmbWD.Text = "" Then
            MsgBox("All text fields must be filled to get new price!")
            Return False
        End If
        Return True
    End Function

    Private Sub CalculateExtraLayoutPrice()
        dblExtraLayoutPrice = Calculator.Calculator.GetExtraLayoutPrice(intExtraLayoutCcts, "GERMANY")
    End Sub

    Private Sub CalculateEdgePlatingPrice()
        dblEdgePlatingPrice = Calculator.Calculator.GetEdgePlatingPrice(intCcts, chkEdgePlating.Checked, strCountry)
    End Sub

    Private Sub CalculateHiSpecPrices()
        dblHiSpecDrillPrice = Calculator.Calculator.GetHiSpecPrice(dblTotalArea, dblPricePerDm, chkHiSpecDrills.Checked)
        dblHiSpecTrackGapPrice = Calculator.Calculator.GetHiSpecPrice(dblTotalArea, dblPricePerDm, chkHiSpecGaps.Checked)
    End Sub

    Private Sub CalculateScorePrice()
        dblScorePrice = Calculator.Calculator.GetScorePrice(intCcts, chkScoring.Checked, strCountry)
    End Sub

    Private Sub CalculateRoutPrice()
        boolDashExists = Database.Library.Database.CheckIfDashOneExists(txtAuNumber.Text)
        dblRoutPrice = Calculator.Calculator.GetRoutPrice(intCcts, chkCutout.Checked, dblTotalArea, strCountry, boolDashExists)
    End Sub

    Private Sub CalculateSilkPrice()
        dblSilkPrice = Calculator.Calculator.GetSilkPrice(intCcts, dblTotalArea, intSilkCount, strCountry)
    End Sub

    Private Sub UpdatePriceLabels()
        strTotalWithVAT = 0.0
        If dblScorePrice > 0.0 Then
            dblRoutPrice = 0.0
        End If

        Dim dblTotalExVAT As Double = Math.Round((dblAreaPrice +
                                                     dblMaskPrice +
                                                     dblSilkPrice +
                                                     dblRoutPrice +
                                                     dblScorePrice +
                                                     dblHiSpecDrillPrice +
                                                     dblHiSpecTrackGapPrice +
                                                     dblEdgePlatingPrice +
                                                     dblExtraLayoutPrice), 2)
        Dim dblVAT As Double = 1 + (CDbl(txtVAT.Text) / 100)
        strTotalWithVAT = Math.Round(dblTotalExVAT * dblVAT, 2)
        Dim strDiscountInfo As String = ""
        If chkDiscount.Checked Then
            strTotalWithVAT = (Math.Round(CDbl(strTotalWithVAT) * 0.9, 2)).ToString
            strDiscountInfo = " (10% discount)"
        End If

        lblTotalPrice.Text = "actual: " + lblTotalAPI.Text
        dblActualPrice = CDbl(strTotalWithVAT)

    End Sub

    Private Sub CalculateMaskPrice()
        If chkSoldermask.Checked = False Then
            dblMaskPrice = 0.0
            Exit Sub
        End If
        dblMaskPrice = Calculator.Calculator.GetMaskPrice(intCcts, dblTotalArea, strCountry)
    End Sub

    Private Sub CalculateAreaPrice()

        dblPricePerDm = Calculator.Calculator.GetPricePerDm(intWhatBase, strPool, intWD, strCountry)

        If dblTotalArea > 1 Then
            dblAreaPrice = dblPricePerDm * dblTotalArea
        Else
            dblAreaPrice = dblPricePerDm
        End If
    End Sub

    Private Sub SetVariables()
        btnCheckPrice.Text = "Price Up To Date"
        btnCheckPrice.BackColor = Color.LightGreen
        strAuNumber = Trim(txtAuNumber.Text)
        dblTotalArea = CDbl(txtTotalArea.Text)
        intCcts = CInt(txtCctCount.Text)
        intWD = CInt(cmbWD.Text)
        intExtraLayoutCcts = CInt(txtExtraLayouts.Text)

        If rdoALU.Checked Then
            strPool = "ALU"
        End If
        If rdoSS.Checked Then
            strPool = "SS"
        End If
        If rdoDS.Checked Then
            strPool = "DS"
        End If
        If rdo4L.Checked Then
            strPool = "4L"
        End If
        If rdo6L.Checked Then
            strPool = "6L"
        End If

        If rdo0Silk.Checked Then
            intSilkCount = 0
        End If
        If rdo1Silk.Checked Then
            intSilkCount = 1
        End If
        If rdo2Silk.Checked Then
            intSilkCount = 2
        End If

        intWhatBase = Int(dblTotalArea)

    End Sub



    Private Sub dataChanged(sender As Object, e As EventArgs)
        btnCheckPrice.Text = "Update Price (Click to refresh)"
        btnCheckPrice.BackColor = Color.Pink
        btnCheckPrice.Enabled = True
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        ImportDataFromDatabase()
        GetAPIPricing()
    End Sub

    Private Sub ImportDataFromDatabase()
        If CheckAuFieldIsFilled() Then
            Exit Sub
        End If
        ClearTextFields()
        GetResultsFromAPI()

        CalculatePriceFromUserInputAndAPI()
        btnCheckPrice.Text = "Price Up To Date"
        btnCheckPrice.BackColor = Color.LightGreen
        btnCheckPrice.Enabled = False

        lblTotalPrice.Text = "actual: " + Replace(lblTotalAPI.Text, "Total:", "")
        'GetOrderRefNumber()
        GetOrderMaterialThickness()
        GetOrderSurface()
        GetCountry()
        GetCurrencySymbol()
        'GetVAT()
        GetOriginalPrice()
        GetWD()
        SetMulti()
        GetTotalActualArea()
        GetTotalMhtArea()
        GetTotalActualCct()
        GetTotalMhtCct()
        GetPool()
        GetSilkCount()
        GetMaskCount()
        GetCutoutCharge()
        GetScoringCharge()
        GetHiSpecDrillCharge()
        GetHiSpecGapCharge()
        GetEdgePlating()
        UpdateCurrency()
        'UpdatePrice()
        'DisableImportButton()
    End Sub

    Private Sub UpdateCurrency()

        Dim strAU As String = Trim(txtAuNumber.Text)
        Dim country As String = ""
        Dim table As String
        Dim multi As Integer = 0
        Dim layouts As Integer = 0
        Dim arr_already_counted As New ArrayList

        table = "C:\in_house_files\local_database\faxback_info.dbf"

        Dim CommandText As String
        CommandText = "SELECT au_num,x_size,y_size,qty,silk,t_gaps_spec,holes_spec,ger,score,rout1 FROM " + table + " WHERE order = '" + strAU + "' order by time_stamp DESC"

        Dim ConnString As String = "Provider=VFPOLEDB.1;Data Source= " + table
        Dim Connection As New OleDbConnection(ConnString)
        Dim CommandResult As New OleDbCommand(CommandText, Connection)

        table = "C:\in_house_files\local_database\orders_info.dbf"
        CommandText = "SELECT country,au_num,currency FROM " + table + " WHERE order_num = '" + strAU + "' order by time_stamp DESC"

        ConnString = "Provider=VFPOLEDB.1;Data Source= " + table
        Connection = New OleDbConnection(ConnString)
        CommandResult = New OleDbCommand(CommandText, Connection)
        Dim language As String = "de"
        Try
            Connection.Open()
            Dim reader2 As OleDbDataReader = CommandResult.ExecuteReader(CommandBehavior.CloseConnection)
            While reader2.Read()

                If InStr(strAU, "Z") Then
                    language = "za"
                    GoTo done_already2
                End If

                Dim currency As String
                country = Trim(reader2.GetString(0))
                currency = Trim(reader2.GetString(2))

                If InStr(strAU, "I") Then
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

                Dim au_num As String = Trim(reader2.GetString(1))

                If InStr(au_num, "-") > 0 And Not arr_already_counted.Contains(au_num) Then
                    arr_already_counted.Add(au_num)
                    If (calculatorAPI.fp = True Or rdoMultiMixed.Checked = True) Then multi = 2
                    layouts = layouts + 1
                End If

done_already2:
            End While
            Connection.Close()
        Catch ex As Exception
            Connection.Close()
        End Try

        cmbLanguage.Text = language
        If language = "us" Then
            chkInches.Checked = True
        Else
            chkInches.Checked = False
        End If

        txtExtraLayouts.Text = layouts.ToString
        If multi > 0 Then
            rdoMultiMixed.Checked = True
        Else
            rdoMultiSingle.Checked = True
        End If

        GetAPIPricing()

    End Sub

    Private Sub GetResultsFromAPI()
        strJSONResult = calculatorAPI.GetTotalPrice(txtAuNumber.Text)
        Dim arr() = strJSONResult.Split(vbLf)

        Dim strVATAPI As String = "0"
        For Each line As String In arr
            If line.Contains("""vat"":") Then
                line = line.Replace("""vat"":", "")
                line = line.Replace(",", "")
                line = line.Trim
                strVATAPI = line
                txtVAT.Text = strVATAPI
            End If
        Next

        For Each line As String In arr
            If line.Contains("""baseprice"":") Then
                line = line.Replace("""baseprice"":", "")
                line = line.Replace(",", "")
                line = line.Trim
                lblAreaPriceAPI.Text = "Area: " + line
            End If

            If line.Contains("""currency"":") Then
                line = line.Replace("""currency"":", "")
                line = line.Replace(",", "")
                line = line.Replace("""", "")
                line = line.Trim
                strAPICurrency = line
            End If

            If line.Contains("""sl"":") Then
                line = line.Replace("""sl"":", "")
                line = line.Replace(",", "")
                line = line.Trim
                lblMaskPriceAPI.Text = "Mask: " + line
            End If

            If line.Contains("""bd"":") Then
                line = line.Replace("""bd"":", "")
                line = line.Replace(",", "")
                line = line.Trim
                lblSilkPriceAPI.Text = "Silk: " + line
            End If

            If line.Contains("""fp"":") Then
                line = line.Replace("""fp"":", "")
                line = line.Replace(",", "")
                line = line.Trim
                lblRoutPriceAPI.Text = "Rout: " + line
            End If

            If line.Contains("""rk"":") Then
                line = line.Replace("""rk"":", "")
                line = line.Replace(",", "")
                line = line.Trim
                lblScoringPriceAPI.Text = "Score: " + line
            End If

            If line.Contains("""dr"":") Then
                line = line.Replace("""dr"":", "")
                line = line.Replace(",", "")
                line = line.Trim
                lblHiSpecDrillsAPI.Text = "H/S Drills: " + line
            End If

            If line.Contains("""lb"":") Then
                line = line.Replace("""lb"":", "")
                line = line.Replace(",", "")
                line = line.Trim
                lblHiSpecTrackGapAPI.Text = "H/S Track/Gap: " + line
            End If

            If line.Contains("""km"":") Then
                line = line.Replace("""km"":", "")
                line = line.Replace(",", "")
                line = line.Trim
                lblEdgePlatingPriceAPI.Text = "Edge Plating: " + line
            End If

            If line.Contains("""cam"":") Then
                line = line.Replace("""cam"":", "")
                line = line.Replace(",", "")
                line = line.Trim
                lblExtraLayoutPriceAPI.Text = "Extra Layouts: " + line
            End If
        Next

    End Sub


    Private Sub GetCountry()
        strCountry = Database.Library.Database.GetCountry(txtAuNumber.Text)
    End Sub

    Private Sub GetCurrencySymbol()
        Select Case strCountry
            Case "USA"
                strCurrencySymbol = "$"
                rdoUSD.Checked = True
            Case "SOUTH_A"
                strCurrencySymbol = " "
                rdoZAR.Checked = True
            Case Else
                strCurrencySymbol = "€"
                rdoEuro.Checked = True
        End Select
    End Sub

    Private Sub DisableImportButton()
        btnImport.Enabled = False
        btnImport.BackColor = Color.LightGreen
    End Sub

    Private Sub GetEdgePlating()
        Dim result As String = Database.Library.Database.GetEdgePlatingCharge(txtAuNumber.Text)
        Select Case result
            Case "0"
                chkEdgePlating.Checked = False
            Case Else
                chkEdgePlating.Checked = True
        End Select
    End Sub

    Private Sub ClearTextFields()
        txtTotalArea.Text = ""
        cmbWD.Text = ""
        txtCctCount.Text = ""
        txtVAT.Text = ""
    End Sub

    Private Sub GetVAT()
        Select Case strCountry
            Case "USA"
                txtVAT.Text = 0
            Case "SOUTH_A"
                txtVAT.Text = 14
            Case Else
                If txtAuNumber.Text.Contains("I") Then
                    txtVAT.Text = "23"
                Else
                    txtVAT.Text = "19"
                End If
        End Select

    End Sub

    Private Sub GetHiSpecGapCharge()
        Dim result As String = Database.Library.Database.GetHiSpecGapCharge(txtAuNumber.Text)
        Select Case result
            Case "0"
                chkHiSpecGaps.Checked = False
            Case Else
                chkHiSpecGaps.Checked = True
        End Select
    End Sub

    Private Sub GetHiSpecDrillCharge()
        Dim result As String = Database.Library.Database.GetHiSpecDrillCharge(txtAuNumber.Text)
        Select Case result
            Case "0"
                chkHiSpecDrills.Checked = False
            Case Else
                chkHiSpecDrills.Checked = True
        End Select
    End Sub

    Private Sub GetScoringCharge()
        Dim result As String = Database.Library.Database.GetScoringCharge(txtAuNumber.Text)
        Select Case result
            Case "0"
                chkScoring.Checked = False
            Case Else
                chkScoring.Checked = True
        End Select
    End Sub

    Private Sub GetCutoutCharge()
        Dim result As String = Database.Library.Database.GetCutoutCharge(txtAuNumber.Text)
        Select Case result
            Case "0"
                chkCutout.Checked = False
            Case Else
                chkCutout.Checked = True
        End Select
    End Sub

    Private Sub GetMaskCount()
        'always charge mask, even if non on the pcb

        'Dim result As String = Database.Library.Database.GetMaskCount(txtAuNumber.Text)
        'Select Case result
        ' Case "0"
        'chkSoldermask.Checked = False
        'Case Else
        'chkSoldermask.Checked = True
        'End Select
        chkSoldermask.Checked = True
        chkSoldermask.Enabled = False
    End Sub

    Private Sub GetSilkCount()
        Dim result As String = Database.Library.Database.GetSilkCount(txtAuNumber.Text)
        Select Case result
            Case "0"
                rdo0Silk.Checked = True
            Case "1"
                rdo1Silk.Checked = True
            Case "2"
                rdo2Silk.Checked = True
            Case Else
                rdo0Silk.Checked = True
        End Select
    End Sub


    Private Sub GetPool()
        Dim result As String = Database.Library.Database.GetPool(txtAuNumber.Text)

        Select Case result
            Case "ALU"
                rdoALU.Checked = True
            Case "8", "810", "811", "812", "813", "814", "815", "816", "817", "818", "819", "820",
                "821", "822", "823", "824", "825", "826", "827", "828", "829", "830", "831",
                "832", "833", "834", "835", "836", "837", "838", "839", "840", "841", "842",
                "843", "844", "845", "846", "847", "848", "849", "850", "83"
                rdoSS.Checked = True
            Case "1", "3", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20",
                "21", "22", "23", "24", "25", "26", "27", "28", "29", "30",
                "31", "32", "33", "34", "35", "36", "37", "38", "39", "40",
                "41", "42", "43", "44", "45", "46", "47", "48", "49", "50"
                rdoDS.Checked = True
            Case "4", "6", "7"
                rdo4L.Checked = True
            Case "5", "2", "9"
                rdo6L.Checked = True
            Case Else

        End Select
    End Sub


    Private Sub GetTotalMhtCct()
        lblMhtCct.Text = "mht:    " + Database.Library.Database.GetMhtCct(txtAuNumber.Text)
    End Sub

    Private Sub GetTotalActualCct()
        txtCctCount.Text = Database.Library.Database.GetActualCct(txtAuNumber.Text)
        lblActualCct.Text = "actual: " + txtCctCount.Text
    End Sub


    Private Sub GetTotalMhtArea()
        Dim strArea As String = Database.Library.Database.GetMhtArea(txtAuNumber.Text)
        Dim intBase As Integer = Math.Truncate(CDbl(strArea.Replace(" (inc Frame)", "")))
        lblTotalMhtArea.Text = "mht:    " + strArea
        If intBase < 1 Then
            intBase = 1
        End If
        If intBase > 50 Then
            intBase = 50
        End If
        lblBaseMht.Text = "mht:    " + intBase.ToString

    End Sub

    'Private Sub GetOrderRefNumber()
    '    Dim ref As String = Database.Library.Database.GetOrderRef(txtAuNumber.Text)
    '    orderRefNumber = ref
    'End Sub

    Private Sub GetOrderMaterialThickness()
        Dim thickness As String = Database.Library.Database.GetOrderThickness(txtAuNumber.Text)
        orderThickness = thickness
    End Sub
    Private Sub GetOrderSurface()
        Dim surface As String = Database.Library.Database.GetOrderSurface(txtAuNumber.Text)
        orderSurface = surface
    End Sub

    Private Sub GetTotalActualArea()
        Dim strArea = Database.Library.Database.GetActualArea(txtAuNumber.Text)
        Dim intBase As Integer = Math.Truncate(CDbl(strArea))
        txtTotalArea.Text = strArea
        lblTotalActualArea.Text = "actual: " + txtTotalArea.Text
        If intBase < 1 Then
            intBase = 1
        End If
        If intBase > 50 Then
            intBase = 50
        End If
        lblBaseActual.Text = "actual: " + intBase.ToString
    End Sub


    Private Sub btnOpenNote_Click(sender As Object, e As EventArgs) Handles btnOpenNote.Click
        Try
            Process.Start("T:\Database3\NOTES\" + strAuNumber + "_note.txt")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click

        If CheckIfPriceAlreadyChecked("S:\PRICE_CONFIRMATIONS\INTERNAL\" + txtAuNumber.Text + "_price.txt") = True Then
            Exit Sub
        End If

        If CheckIfPriceAlreadyChecked("S:\PRICE_CONFIRMATIONS\" + txtAuNumber.Text + "_price.txt") = True Then
            Exit Sub
        End If

        If chkPriceChangeOthers.Checked And txtPriceChangeOthers.Text = "" Then
            MsgBox("Please enter text for other price change reason!")
            Exit Sub
        End If


        If chkPriceChangeLayouts.Checked Or
            chkPriceChangeCct.Checked Or
            chkPriceChangeArea.Checked Or
            chkPriceChangeCutouts.Checked Or
            chkPriceChangeScoring.Checked Or
            chkPriceChangeDrills.Checked Or
            chkPriceChangeGaps.Checked Or
            chkPriceChangeEdgePlating.Checked Or
            chkPriceChangeSilk.Checked Or
            chkPriceChangeOthers.Checked Then

            CreatePriceChangeInfoFiles()
        End If

        If chkUpdatePrice.Checked Then
            Database.Library.Database.UpdatePriceToInOrdersInfo(txtAuNumber.Text, dblActualPrice.ToString)
            Database.Library.Database.UpdatePriceToInOrdersInfoSQL(txtAuNumber.Text, dblActualPrice.ToString)
        End If
        Database.Library.Database.UpdateDannedToTrue(txtAuNumber.Text)
        Database.Library.Database.UpdateDannedToTrueFaxBack(txtAuNumber.Text)

        End
    End Sub

    Private Function CheckIfPriceAlreadyChecked(strPath As String)
        If File.Exists(strPath) Then
            Try
                Dim TextLine As String = ""
                Dim objReader As New StreamReader(strPath)
                Do While objReader.Peek() <> -1
                    TextLine = TextLine & objReader.ReadLine() & vbNewLine
                Loop

                MsgBox("Price already checked! No changes will be made." + vbNewLine + vbNewLine + TextLine, MsgBoxStyle.Exclamation, "Already checked")
            Catch ex As Exception

            End Try
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub CreatePriceChangeInfoFiles()

        Dim strRunning As String = ""

        Dim strName As String = """" + strOperName + """"
        Dim strDate As String = """" + Date.Now.ToString + """"
        Dim strLayouts As String = "null"
        Dim strCCT As String = "null"
        Dim strArea As String = "null"
        Dim strCutouts As String = "null"
        Dim strScoring As String = "null"
        Dim strHiSpecDrills As String = "null"
        Dim strHiSpecGapsTracks As String = "null"
        Dim strEdgePlating As String = "null"
        Dim strSilk As String = "null"
        Dim strLeadtime As String = "null"
        Dim strOther As String = "null"



        strRunning += strOperName + " " + Date.Now + vbNewLine
        strRunning += "----------------" + vbNewLine +
                      "Reasons for price change" + vbNewLine +
                      "----------------" + vbNewLine


        If chkPriceChangeLayouts.Checked Then
            strRunning += "Number Layouts " + vbNewLine +
               "Number of extra layouts has changed to " +
               txtExtraLayouts.Text + vbNewLine +
               "----------------" + vbNewLine

            strLayouts = """" + "number of extra layouts is " + " " + txtExtraLayouts.Text + """"

        End If

        If chkPriceChangeArea.Checked Then
            strRunning += "Area " + vbNewLine +
                lblTotalMhtArea.Text + vbNewLine +
                lblTotalActualArea.Text + vbNewLine +
                "----------------" + vbNewLine

            strArea = """" + lblTotalMhtArea.Text.Replace("    ", " ") + " " + lblTotalActualArea.Text + """"
        End If

        If chkPriceChangeCct.Checked Then
            strRunning += "Cct " + vbNewLine +
                lblMhtCct.Text + vbNewLine +
                lblActualCct.Text + vbNewLine +
                "----------------" + vbNewLine

            strCCT = """" + lblMhtCct.Text.Replace("    ", " ") + " " + lblActualCct.Text + """"
        End If

        If chkPriceChangeCutouts.Checked Then
            If chkCutout.Checked = True Then
                strRunning += "Cutout charge added " + vbNewLine +
                    "----------------" + vbNewLine

                strCutouts = """added"""
            Else
                strRunning += "Cutout charge removed " + vbNewLine +
                    "----------------" + vbNewLine

                strCutouts = """removed"""
            End If
        End If

        If chkPriceChangeScoring.Checked Then
            If chkPriceChangeScoring.Checked = True Then
                strRunning += "Scoring charge added " + vbNewLine +
                    "----------------" + vbNewLine

                strScoring = """added"""
            Else
                strRunning += "Scoring charge removed " + vbNewLine +
                    "----------------" + vbNewLine

                strScoring = """removed"""
            End If
        End If

        If chkPriceChangeDrills.Checked Then
            If chkHiSpecDrills.Checked = True Then
                strRunning += "Hi Spec drill charge added " + vbNewLine +
                    "----------------" + vbNewLine

                strHiSpecDrills = """added"""
            Else
                strRunning += "Hi Spec drill charge removed " + vbNewLine +
                    "----------------" + vbNewLine

                strHiSpecDrills = """removed"""
            End If
        End If

        If chkPriceChangeGaps.Checked Then
            If chkHiSpecGaps.Checked = True Then
                strRunning += "Hi Spec gaps/tracks charge added " + vbNewLine +
                    "----------------" + vbNewLine

                strHiSpecGapsTracks = """added"""
            Else
                strRunning += "Hi Spec gaps/tracks charge removed " + vbNewLine +
                    "----------------" + vbNewLine

                strHiSpecGapsTracks = """removed"""
            End If
        End If

        If chkPriceChangeEdgePlating.Checked Then
            If chkEdgePlating.Checked = True Then
                strRunning += "Edge plating charge added " + vbNewLine +
                    "----------------" + vbNewLine

                strEdgePlating = """added"""
            Else
                strRunning += "Edge plating gaps/tracks charge removed " + vbNewLine +
                    "----------------" + vbNewLine

                strEdgePlating = """removed"""
            End If
        End If

        If chkPriceChangeSilk.Checked Then
            strRunning += "Charge only one silkscreen instead of two " + vbNewLine +
                "----------------" + vbNewLine

            strSilk = """Charge only one silkscreen instead of two"""

        End If

        '  If chkPriceChangeLeadtime.Checked Then
        '  strRunning += "Charge only one silkscreen instead of two " + vbNewLine +
        '          "----------------" + vbNewLine
        '
        '        strSilk = """Charge only one silkscreen instead of two"""
        '
        '        End If


        If chkPriceChangeOthers.Checked Then
            strRunning += "Others " + vbNewLine +
                txtPriceChangeOthers.Text + vbNewLine +
                "----------------" + vbNewLine

            strOther = """" + txtPriceChangeOthers.Text + """"
            strOther = strOther.Replace("""", "")
            strOther = """" + strOther + """"
        End If


        strRunning += vbNewLine + vbNewLine + vbNewLine

        strRunning += "Comparison (all including VAT)" + vbNewLine +
            "----------------" + vbNewLine +
            lblOriginalOrderPrice.Text + vbNewLine +
            lblTotalPrice.Text + vbNewLine +
            lblPriceDiff.Text + vbNewLine

        Dim strMhtIncorrect As String = ""
        Try
            My.Computer.FileSystem.WriteAllText("S:\PRICE_CONFIRMATIONS\" + txtAuNumber.Text + "_price.txt", strRunning, False)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Try
            My.Computer.FileSystem.WriteAllText("S:\PRICE_CONFIRMATIONS\MAIN_STATUS_LIST\PRICE_CONFIRMATION_LIST.txt", txtAuNumber.Text + " _ CHANGED (" + strOperName + ")" + strMhtIncorrect + vbNewLine, True)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        BuildJSONfilePriceChange(orderSurface, strName, strDate, strLayouts, strCCT, strArea, strCutouts, strScoring, strHiSpecDrills, strHiSpecGapsTracks, strEdgePlating, strSilk, strOther)

    End Sub

    Private Sub btnGetPrice_Click(sender As Object, e As EventArgs) Handles btnGetPrice.Click
        CalculatePriceFromUserInputAndAPI()
    End Sub

    Private Sub CalculatePriceFromUserInputAndAPI()

        Dim material As String
            If rdoALU.Checked Then material = "3" Else material = "1"

            Dim layerCount As String = "1"
            If rdoALU.Checked Or rdoSS.Checked Then layerCount = "1"
            If rdoDS.Checked Then layerCount = "2"
            If rdo4L.Checked Then layerCount = "4"
            If rdo6L.Checked Then layerCount = "6"

            Dim silkCount As String = "0"
            If rdo0Silk.Checked Then silkCount = "0"
            If rdo1Silk.Checked Then silkCount = "1"
            If rdo2Silk.Checked Then silkCount = "3"

        Dim hiSpecGaps As String = "1"
            If chkHiSpecGaps.Checked Then hiSpecGaps = "2"

        Dim hiSpecDrill As String = "1"
            If chkHiSpecDrills.Checked Then hiSpecDrill = "2"

            Dim edgePlating As String = "0"
            If chkEdgePlating.Checked Then edgePlating = "1"

            Dim panelProcessing As String = "3"
            If chkCutout.Checked Then panelProcessing = "4"
            If chkScoring.Checked Then panelProcessing = "1"

        Dim multi As String = "0"
        If rdoMultiSingle.Checked = True Then multi = "0"
        If rdoMultiMixed.Checked = True Then multi = "1"
        If rdoMultiMixed.Checked = True And (calculatorAPI.fp = True) Then multi = "2"

        Dim strAreaTemp As String = txtTotalArea.Text
        If strAreaTemp = "" Then strAreaTemp = "0.0"

        'If chkInches.Checked Then
        '    strAreaTemp = (CDbl(strAreaTemp) * 15.5).ToString
        'End If

        Dim inputNumber As String = txtAuNumber.Text

        Dim precamNumber, precamNumberType As String
        Dim TheYear As String = CStr(Now.Year)
        Dim TheMonth As String = CStr(Now.Month)

        'If InStr(inputNumber, "P") Then
        ' precamNumberType = "au_number"
        ' precamNumber = "AU-" & TheYear & TheMonth & "/" & inputNumber
        ' Else
        precamNumberType = "ref_number"
        precamNumber = calculatorAPI.ref_number
        'End If

        strJSONResult = calculatorAPI.CalculatePrice(precamNumberType, precamNumber, cmbLanguage.Text, material, layerCount, txtCctCount.Text, strAreaTemp, silkCount, hiSpecGaps, hiSpecDrill, edgePlating, txtExtraLayouts.Text, panelProcessing, cmbWD.Text, multi)

        'strJSONResult = calculatorAPI.CalculatePrice(cmbLanguage.Text, material, layerCount, txtCctCount.Text, strAreaTemp, silkCount, hiSpecGaps, hiSpecDrill, edgePlating, txtExtraLayouts.Text, panelProcessing, cmbWD.Text, multi)

        Dim arr() = strJSONResult.Split(vbLf)

        Dim strVATAPI As String = "0"
            For Each line As String In arr
                If line.Contains("""vat"":") Then
                    line = line.Replace("""vat"":", "")
                    line = line.Replace(",", "")
                    line = line.Replace("""", "")
                    line = line.Trim
                    strVATAPI = line
                    txtVAT.Text = line
                End If
            Next

        Dim strCurrencySymbol As String = "€"
        rdoEuro.Checked = True
        Select Case cmbLanguage.Text
                Case "us"
                strCurrencySymbol = "$"
                rdoUSD.Checked = True
            Case "za"
                strCurrencySymbol = ""
                rdoZAR.Checked = True
        End Select

            For Each line As String In arr
                If line.Contains("""baseprice"":") Then
                    line = line.Replace("""baseprice"":", "")
                    line = line.Replace(",", "")
                    line = line.Trim
                    lblAreaPriceAPI.Text = "Area: " + strCurrencySymbol + line
                End If

                If line.Contains("""sl"":") Then
                    line = line.Replace("""sl"":", "")
                    line = line.Replace(",", "")
                    line = line.Trim
                    lblMaskPriceAPI.Text = "Mask: " + strCurrencySymbol + line
                End If

                If line.Contains("""bd"":") Then
                    line = line.Replace("""bd"":", "")
                    line = line.Replace(",", "")
                    line = line.Trim
                    lblSilkPriceAPI.Text = "Silk: " + strCurrencySymbol + line
                End If

                If line.Contains("""fp"":") Then
                    line = line.Replace("""fp"":", "")
                    line = line.Replace(",", "")
                    line = line.Trim
                    lblRoutPriceAPI.Text = "Rout: " + strCurrencySymbol + line
                End If

                If line.Contains("""rk"":") Then
                    line = line.Replace("""rk"":", "")
                    line = line.Replace(",", "")
                    line = line.Trim
                    lblScoringPriceAPI.Text = "Score: " + strCurrencySymbol + line
                End If

                If line.Contains("""dr"":") Then
                    line = line.Replace("""dr"":", "")
                    line = line.Replace(",", "")
                    line = line.Trim
                    lblHiSpecDrillsAPI.Text = "H/S Drills: " + strCurrencySymbol + line
                End If

                If line.Contains("""lb"":") Then
                    line = line.Replace("""lb"":", "")
                    line = line.Replace(",", "")
                    line = line.Trim
                    lblHiSpecTrackGapAPI.Text = "H/S Track/Gap: " + strCurrencySymbol + line
                End If

                If line.Contains("""km"":") Then
                    line = line.Replace("""km"":", "")
                    line = line.Replace(",", "")
                    line = line.Trim
                    lblEdgePlatingPriceAPI.Text = "Edge Plating: " + strCurrencySymbol + line
                End If

                If line.Contains("""cam"":") Then
                    line = line.Replace("""cam"":", "")
                    line = line.Replace(",", "")
                    line = line.Trim
                    lblExtraLayoutPriceAPI.Text = "Extra Layouts: " + strCurrencySymbol + line
                End If

                If line.Contains("""total"":") Then
                    line = line.Replace("""total"":", "")
                    line = line.Replace(",", "")
                line = line.Trim
                lblTotalAPI.Text = "Total: " + strCurrencySymbol + line

                Try
                    lblExVAT.Text = "(" + strCurrencySymbol + (Math.Round((CDbl(line) / (txtVAT.Text + 100.0)) * 100, 2)).ToString + " ex VAT)"
                    dblActualPrice = CDbl(line)
                Catch

                    ' Missing text file 

                End Try

            End If

            Next

    End Sub

    Private Sub btnOpenMht_Click(sender As Object, e As EventArgs) Handles btnOpenMht.Click
        Try

            Process.Start("S:\Job\in_house_files\capture_data\views\" + Trim(txtAuNumber.Text) + "_v.mht")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub txtAuNumber_TextChanged(sender As Object, e As EventArgs) Handles txtAuNumber.TextChanged
        btnImport.Enabled = True
        btnImport.BackColor = Color.Pink
    End Sub

    Private Sub BuildJSONfilePriceChange(strSurface As String, strName As String, strDate As String, strLayouts As String, strCCT As String, strArea As String, strCutouts As String, strScoring As String, strHiSpecDrills As String, strHiSpecGapsTracks As String, strEdgePlating As String, strSilk As String, strOther As String)

        Dim order As String = txtAuNumber.Text
        Dim strJSON As String

        Dim material As Integer = orderThickness
        If rdoALU.Checked = True Then
            material = 3 ' 3-ALU
        End If

        Dim charge_area As Double = CDbl(txtTotalArea.Text)
        Dim language As String = cmbLanguage.Text 'de,uk,za,us,fr,it,nl
        Dim quantity As Integer = txtCctCount.Text

        Dim layer_count As Integer = 2
        If rdoALU.Checked = True Or rdoSS.Checked = True Then
            layer_count = 1
        End If
        If rdo4L.Checked = True Then
            layer_count = 4
        End If
        If rdo6L.Checked = True Then
            layer_count = 6
        End If

        ' Dim soldermask As Integer = 1

        Dim silkscreen As Integer = 0
        If rdo1Silk.Checked = True Then
            silkscreen = 1
        End If
        If rdo2Silk.Checked = True Then
            silkscreen = 3  ' 1->TopOnly 2->BotOnly 3->TopAndBot
        End If

        Dim track_gap_size As Integer = 1
        If chkHiSpecGaps.Checked = True Then
            track_gap_size = 3
        End If

        Dim drill_diameter As Integer = 1
        If chkHiSpecDrills.Checked = True Then
            drill_diameter = 4
        End If
        If rdoALU.Checked = True Then
            drill_diameter = 3 ' 3-ALU
        End If

        Dim edge_metalization As Integer = 2
        If chkEdgePlating.Checked = True Then
            edge_metalization = 1
        End If

        Dim scoring As Integer = 0
        If chkScoring.Checked = True Then
            scoring = 1
        End If

        Dim cutout As Integer = 0
        If chkCutout.Checked = True Then
            cutout = 1
        End If

        Dim panel_processing As Integer = 3  ' 1-Scoring,2-Pips,3-SinglePieces,4-Cutout
        If cutout > 0 Then
            panel_processing = 4
        End If
        If scoring > 0 Then
            panel_processing = 1
        End If

        Dim lead_time As Integer = cmbWD.Text
        Dim layouts As Integer = txtExtraLayouts.Text

        Dim multi As Integer = 0
        If rdoMultiMixed.Checked = True Then
            multi = 2
        End If

        Dim TheYear As String = CStr(Now.Year)
        Dim TheMonth As String = CStr(Now.Month)
        Dim precamNumber, precamNumberType As String

        If Len(TheMonth) < 2 Then TheMonth = "0" & TheMonth

        'If InStr(order, "P") Then
        ' precamNumberType = "au_number"
        ' precamNumber = "AU-" & TheYear & TheMonth & "/" & order
        ' Else
        precamNumberType = "ref_number"
        precamNumber = calculatorAPI.ref_number
        'End If

        Dim extraInfo As String = "["
        If (calculatorAPI.pcbAreasAdded < CDbl(charge_area) + 0.1) And (calculatorAPI.pcbAreasAdded > CDbl(charge_area) - 0.1) Then
            For Each info In calculatorAPI.pcbInfoList
                extraInfo = extraInfo & "{" & """pcb_width"":" & """" & info.pcbWidth & """" & ",""pcb_length"":" & """" & info.pcbLength & """" & ",""pcb_quantity"":" & """" & info.pcbQuantity & """" & ",""pcb_precam_name"":" & """" & info.pcbPrecamName & """" & ",""pcb_ref"":" & """" & info.pcbRef & """" & "},"
            Next
        Else
            extraInfo = extraInfo & ","
        End If
        extraInfo = Mid(extraInfo, 1, extraInfo.Length - 1) & "]"


        strJSON =
"{  
    ""language"": """ & language & """,  
    """ & precamNumberType & """: """ & precamNumber & """ ,
    ""actual_product"": { 
        ""materialdicke"":" & material & ",
        ""lagenzahl"":" & layer_count & ",
        ""oberflaeche"":" & strSurface & "," + vbNewLine + " 
        ""total_quantity"":" & quantity & ",
        ""total_area"":" & charge_area & ",
        ""bestueckungsdruck"":" & silkscreen & ",
        ""leiterbahn"":" & track_gap_size & ",
        ""bohrdurchmesser"":" & drill_diameter & ",
        ""kantenmetallisierung"":" & edge_metalization & ",
        ""add_layoutfiles"":" & layouts & ",
        ""mehrfachnutzen"":" & multi & ",
        ""mehrfachnutzen_verarbeitung"":" & panel_processing & ",
        ""lieferzeit_arbeitstage"":" & lead_time & ",
        ""extra_info"":" &
        extraInfo & " " + vbNewLine + " 
        } ,
    ""original_price"": """ & strOriginalPrice & """, 
    ""actual_price"": """ & CStr(dblActualPrice) & """,  
    ""reason_for_price_change"":{
        ""checked_by"":" + strName + ",
        ""checked_date"":" + strDate + ",
        ""quantity"":" + strCCT + ",
        ""area"":" + strArea + ",
        ""silkscreen"":" + strSilk + ",
        ""track_gap_size"":" + strHiSpecGapsTracks + ",
        ""drill_diameter"":" + strHiSpecDrills + ",
        ""edge_metallisation"":" + strEdgePlating + ",
        ""add_layoutfiles"":" + strLayouts + ",
        ""cutouts"":" + strCutouts + ",
        ""scoring"":" + strScoring + ",  
        ""other"":" + strOther + "
    },
    ""pricing_api_response"":" & calculatorAPI.priceAPIResponse + vbNewLine +
    "}"
        Try
            My.Computer.FileSystem.WriteAllText("S:\PRICE_CONFIRMATIONS\JSON\" + txtAuNumber.Text + ".json", strJSON, False)
        Catch ex As Exception
        End Try


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try

            Process.Start("S:\Ordersforcam\ForCAM\" + Trim(txtAuNumber.Text) + ".gwk")

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try

            Process.Start("S:\Ordersforcam\ForCAM")

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


End Class
