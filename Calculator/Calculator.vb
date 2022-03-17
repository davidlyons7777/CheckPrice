Public Class Calculator

    Public Shared Function GetPricePerDm(intWhatBase As Integer, strType As String, intWD As Integer, strCountry As String) As String

        If intWD < 0 Then Return -1
        If intWD > 6 Then intWD = 6
        If intWhatBase > 10 And strType = "4L" Then intWhatBase = 10
        If intWhatBase > 10 And strType = "6L" Then intWhatBase = 10
        If intWhatBase > 50 Then intWhatBase = 50
        If intWhatBase < 1 Then intWhatBase = 1

        Dim arrBasePriceALU_EURO As New ArrayList From {0, 83.2094, 66.5676, 49.9257, 41.6048, 38.7619, 35.959, 33.1361, 30.3132, 27.4903, 24.6303, 24.1118, 23.6391, 23.1663, 22.6935, 22.2208, 21.7479, 21.2751, 20.8024, 20.172, 19.5368, 18.7851, 18.1233, 17.5244, 16.9729, 16.4686, 15.9958, 15.523, 15.129, 14.766, 14.4419, 14.1048, 13.8053, 13.5216, 13.2538, 13.0016, 12.7493, 12.5129, 12.2924, 12.0875, 11.8873, 11.6776, 11.5044, 11.331, 11.1576, 10.9843, 10.8109, 10.6534, 10.4958, 10.3383, 10.1899}
        Dim arrBasePriceSS_EURO As New ArrayList From {0, 37.5, 21.0715, 20.071, 19.0705, 18.07, 17.0695, 16.069, 15.0685, 14.068, 13.0697, 12.7861, 12.5445, 12.2924, 12.0402, 11.788, 11.5359, 11.2837, 11.0316, 10.6954, 10.3592, 9.9599, 9.6132, 9.2981, 9.0039, 8.7307, 8.4891, 8.2369, 8.0268, 7.8376, 7.6591, 7.4804, 7.3229, 7.1758, 7.0287, 6.8921, 6.7661, 6.6399, 6.5244, 6.4088, 6.3038, 6.1987, 6.1041, 6.0096, 5.915, 5.8309, 5.7364, 5.6524, 5.5683, 5.4843, 5.4002}
        Dim arrBasePriceDS_EURO As New ArrayList From {0, 49.5, 26.4821, 25.2246, 23.9668, 22.7091, 21.4513, 20.1905, 18.9358, 17.6781, 16.4203, 16.046, 15.7594, 15.4442, 15.129, 14.8139, 14.4987, 14.1834, 13.8682, 13.448, 13.0245, 12.5234, 12.0822, 11.6829, 11.3152, 10.9791, 10.6639, 10.3487, 10.086, 9.8444, 9.6279, 9.4031, 9.2034, 9.0144, 8.8358, 8.6676, 8.4996, 8.3419, 8.1948, 8.0583, 7.9249, 7.7851, 7.6696, 7.554, 7.4384, 7.3229, 7.2073, 7.1022, 6.9971, 6.8921, 6.7934}
        Dim arrBasePrice4L_EURO As New ArrayList From {0, 92.83, 74.2677, 55.7008, 46.4173, 45.6853, 44.9532, 44.2212, 43.4891, 42.7571, 42.025}
        Dim arrBasePrice6L_EURO As New ArrayList From {0, 116.0433, 92.8347, 69.626, 58.0217, 57.1067, 56.1916, 55.2766, 54.3615, 53.4465, 52.5313}

        Dim arrBasePriceALU_US As New ArrayList From {0, 2.9795, 2.3836, 1.7877, 1.4897, 1.3887, 1.2876, 1.1865, 1.0854, 0.9844, 0.8819, 0.8634, 0.8465, 0.8295, 0.8126, 0.7957, 0.7787, 0.7618, 0.7449, 0.7223, 0.6996, 0.6727, 0.649, 0.6275, 0.6078, 0.5897, 0.5728, 0.5558, 0.5417, 0.5288, 0.5172, 0.5051, 0.4943, 0.4842, 0.4746, 0.4656, 0.4565, 0.4481, 0.4402, 0.4328, 0.4257, 0.4182, 0.412, 0.4057, 0.3995, 0.3933, 0.3872, 0.3815, 0.3758, 0.3702, 0.3649}
        Dim arrBasePriceSS_US As New ArrayList From {0, 2.6855, 1.509, 1.4374, 1.3657, 1.2941, 1.2224, 1.1508, 1.0791, 1.0075, 0.936, 0.9156, 0.8983, 0.8803, 0.8622, 0.8442, 0.8261, 0.8081, 0.79, 0.7659, 0.7419, 0.7133, 0.6884, 0.6659, 0.6448, 0.6252, 0.6079, 0.5899, 0.5748, 0.5613, 0.5485, 0.5357, 0.5244, 0.5139, 0.5033, 0.4936, 0.4845, 0.4755, 0.4672, 0.459, 0.4514, 0.4439, 0.4371, 0.4304, 0.4236, 0.4176, 0.4108, 0.4048, 0.3988, 0.3927, 0.3867}
        Dim arrBasePriceDS_US As New ArrayList From {0, 3.544, 1.8965, 1.8064, 1.7164, 1.6263, 1.5362, 1.4459, 1.3561, 1.266, 1.1759, 1.1511, 1.1286, 1.106, 1.0834, 1.0609, 1.0383, 1.0157, 0.9931, 0.9631, 0.9327, 0.8968, 0.8652, 0.8366, 0.8103, 0.7862, 0.7637, 0.7411, 0.7223, 0.705, 0.6895, 0.6734, 0.6591, 0.6455, 0.6328, 0.6207, 0.6087, 0.5974, 0.5869, 0.5771, 0.5675, 0.5575, 0.5492, 0.541, 0.5327, 0.5244, 0.5161, 0.5086, 0.5011, 0.4936, 0.4865}
        Dim arrBasePrice4L_US As New ArrayList From {0, 6.6482, 5.3185, 3.9889, 3.3241, 3.2576, 3.2193, 3.1668, 3.1144, 3.062, 3.0095}
        Dim arrBasePrice6L_US As New ArrayList From {0, 8.3102, 6.6482, 4.9862, 4.1552, 4.0896, 4.0241, 3.9585, 3.893, 3.8275, 3.7619}

        Dim arrBasePriceALU_ZA As New ArrayList From {0, 1464.4869, 1464.4869, 1464.4869, 1464.4869, 665.6751, 665.6751, 665.6751, 665.6751, 665.6751, 433.4925, 424.3685, 416.0478, 407.7269, 399.4061, 391.0839, 382.763, 374.4423, 366.1229, 355.0274, 343.8478, 330.6183, 318.9696, 308.429, 298.7223, 289.8469, 281.5246, 273.2053, 266.2703, 259.8914, 254.178, 248.2428, 242.9718, 237.9801, 233.2651, 228.8281, 224.3868, 220.2278, 216.3454, 212.7396, 209.2168, 205.5266, 202.4769, 199.4258, 196.3731, 193.3234, 190.2723, 187.4991, 184.7246, 181.9515, 179.3425}
        Dim arrBasePriceSS_ZA As New ArrayList From {0, 466.1592, 466.1592, 466.1592, 466.1592, 351.3259, 351.3259, 351.3259, 351.3259, 351.3259, 276.0331, 225.036, 220.7824, 216.3454, 211.9069, 207.4685, 203.0315, 198.5631, 194.1561, 188.2387, 182.3212, 175.2945, 169.1922, 163.1922, 158.4681, 153.6598, 149.4077, 144.9693, 141.2718, 137.9426, 134.7997, 131.6554, 128.8823, 126.294, 123.7044, 121.301, 119.0825, 116.8625, 114.8289, 112.7953, 110.9466, 109.0963, 107.4325, 105.7686, 104.1047, 102.6243, 100.9604, 99.4829, 98.0025, 96.5234, 95.043}
        Dim arrBasePriceDS_ZA As New ArrayList From {0, 585.8696, 585.8696, 585.8696, 585.8696, 441.5464, 441.5464, 441.5464, 441.5464, 441.5464, 346.7957, 282.9133, 277.3656, 271.8179, 266.2703, 260.724, 255.1763, 249.6286, 244.081, 236.6845, 229.2319, 220.4126, 212.6465, 205.6197, 199.1477, 193.2317, 187.684, 182.1363, 177.2609, 173.2609, 169.451, 165.4947, 161.9806, 158.6529, 155.5101, 152.5506, 149.5926, 146.818, 144.2283, 141.8264, 139.4778, 137.0183, 134.9846, 132.951, 130.9159, 128.8823, 126.8487, 124.9984, 123.1497, 121.301, 119.5632}
        Dim arrBasePrice4L_ZA As New ArrayList From {0, 1633.8891, 1633.8891, 1633.8891, 1633.8891, 746.228, 746.228, 746.228, 746.228, 746.228, 739.6407}
        Dim arrBasePrice6L_ZA As New ArrayList From {0, 2042.3617, 2042.3617, 2042.3617, 2042.3617, 932.8777, 932.8777, 932.8777, 932.8777, 932.8777, 924.5502}


        Dim dblWDMultiplier As Double
        dblWDMultiplier = SetMultiplier(intWD, dblWDMultiplier, strType)

        Dim arrBasePriceALU As New ArrayList
        Dim arrBasePriceSS As New ArrayList
        Dim arrBasePriceDS As New ArrayList
        Dim arrBasePrice4L As New ArrayList
        Dim arrBasePrice6L As New ArrayList
        Dim dblUSMultiplier As Double = 1.0

        Select Case strCountry
            Case "USA"
                arrBasePriceALU = arrBasePriceALU_US
                arrBasePriceSS = arrBasePriceSS_US
                arrBasePriceDS = arrBasePriceDS_US
                arrBasePrice4L = arrBasePrice4L_US
                arrBasePrice6L = arrBasePrice6L_US
                dblUSMultiplier = 16.0
            Case "SOUTH_A"
                arrBasePriceALU = arrBasePriceALU_ZA
                arrBasePriceSS = arrBasePriceSS_ZA
                arrBasePriceDS = arrBasePriceDS_ZA
                arrBasePrice4L = arrBasePrice4L_ZA
                arrBasePrice6L = arrBasePrice6L_ZA
            Case Else
                arrBasePriceALU = arrBasePriceALU_EURO
                arrBasePriceSS = arrBasePriceSS_EURO
                arrBasePriceDS = arrBasePriceDS_EURO
                arrBasePrice4L = arrBasePrice4L_EURO
                arrBasePrice6L = arrBasePrice6L_EURO
        End Select

        Select Case strType
            Case "ALU"
                Return arrBasePriceALU(intWhatBase) * dblWDMultiplier * dblUSMultiplier
            Case "SS"
                Return arrBasePriceSS(intWhatBase) * dblWDMultiplier * dblUSMultiplier
            Case "DS"
                Return arrBasePriceDS(intWhatBase) * dblWDMultiplier * dblUSMultiplier
            Case "4L"
                Return arrBasePrice4L(intWhatBase) * dblWDMultiplier * dblUSMultiplier
            Case "6L"
                Return arrBasePrice6L(intWhatBase) * dblWDMultiplier * dblUSMultiplier
            Case Else
                Return -1
        End Select

    End Function




    Private Shared Function SetMultiplier(intWD As Integer, dblMultiplier As Double, strType As String) As Double
        Select Case intWD
            Case 6
                dblMultiplier = 1.0
            Case 5
                If strType = "4L" Or strType = "6L" Then
                    dblMultiplier = 2.0
                Else
                    dblMultiplier = 2.0
                End If
            Case 4
                dblMultiplier = 3.0
            Case 3
                If strType = "4L" Or strType = "6L" Then
                    dblMultiplier = 2.8
                Else
                    dblMultiplier = 3.5
                End If
            Case 2
                If strType = "4L" Or strType = "6L" Then
                    dblMultiplier = 3.2
                Else
                    dblMultiplier = 4.0
                End If
            Case 1
                dblMultiplier = 4.6
            Case 0
                dblMultiplier = 6.0
        End Select

        Return dblMultiplier
    End Function



    Public Shared Function GetSilkPrice(intCcts As Integer, dblArea As Double, intSilkCount As Integer, strCountry As String) As Double

        If intSilkCount = 0 Then
            Return 0.0
        End If

        Dim dblBasePrice = 0.0
        Dim dblFlexPrice1 = 0.0
        Dim dblFlexPrice2 = 0.0
        Dim dblMaxPrice = 0.0

        Select Case strCountry
            Case "USA"
                dblBasePrice = 6.771
                dblFlexPrice1 = 0.0
                dblFlexPrice2 = 0.0612
                dblMaxPrice = 39.9489
                If intSilkCount = 2 Then
                    dblBasePrice = 13.542
                    dblFlexPrice1 = 0.0
                    dblFlexPrice2 = 0.1223
                    dblMaxPrice = 79.8978
                End If
            Case "SOUTH_A"
                If strCountry = "SOUTH_A" Then
                    dblBasePrice = 88.0
                    dblFlexPrice1 = 0.00
                    dblFlexPrice2 = 12.32
                    dblMaxPrice = 519.2
                    If intSilkCount = 2 Then
                        dblBasePrice = 176.0
                        dblFlexPrice1 = 0.0
                        dblFlexPrice2 = 24.64
                        dblMaxPrice = 1038.4
                    End If
                End If
            Case Else
                dblBasePrice = 5.0
                dblFlexPrice1 = 0.0
                dblFlexPrice2 = 0.7
                dblMaxPrice = 29.5

                If intSilkCount = 2 Then
                    dblBasePrice = 10.0
                    dblFlexPrice1 = 0.0
                    dblFlexPrice2 = 1.4
                    dblMaxPrice = 59.0
                End If
        End Select

        Dim dblTotalPrice = dblBasePrice + (intCcts * dblFlexPrice1) + (dblArea * dblFlexPrice2)
        If dblTotalPrice > dblMaxPrice And intSilkCount = 1 Then
            dblTotalPrice = dblMaxPrice
        End If

        If dblTotalPrice > dblMaxPrice And intSilkCount = 2 Then
            dblTotalPrice = dblMaxPrice
        End If

        Return dblTotalPrice
    End Function


    Public Shared Function GetMaskPrice(intCcts As Integer, dblArea As Double, strCountry As String) As Double
        Dim dblBasePrice = 5.0
        Dim dblFlexPrice1 = 0.0
        Dim dblFlexPrice2 = 0.7
        Dim dblMaxPrice = 49.0

        If strCountry = "USA" Then
            dblBasePrice = 6.771
            dblFlexPrice1 = 0.0
            dblFlexPrice2 = 0.0612
            dblMaxPrice = 66.3558
        End If

        If strCountry = "SOUTH_A" Then
            dblBasePrice = 88.0
            dblFlexPrice1 = 0.0
            dblFlexPrice2 = 12.32
            dblMaxPrice = 862.4
        End If

        Dim totalPrice = dblBasePrice + (intCcts * dblFlexPrice1) + (dblArea * dblFlexPrice2)
        If totalPrice > dblMaxPrice Then
            totalPrice = dblMaxPrice
        End If
        Return totalPrice
    End Function



    Public Shared Function GetRoutPrice(intCcts As Integer, boolCutouts As Boolean, dblArea As Double, strCountry As String, boolDashExists As Boolean) As Double
        Dim price As Double
        Select Case strCountry
            Case "USA"
                price = 4.3334 + (intCcts * 1.3542)
                If price > 66.3558 Then
                    price = 66.3558
                End If
            Case "SOUTH_A"
                price = 56.32 + (intCcts * 17.6)
                If price > 862.4 Then
                    price = 862.4
                End If
            Case Else
                price = 3.2 + (intCcts * 1.0)
                If price > 49.0 Then
                    price = 49.0
                End If
        End Select

        If boolCutouts = True Then
            Return price
        End If
        If boolDashExists = True Then
            Return price
        End If
        If CDbl(intCcts) <= dblArea Then
            price = 0.0
        End If
        If intCcts = 1 Then
            price = 0.0
        End If
        Return price
    End Function



    Public Shared Function GetScorePrice(intCcts As Integer, boolScoring As Boolean, strCountry As String) As Double
        Dim dblBase As Double
        Dim dblFlex1 As Double
        Dim dblMax As Double

        Select Case strCountry
            Case "USA"
                dblBase = 39.2718
                dblFlex1 = 1.0834
                dblMax = 93.4398
            Case "SOUTH_A"
                dblBase = 510.4
                dblFlex1 = 14.08
                dblMax = 1214.4
            Case Else
                dblBase = 29.0
                dblFlex1 = 0.8
                dblMax = 69.0
        End Select


        If boolScoring = True Then
            Dim dblPrice As Double = dblBase + intCcts * dblFlex1
            If dblPrice > dblMax Then
                dblPrice = dblMax
            End If
            Return dblPrice
        Else
            Return 0
        End If

    End Function




    Public Shared Function GetHiSpecPrice(dblArea As Double, dblPriceFromSheet As Double, boolHiSpecRequired As Boolean) As Double
        If boolHiSpecRequired = True Then
            Return 0.05 * dblArea * dblPriceFromSheet
        Else
            Return 0
        End If
    End Function


    Public Shared Function GetEdgePlatingPrice(intCcts As Integer, boolEdgePlating As Boolean, strCountry As String) As Double
        Dim dblBase As Double
        Dim dblFlex1 As Double

        Select Case strCountry
            Case "USA"
                dblBase = 12.1878
                dblFlex1 = 0.7854
            Case "SOUTH_A"
                dblBase = 158.4
                dblFlex1 = 10.208
            Case Else
                dblBase = 9.0
                dblFlex1 = 0.58
        End Select

        If boolEdgePlating = True Then
            Return dblBase + intCcts * dblFlex1
        Else
            Return 0
        End If
    End Function


    Public Shared Function GetExtraLayoutPrice(intCcts As Integer, strCountry As String) As Double
        If intCcts <= 0 Then
            Return 0
        End If
        Select Case strCountry
            Case "USA"
                Return intCcts * 13.2035
            Case "SOUTH_A"
                Return intCcts * 171.6
            Case Else
                Return intCcts * 9.75
        End Select


    End Function


    Private Shared Function AreaPrice(strArea As String, intCcts As Integer, dblPricePerDm As String) As Double
        Dim dblAreaPrice As Double

        If CDbl(strArea) > 1 Then
            dblAreaPrice = dblPricePerDm * CDbl(strArea)
        Else
            dblAreaPrice = dblPricePerDm
        End If

        Return dblAreaPrice
    End Function


    Private Shared Function GetAreaPrice(strWD As String, intWhatBase As Integer, strPool As String, strCountry As String) As String
        Return GetPricePerDm(intWhatBase, strPool, CInt(strWD), strCountry)
    End Function


    Private Shared Function ConvertPoolString(strPool)
        Select Case strPool
            Case "ALU"
                strPool = "ALU"
            Case "8", "810", "811", "812", "813", "814", "815", "816", "817", "818", "819", "820",
                "821", "822", "823", "824", "825", "826", "827", "828", "829", "830", "831",
                "832", "833", "834", "835", "836", "837", "838", "839", "840", "841", "842",
                "843", "844", "845", "846", "847", "848", "849", "850", "83"
                strPool = "SS"
            Case "1", "3", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20",
                "21", "22", "23", "24", "25", "26", "27", "28", "29", "30",
                "31", "32", "33", "34", "35", "36", "37", "38", "39", "40",
                "41", "42", "43", "44", "45", "46", "47", "48", "49", "50"
                strPool = "DS"
            Case "4", "6", "7"
                strPool = "4L"
            Case "5", "2", "9"
                strPool = "6L"
            Case Else
                strPool = "DS"
        End Select
        Return strPool
    End Function

    Public Shared Function GetTotalPrice(strAU As String) As String

        'information from database
        Dim dblTotalArea As Double = CDbl(Database.Library.Database.GetActualArea(strAU))
        Dim strWD As String = Database.Library.Database.GetInfoFromDatabase("work_days",
                                                                            "orders_info",
                                                                            "au_num",
                                                                            strAU,
                                                                            "C:\in_house_files\local_database\orders_info.dbf")
        Dim strCountry As String = Database.Library.Database.GetCountry(strAU)
        Dim strCctCount = Database.Library.Database.GetActualCct(strAU)

        Dim strVAT As String = "19"
        If strAU.Contains("I") Then strVAT = "23"
        If strCountry = "USA" Then strVAT = "0"
        If strCountry = "SOUTH_A" Then strVAT = "14"

        Dim intWhatBase As Integer = Math.Truncate(dblTotalArea)
        If intWhatBase < 1 Then intWhatBase = 1
        If intWhatBase > 50 Then intWhatBase = 50

        Dim intCcts As Integer = Database.Library.Database.GetActualCct(strAU)
        Dim strPool As String = ConvertPoolString(Database.Library.Database.GetPool(strAU))
        Dim intSilkCount As String = Database.Library.Database.GetSilkCount(strAU)
        Dim strCutoutCharge As String = Database.Library.Database.GetCutoutCharge(strAU)
        Dim boolCutoutChargeRequired As Boolean = False
        If strCutoutCharge = "1" Then
            boolCutoutChargeRequired = True
        End If

        Dim strScoringCharge As String = Database.Library.Database.GetScoringCharge(strAU)
        Dim boolScoringRequired As Boolean = False
        If strScoringCharge = "1" Then
            boolScoringRequired = True
        End If

        Dim strHiSpecDrill As String = Database.Library.Database.GetHiSpecDrillCharge(strAU)
        Dim boolHiSpecDrillsRequired As Boolean = False
        If strHiSpecDrill = "1" Then
            boolHiSpecDrillsRequired = True
        End If

        Dim strHiSpecGap As String = Database.Library.Database.GetHiSpecGapCharge(strAU)
        Dim boolHiSpecGapsRequired As Boolean = False
        If strHiSpecGap = "1" Then
            boolHiSpecGapsRequired = True
        End If

        Dim strEdgePlating As String = Database.Library.Database.GetEdgePlatingCharge(strAU)
        Dim boolEdgePlatingRequired As Boolean = False
        If strEdgePlating = "1" Then
            boolEdgePlatingRequired = True
        End If

        'calculate price per dm
        Dim dblPricePerDm As String = GetAreaPrice(strWD, intWhatBase, strPool, strCountry)

        'check if dash part exists for rout charge
        Dim boolDashExists As Boolean = Database.Library.Database.CheckIfDashOneExists(strAU)

        'now that we have total area, price per dm and cct the other prices can be calculated
        Dim dblAreaPrice As Double = AreaPrice(dblTotalArea, intCcts, dblPricePerDm)
        Dim dblMaskPrice As Double = GetMaskPrice(intCcts, dblTotalArea, strCountry)
        Dim dblSilkPrice As Double = GetSilkPrice(intCcts, dblTotalArea, intSilkCount, strCountry)
        Dim dblRoutPrice As Double = GetRoutPrice(intCcts, boolCutoutChargeRequired, dblTotalArea, strCountry, boolDashExists)
        Dim dblScorePrice As Double = GetScorePrice(intCcts, boolScoringRequired, strCountry)
        Dim dblHiSpecDrillPrice As Double = GetHiSpecPrice(dblTotalArea, dblPricePerDm, boolHiSpecDrillsRequired)
        Dim dblHiSpecTrackGapPrice As Double = GetHiSpecPrice(dblTotalArea, dblPricePerDm, boolHiSpecGapsRequired)
        Dim dblEdgePlatingPrice As Double = GetEdgePlatingPrice(intCcts, boolEdgePlatingRequired, strCountry)

        'don't charge rout if there is scoring
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
                                                     dblEdgePlatingPrice), 2)
        Dim dblVAT As Double = 1 + (CDbl(strVAT) / 100)
        Dim strTotalWithVAT = Math.Round(dblTotalExVAT * dblVAT, 2)

        Return strTotalWithVAT

    End Function



End Class
