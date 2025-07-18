VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "holiday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public holidayName As String

Public Function holidayHantei1(hanteiBi As Date) As Boolean
'国民の祝日を判定し、祝日ならTrue,そうでないならfalseを返す
'holidayNameに祝日名を入れる。
Dim uminohi As Long
Dim yamanohi As Long
Dim supotsuhi As Long
Dim seijinhi As Long
Dim keirouhi As Long
Dim shunbunhi As Long
Dim shuubunhi As Long

holidayHantei1 = False

'祝日が日と曜日で決まっているもの
'元日
    Select Case Month(hanteiBi)
        Case 1
            '1月1日（元日）
            If Day(hanteiBi) = 1 Then
                holidayHantei1 = True
                holidayName = "元日"
            End If
            '1月の第２月曜日（成人の日）
            seijinhi = hizukeFind(hanteiBi, 2)
            If Day(hanteiBi) = seijinhi Then
                holidayHantei1 = True
                holidayName = "成人の日"
            End If
        Case 2
            '2月11日（建国記念の日：政令により定められる）
            If Day(hanteiBi) = 11 Then
                holidayHantei1 = True
                holidayName = "建国記念の日"
            End If
            '2月23日（天皇誕生日）
            If Day(hanteiBi) = 23 Then
                holidayHantei1 = True
                holidayName = "天皇誕生日"
            End If
        Case 3
            '春分の日
            shunbunhi = Int(20.8431 + 0.242194 * (year(hanteiBi) - 1980) - Int((year(hanteiBi) - 1980) / 4))
             If Day(hanteiBi) = shunbunhi Then
                holidayHantei1 = True
                holidayName = "春分の日"
            End If
        Case 4
            '4月29日（昭和の日）
            If Day(hanteiBi) = 29 Then
                holidayHantei1 = True
                holidayName = "昭和の日"
            End If
        Case 5
            '5月3日（憲法記念日）
            If Day(hanteiBi) = 3 Then
                holidayHantei1 = True
                holidayName = "憲法記念日"
            End If
            '5月4日（みどりの日）
            If Day(hanteiBi) = 4 Then
                holidayHantei1 = True
                holidayName = "みどりの日"
            End If
            '5月5日（こどもの日）
            If Day(hanteiBi) = 5 Then
                holidayHantei1 = True
                holidayName = "こどもの日"
            End If
        Case 7
            '7月第３月曜日（海の日：ただし2020年は7月23日、2021年は7月22日）
            If year(hanteiBi) = 2020 Then
                uminohi = 23
            ElseIf year(hanteiBi) = 2021 Then
                uminohi = 22
            Else
                uminohi = hizukeFind(hanteiBi, 3)
            End If
            If Day(hanteiBi) = uminohi Then
                holidayHantei1 = True
                holidayName = "海の日"
            End If
            '2020年のみスポーツの日が7月24日
            '2021年のみスポーツの日が7月23日
            If year(hanteiBi) = 2020 Then
                supotsuhi = 24
            ElseIf year(hanteiBi) = 2021 Then
                supotsuhi = 23
            End If
            If Day(hanteiBi) = supotsuhi Then
                holidayHantei1 = True
                holidayName = "スポーツの日"
            End If
        Case 8
            '8月11日（山の日）
            If year(hanteiBi) = 2020 Then
                yamanohi = 10
            ElseIf year(hanteiBi) = 2021 Then
                yamanohi = 8
            Else
                yamanohi = 11
            End If
            If Day(hanteiBi) = yamanohi Then
                holidayHantei1 = True
                holidayName = "山の日"
            End If
        Case 9
            '秋分の日
            shuubunhi = Int(23.2488 + 0.242194 * (year(hanteiBi) - 1980)) - Int((year(hanteiBi) - 1980) / 4)
            If Day(hanteiBi) = shuubunhi Then
                holidayHantei1 = True
                holidayName = "秋分の日"
            End If
            '9月第３月曜日（敬老の日）
            keirouhi = hizukeFind(hanteiBi, 3)
            If Day(hanteiBi) = keirouhi Then
                holidayHantei1 = True
                holidayName = "敬老の日"
            End If
        Case 10
            If year(hanteiBi) <> 2020 And year(hanteiBi) <> 2021 Then
                supotsuhi = hizukeFind(hanteiBi, 2)
            End If
            If Day(hanteiBi) = supotsuhi Then
                holidayHantei1 = True
                holidayName = "スポーツの日"
            End If
        Case 11
            '11月3日（文化の日）
            If Day(hanteiBi) = 3 Then
                holidayHantei1 = True
                holidayName = "文化の日"
            End If
            '11月23日（勤労感謝の日）
            If Day(hanteiBi) = 23 Then
                holidayHantei1 = True
                holidayName = "勤労感謝の日"
            End If
    End Select
End Function

Public Function holidayHantei2(hanteiBi As Date, holiday1 As Boolean) As Boolean
    '国民の祝日に関する法律３条２項判定
    Dim zenjitsu As Date
    Dim zenhori As Date
    
    zenjitsu = DateSerial(year(hanteiBi), Month(hanteiBi), Day(hanteiBi) - 1)
    If holiday1 = False Then
        zenhori = zenjitsu
        Do While holidayHantei1(zenhori) = True
            If Weekday(zenhori) = vbSunday Then
                holidayHantei2 = True
            End If
            zenhori = DateSerial(year(zenhori), Month(zenhori), Day(zenhori) - 1)
        Loop
    Else
        holidayHantei2 = True
    End If
    '国民の祝日に関する法律３条３項判定
    Dim yokujitsu As Date
    
    yokujitsu = DateSerial(year(hanteiBi), Month(hanteiBi), Day(hanteiBi) + 1)
    If holiday1 = False Then
        If (holidayHantei1(zenjitsu) = True) And (holidayHantei1(yokujitsu) = True) Then
            holidayHantei2 = True
        End If
    End If
End Function

Function hizukeFind(hanteiBi As Date, youbikai As Long) As Long
'当該月の最初の指定曜日の日付を返す
Dim i As Long
Dim count As Long

    count = 0
    For i = 1 To 21
        If Weekday(DateSerial(year(hanteiBi), Month(hanteiBi), i), vbSunday) = vbMonday Then
            count = count + 1
            If count >= youbikai Then
                hizukeFind = i
                Exit For
            End If
        End If
    Next i
End Function

Public Function preEigyoBi(searchDay As Date) As Date
    '直前の営業日を返す
    '2024年5月6日　作成開始
    '作成者：森　一義
    Dim hanteiFlag As Boolean
    hanteiFlag = True
    searchDay = DateSerial(year(searchDay), Month(searchDay), Day(searchDay) - 1)
    Do While hanteiFlag
        If Weekday(searchDay) = vbSunday Then
            hanteiFlag = True
        ElseIf Weekday(searchDay) = vbSaturday Then
            hanteiFlag = True
        ElseIf holidayHantei2(searchDay, holidayHantei1(searchDay)) Then
            hanteiFlag = True
        Else
            hanteiFlag = False
        End If
        If hanteiFlag Then
            searchDay = DateSerial(year(searchDay), Month(searchDay), Day(searchDay) - 1)
        End If
    Loop
    preEigyoBi = searchDay
End Function

Public Function nextEigyoBi(searchDay As Date) As Date
    '直後の営業日を返す。
    '2024年5月6日　作成開始
    '作成者：森　一義
    Dim hanteiFlag As Boolean
    hanteiFlag = True
    searchDay = DateSerial(year(searchDay), Month(searchDay), Day(searchDay) + 1)
    Do While hanteiFlag
        If Weekday(searchDay) = vbSunday Then
            hanteiFlag = True
        ElseIf Weekday(searchDay) = vbSaturday Then
            hanteiFlag = True
        ElseIf holidayHantei2(searchDay, holidayHantei1(searchDay)) Then
            hanteiFlag = True
        Else
            hanteiFlag = False
        End If
        If hanteiFlag Then
            searchDay = DateSerial(year(searchDay), Month(searchDay), Day(searchDay) + 1)
        End If
    Loop
    nextEigyoBi = searchDay
End Function


