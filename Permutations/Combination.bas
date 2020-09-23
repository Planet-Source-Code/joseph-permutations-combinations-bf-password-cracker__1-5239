Attribute VB_Name = "combination"
'Source code developed by Joseph Ninan
'2nd Year BTech Computer Science and Engineering
'Sree Chitra Thirunal College of Engineering
'Papanamcode, Trivandrum-18
'Affliated to University of Kerala
'Residential Address
'Liju Bhavan, Muttampuram Lane, Sreekariyam PO
'Trivandrum
'Kerala state
'India
'PIN 695017
'Tel No 0091-471-449977
'email josephninan@crosswinds.net   liju_trv@yahoo.com
'Web Site http://www.jofu.8m.com
Public char(260) As String
Public OChoice, TotalChar, NextBlock, ASCIIStart, Max, Counter As Integer
Public Ca, Cb, Cc, Cd, Ce, Cf, Cg, Ch, Cbi As Integer
Public NoOfComb As Long
Public result As Variant
Public Const TOTALUP = 26
Public Const TOTALLOW = 26
Public Const TOTALDIGITS = 10
Public fact As Double

Public Sub Initialize()
    If frmMain.optOutTextBox.Value = True Then
        OChoice = 1
    ElseIf frmMain.optOutFile.Value = True Then
        OChoice = 2
    ElseIf frmMain.optOutSendkeys.Value = True Then
        OChoice = 3
    ElseIf frmMain.optByCode.Value = True Then
        OChoice = 4
    End If
    
    pcount = 0
End Sub
Public Sub BFInitialise()
If frmMain.optPmode1.Value = True Then
    bf2dummy = frmMain.txtBFSetup.Text
    bf2len = Len(bf2dummy)
    a = InStr(1, bf2dummy, vbCrLf)
    While (a <> bf2len + 1 And a <> 0)
        bf2dummy = Left(bf2dummy, a - 1) & Mid(bf2dummy, a + 2)
        bf2len = Len(bf2dummy)
        a = InStr(1, bf2dummy, vbCrLf)
    Wend
    SendKeys bf2dummy, True
Else
    bfStart = 1
    BFL = Len(frmMain.txtBFSetup.Text)
    For bfi = 1 To BFL
        bfStop = InStr(bfStart, frmMain.txtBFSetup.Text, vbCrLf)
        If bfStop = 0 Or bfStop = BFL + 1 Then bfStop = BFL + 1
        bfkey = Mid(frmMain.txtBFSetup.Text, bfStart, (bfStop - bfStart))
        bfi = bfStop + 1
        bfStart = bfStop + 2
        SendKeys bfkey
    Next bfi
End If
End Sub

'Now what should i do
'I have got all the input chars in char(Totalchar)
'I have got a fixed length in n=frmmain.txtlength
'First i have to get all the n letter combinations
'Then i have to permute them
'How to do?
'I got the code for doing nC3
'Will try to do the rest later

Public Sub GenOutput(cho As Integer)
Max = TotalChar + 1 - c
StartTime = Time
Select Case cho
Case 1:
    For Ca = 0 To Max - 1
            OpenForms = DoEvents
            NoOfComb = NoOfComb + 1
            comb = char(Ca)
            If frmMain.optCPMode.Value = True Then
                permute (comb)
            Else
                Select Case OChoice
                Case 1:
                result = result & vbCrLf & comb
                Case 2:
                Print #1, comb
                Case 3:
                SendKeys comb: BFInitialise
                Case 4:
                processcode (comb)
                End Select
            End If
            
            'NoOfComb = NoOfComb + 1
            'Debug.Print comb, NoOfComb
    Next Ca

Case 2:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cbi = 0 To Max - Ca - 2
            NoOfComb = NoOfComb + 1
            comb = char(Cbi) & char(Cbi + Ca + 1)
            If frmMain.optCPMode.Value = True Then
                        permute (comb)
            Else
                Select Case OChoice
                Case 1:
                    result = result & vbCrLf & comb
                Case 2:
                    Print #1, comb
                Case 3:
                    SendKeys comb: BFInitialise
                Case 4:
                processcode (comb)
                
                End Select
                'NoOfComb = NoOfComb + 1
                'Debug.Print comb, NoOfComb
            End If
        Next Cbi
    Next Ca
Case 3:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cb = 0 To Max - Ca
            For Cbi = 0 To Max - Ca - Cb - 3
                OpenForms = DoEvents
                NoOfComb = NoOfComb + 1
                comb = char(Cbi) & char(Cbi + Ca + 1) & char(Cbi + Ca + Cb + 2)
                If frmMain.optCPMode.Value = True Then
                        permute (comb)
                    Else
                        Select Case OChoice
                        Case 1:
                            result = result & vbCrLf & comb
                        Case 2:
                            Print #1, comb
                        Case 3:
                            SendKeys comb: BFInitialise
                        Case 4:
                            processcode (comb)
                        
                        End Select
                End If
                'NoOfComb = NoOfComb + 1
                'Debug.Print comb, NoOfComb
            Next Cbi
        Next Cb
    Next Ca
Case 4:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cb = 0 To Max - Ca
            For Cc = 0 To Max - Ca - Cb
                OpenForms = DoEvents
                For Cbi = 0 To Max - Ca - Cb - Cc - 4
                    NoOfComb = NoOfComb + 1
                    comb = char(Cbi) & char(Cbi + Ca + 1) & char(Cbi + Ca + Cb + 2) & char(Cbi + Ca + Cb + Cc + 3)
                    If frmMain.optCPMode.Value = True Then
                        permute (comb)
                    Else
                        Select Case OChoice
                        Case 1:
                            result = result & vbCrLf & comb
                        Case 2:
                            Print #1, comb
                        Case 3:
                            SendKeys comb: BFInitialise
                        Case 4:
                            processcode (comb)
                        End Select
                    End If
                    'NoOfComb = NoOfComb + 1
                    'Debug.Print comb, NoOfComb
                Next Cbi
            Next Cc
        Next Cb
    Next Ca
Case 5:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cb = 0 To Max - Ca
            For Cc = 0 To Max - Ca - Cb
                OpenForms = DoEvents
                For Cd = 0 To Max - Ca - Cb - Cc
                    For Cbi = 0 To Max - Ca - Cb - Cc - Cd - 5
                    OpenForms = DoEvents
                    NoOfComb = NoOfComb + 1
                    comb = char(Cbi) & char(Cbi + Ca + 1) & char(Cbi + Ca + Cb + 2) & char(Cbi + Ca + Cb + Cc + 3) & char(Cbi + Ca + Cb + Cc + Cd + 4)
                    If frmMain.optCPMode.Value = True Then
                        permute (comb)
                    Else
                        Select Case OChoice
                        Case 1:
                            result = result & vbCrLf & comb
                        Case 2:
                            Print #1, comb
                        Case 3:
                            SendKeys comb: BFInitialise
                        Case 4:
                            processcode (comb)
                        End Select
                    End If
                        'NoOfComb = NoOfComb + 1
                        'Debug.Print comb, NoOfComb
                    Next Cbi
                Next Cd
            Next Cc
        Next Cb
    Next Ca
Case 6:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cb = 0 To Max - Ca
            For Cc = 0 To Max - Ca - Cb
                OpenForms = DoEvents
                For Cd = 0 To Max - Ca - Cb - Cc
                    For Ce = 0 To Max - Ca - Cb - Cc - Cd
                        OpenForms = DoEvents
                        For Cbi = 0 To Max - Ca - Cb - Cc - Cd - Ce - 6
                            NoOfComb = NoOfComb + 1
                            comb = char(Cbi) & char(Cbi + Ca + 1) & char(Cbi + Ca + Cb + 2) & char(Cbi + Ca + Cb + Cc + 3) & char(Cbi + Ca + Cb + Cc + Cd + 4) & char(Cbi + Ca + Cb + Cc + Cd + Ce + 5)
                            If frmMain.optCPMode.Value = True Then
                                permute (comb)
                            Else
                                Select Case OChoice
                                Case 1:
                                    result = result & vbCrLf & comb
                                Case 2:
                                    Print #1, comb
                                Case 3:
                                    SendKeys comb: BFInitialise
                                Case 4:
                                    processcode (comb)
                                End Select
                                'NoOfComb = NoOfComb + 1
                                'Debug.Print comb, NoOfComb
                            End If
                        Next Cbi
                    Next Ce
                Next Cd
            Next Cc
        Next Cb
    Next Ca
Case 7:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cb = 0 To Max - Ca
            For Cc = 0 To Max - Ca - Cb
                OpenForms = DoEvents
                For Cd = 0 To Max - Ca - Cb - Cc
                    For Ce = 0 To Max - Ca - Cb - Cc - Cd
                        OpenForms = DoEvents
                        For Cf = 0 To Max - Ca - Cb - Cc - Cd - Ce
                            For Cbi = 0 To Max - Ca - Cb - Cc - Cd - Ce - Cf - 7
                                OpenForms = DoEvents
                                NoOfComb = NoOfComb + 1
                                comb = char(Cbi) & char(Cbi + Ca + 1) & char(Cbi + Ca + Cb + 2) & char(Cbi + Ca + Cb + Cc + 3) & char(Cbi + Ca + Cb + Cc + Cd + 4) & char(Cbi + Ca + Cb + Cc + Cd + Ce + 5) & char(Cbi + Ca + Cb + Cc + Cd + Ce + Cf + 6)
                                If frmMain.optCPMode.Value = True Then
                                    permute (comb)
                                Else
                                    Select Case OChoice
                                    Case 1:
                                        result = result & vbCrLf & comb
                                    Case 2:
                                        Print #1, comb
                                    Case 3:
                                        SendKeys comb: BFInitialise
                                    Case 4:
                                        processcode (comb)
                                    End Select
                                    'NoOfComb = NoOfComb + 1
                                    'Debug.Print comb, NoOfComb
                                End If
                            Next Cbi
                        Next Cf
                    Next Ce
                Next Cd
            Next Cc
        Next Cb
    Next Ca
Case 8:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cb = 0 To Max - Ca
            For Cc = 0 To Max - Ca - Cb
                OpenForms = DoEvents
                For Cd = 0 To Max - Ca - Cb - Cc
                    For Ce = 0 To Max - Ca - Cb - Cc - Cd
                        OpenForms = DoEvents
                        For Cf = 0 To Max - Ca - Cb - Cc - Cd - Ce
                            For Cg = 0 To Max - Ca - Cb - Cc - Cd - Ce - Cf
                                For Cbi = 0 To Max - Ca - Cb - Cc - Cd - Ce - Cf - Cg - 8
                                    OpenForms = DoEvents
                                    NoOfComb = NoOfComb + 1
                                    comb = char(Cbi) & char(Cbi + Ca + 1) & char(Cbi + Ca + Cb + 2) & char(Cbi + Ca + Cb + Cc + 3) & char(Cbi + Ca + Cb + Cc + Cd + 4) & char(Cbi + Ca + Cb + Cc + Cd + Ce + 5) & char(Cbi + Ca + Cb + Cc + Cd + Ce + Cf + 6) & char(Cbi + Ca + Cb + Cc + Cd + Ce + Cf + Cg + 7)
                                    If frmMain.optCPMode.Value = True Then
                                        permute (comb)
                                    Else
                                        Select Case OChoice
                                        Case 1:
                                            result = result & vbCrLf & comb
                                        Case 2:
                                            Print #1, comb
                                        Case 3:
                                            SendKeys comb: BFInitialise
                                        Case 4:
                                            processcode (comb)
                                        End Select
                                        'NoOfComb = NoOfComb + 1
                                        'Debug.Print comb, NoOfComb
                                    End If
                                Next Cbi
                            Next Cg
                        Next Cf
                    Next Ce
                Next Cd
            Next Cc
        Next Cb
    Next Ca
Case 9:
    For Ca = 0 To Max
        OpenForms = DoEvents
        For Cb = 0 To Max - Ca
            For Cc = 0 To Max - Ca - Cb
                For Cd = 0 To Max - Ca - Cb - Cc
                    For Ce = 0 To Max - Ca - Cb - Cc - Cd
                        For Cf = 0 To Max - Ca - Cb - Cc - Cd - Ce
                            For Cg = 0 To Max - Ca - Cb - Cc - Cd - Ce - Cf
                                For Ch = 0 To Max - Ca - Cb - Cc - Cd - Ce - Cf - Cg
                                    For Cbi = 0 To Max - Ca - Cb - Cc - Cd - Ce - Cf - Cg - Ch - 9
                                        OpenForms = DoEvents
                                        NoOfComb = NoOfComb + 1
                                        comb = char(Cbi) & char(Cbi + Ca + 1) & char(Cbi + Ca + Cb + 2) & char(Cbi + Ca + Cb + Cc + 3) & char(Cbi + Ca + Cb + Cc + Cd + 4) & char(Cbi + Ca + Cb + Cc + Cd + Ce + 5) & char(Cbi + Ca + Cb + Cc + Cd + Ce + Cf + 6) & char(Cbi + Ca + Cb + Cc + Cd + Ce + Cf + Cg + 7) & char(Cbi + Ca + Cb + Cc + Cd + Ce + Cf + Cg + Ch + 8)
                                        If frmMain.optCPMode.Value = True Then
                                            permute (comb)
                                        Else
                                            Select Case OChoice
                                            Case 1:
                                                result = result & vbCrLf & comb
                                            Case 2:
                                                Print #1, comb
                                            Case 3:
                                                SendKeys comb: BFInitialise
                                            Case 4:
                                                processcode (comb)
                                            End Select
                                            'NoOfComb = NoOfComb + 1
                                            'Debug.Print comb, NoOfComb
                                        End If
                                    Next Cbi
                                Next Ch
                            Next Cg
                        Next Cf
                    Next Ce
                Next Cd
            Next Cc
        Next Cb
    Next Ca

End Select
frmMain.txtStatus.Text = result
'Debug.Print Len(result) / (cho + 2)
Debug.Print "Total No of combinations processed"; NoOfComb
Debug.Print StartTime, "Started"
Debug.Print Time, "Ended"
Debug.Print "Over"


End Sub

