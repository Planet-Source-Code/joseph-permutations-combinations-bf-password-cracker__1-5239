Attribute VB_Name = "PowerUser"
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
' or planet@jofu.8m.com
'Web Site http://www.jofu.8m.com

Public PUFlag As Integer
Public StartTime As Variant

Public Sub processcode(PCCode As String)
If PUFlag = 0 Then
    PUFlag = 1
    a = MsgBox("This is a demonstration of how the permuted word can be used internally in code to check out for possible password combination.Here when ever the permuted word is JOSEPH the code breaks and the thus you can determine the internal speed. All other external time delay as seen in the other bruteforce attack has been eliminated here", vbInformation, "Bruteforce for PowerUsers")
End If
If PCCode = "JOSEPH" Then
    Debug.Print PCCode
    a = MsgBox("Password cracked & code is " & PCCode & vbCrLf & "No of permutations tested " & pcount & vbCrLf & "Job started at " & StartTime & vbCrLf & "Job ended at " & Time, vbExclamation, "Bruteforce Power User Module")
End If

End Sub
