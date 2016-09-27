Attribute VB_Name = "Worker"
Sub updateHeadCount()

    Application.ScreenUpdating = False

    Sheets("hcData").Activate
    s = "2:" & Lib.lastRow
    Rows(s).ClearContents

    Dim tmonth As Integer
    Dim Employees As New Collection
    Dim Emp As CEmployee
    Dim tname As String

    tmonth = CInt(InputBox("Insert actual month"))
    sheetNames = Array("DR", "ASM", "REP")

    For Each sheetName In sheetNames
        Sheets(sheetName).Activate
        For i = 2 To Lib.lastRow()
            If Month(Cells(i, 1).Value) = tmonth Then
                tname = Cells(i, Lib.getCol(1, "Employee")).Value
                If Not checkName(tname, Employees) Then
                    Set Emp = New CEmployee
                    With Emp
                        .name = tname
                        .Reg = Cells(i, Lib.getCol(1, "mReg")).Value
                        .Role = sheetName
                        .Dir = Cells(i, Lib.getCol(1, "Status")).Value
                        .Spec = Cells(i, Lib.getCol(1, "Specialization")).Value
                        .Chief = Cells(i, Lib.getCol(1, "Chief")).Value
                        .Sex = getParam("Sex", tname)
                        .Uid = getParam("ID", tname)
                        .Mail = getParam("Mail", tname)
                        .SubRole = getSubRole(.Uid)
                    End With
                    Employees.Add Emp
                End If
            End If
        Next i
    Next sheetName
    
    Sheets("hcData").Activate
    i = 2
    For Each empl In Employees
        Cells(i, 2).Value = empl.Reg
        Cells(i, 3).Value = empl.name
        Cells(i, 4).Value = empl.Dir
        Cells(i, 5).Value = empl.Spec
        Cells(i, 6).Value = empl.Chief
        If InStr(1, empl.name, "‚‡Í‡ÌÒ", vbTextCompare) > 0 Then Cells(i, 7).Value = 1
        Cells(i, 8).Value = empl.Role
        Cells(i, 9).Value = empl.Sex
        Cells(i, 10).Value = empl.Uid
        Cells(i, 11).Value = empl.Mail
        Cells(i, 12).Value = empl.SubRole
        i = i + 1
    Next empl
    
    Application.ScreenUpdating = True
    
End Sub

Sub updateRotation()

    Application.ScreenUpdating = False

    Sheets("rData").Activate
    s = "2:" & Lib.lastRow
    Rows(s).ClearContents

    Dim tmonth As Integer
    Dim Employees As New Collection
    Dim Emp As CEmployee
    Dim tname As String

    tmonth = CInt(InputBox("Insert actual month"))
    sheetNames = Array("DR", "ASM", "REP")
    counter = 2

    For curMonth = 1 To tmonth

        For Each sheetName In sheetNames
            Sheets(sheetName).Activate
            For i = 2 To Lib.lastRow()
                If Month(Cells(i, 1).Value) = curMonth Then
                    tname = Cells(i, Lib.getCol(1, "Employee")).Value
                    If Not checkName(tname, Employees) Then
                        Set Emp = New CEmployee
                        With Emp
                            .name = tname
                            .Reg = Cells(i, Lib.getCol(1, "mReg")).Value
                            .Role = sheetName
                            .Dir = Cells(i, Lib.getCol(1, "Status")).Value
                            .Spec = Cells(i, Lib.getCol(1, "Specialization")).Value
                            .Chief = Cells(i, Lib.getCol(1, "Chief")).Value
                            .Sex = getParam("Sex", tname)
                        End With
                        Employees.Add Emp
                    End If
                End If
            Next i
        Next sheetName
        
        Sheets("rData").Activate
        For Each empl In Employees
            Cells(counter, 2).Value = empl.Reg
            Cells(counter, 3).Value = empl.name
            Cells(counter, 4).Value = empl.Dir
            Cells(counter, 5).Value = empl.Spec
            Cells(counter, 6).Value = empl.Chief
            If empl.name Like "*‡Í‡ÌÒ*" Then Cells(counter, 7).Value = 1
            Cells(counter, 8).Value = empl.Role
            Cells(counter, 9).Value = empl.Sex
            Cells(counter, 1).Value = curMonth
            counter = counter + 1
        Next empl
        Set Employees = New Collection
    Next curMonth

    Cells(2, 11).FormulaLocal = "=≈—À»(»(G2="""";—◊®“≈—À»ÃÕ(C:C;C2;A:A;$A$1)=0;Ò˜∏ÚÂÒÎËÏÌ($C$1:C1;C2;$K$1:K1;1)=0);1;"""")"
    Cells(2, 12).FormulaLocal = "=≈—À»(Ë(G2="""";A2=$A$1);1;"""")"
    Range(Cells(2, 11), Cells(Lib.lastRow, 12)).FillDown
    Application.Calculate
    Range(Cells(2, 11), Cells(Lib.lastRow, 12)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False

    Application.ScreenUpdating = True

End Sub

Function getSubRole(Uid As String) As String
    Set ws = ActiveSheet
    Dim s As String
    s = ""
    Sheets("HR").Activate
    lRow = Lib.lastRow
    For i = 2 To lRow
        If Cells(i, 2).Value = Uid Then
            s = Cells(i, 19).Value
            Exit For
        End If
    Next i
    ws.Activate
    getSubRole = s
End Function


Function checkName(name As String, col As Collection) As Boolean
    result = False
    For Each Item In col
        If Item.name = name Then
            result = True
            Exit For
        End If
    Next Item
    checkName = result
End Function

Function getParam(param As String, name As String) As String
    Set ws = ActiveSheet
    Dim s As String
    s = ""
    Sheets("EMPLOYEES").Activate
    For i = 2 To Lib.lastRow
        If Cells(i, Lib.getCol(1, "Employee")).Value = name Then
            s = Cells(i, Lib.getCol(1, param)).Value
            Exit For
        End If
    Next i
    ws.Activate
    getParam = s
End Function

Function getSex(name As String) As String
    Set ws = ActiveSheet
    Dim s As String
    s = "Male"
    Sheets("EMPLOYEES").Activate
    For i = 2 To Lib.lastRow()
        If Cells(i, 1).Value = name Then
            s = Cells(i, 5).Value
            Exit For
        End If
    Next i
    ws.Activate
    getSex = s
End Function

Sub updateHr()
    Application.ScreenUpdating = False

    Dim s As String
    cNames = Array("Carol ID", "Local ID", "Last Name", "First Name", "Specialization Magnitude Code", "Carol Title", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    Set hrSheet = ActiveWorkbook.Sheets("HR")
    Set wb = Workbooks.Open(Lib.selectFile)
    wb.Activate
    Set fSheet = wb.Sheets("List Statutory Headcount")
    fSheet.Activate
    Cells.Select
    Selection.UnMerge
    cRow = Lib.getRow(1, "Carol ID")
    s = "1:" & cRow - 1
    Rows(s).Select
    Selection.Delete
    For Each cName In cNames
        s = cName
        Columns(Lib.getCol(1, s)).Select
        Selection.Copy
        hrSheet.Activate
        Cells(1, Lib.getCol(1, s)).Select
        Selection.PasteSpecial xlPasteValues
        fSheet.Activate
    Next cName
    
    Application.ScreenUpdating = False
End Sub

Sub modifyHr()
    Application.ScreenUpdating = False
    
    Sheets("HR").Activate
    Cells(2, Lib.getCol(1, "TYPE")).FormulaLocal = "=¬œ–(F2;JOBS!A:C;2;0)"
    Cells(2, Lib.getCol(1, "SUBTYPE")).FormulaLocal = "=¬œ–(F2;JOBS!A:C;3;0)"
    Cells(2, Lib.getCol(1, "NAME")).FormulaLocal = "=»Õƒ≈ —(EMPLOYEES!A:A;œŒ»— œŒ«(«Õ¿◊≈Õ(B2);EMPLOYEES!B:B;0))"
    Range(Cells(2, Lib.getCol(1, "TYPE")), Cells(Lib.lastRow, Lib.getCol(1, "NAME"))).FillDown
    Application.Calculate
    Range(Cells(2, Lib.getCol(1, "TYPE")), Cells(Lib.lastRow, Lib.getCol(1, "NAME"))).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
    
    Application.ScreenUpdating = True

End Sub

Sub modifySpecialization()
      'Application.ScreenUpdating = False
      
      Sheets("REP").Activate
      Cells(2, 15).FormulaLocal = "=≈—À»(—◊®“≈—À»ÃÕ($B:$B;O$1;$A:$A;$A2;$E:$E;$E2)>0;O$1;"""")"
      Cells(2, 16).FormulaLocal = "=≈—À»(»(O2<>"""";Q2&S2&U2&W2&Y2&AA2<>"""");""+"";"""")"
      Cells(2, 17).FormulaLocal = "=≈—À»(—◊®“≈—À»ÃÕ($B:$B;Q$1;$A:$A;$A2;$E:$E;$E2)>0;Q$1;"""")"
      Cells(2, 18).FormulaLocal = "=≈—À»(»(Q2<>"""";S2&U2&W2&Y2&AA2<>"""");""+"";"""")"
      Cells(2, 19).FormulaLocal = "=≈—À»(—◊®“≈—À»ÃÕ($B:$B;S$1;$A:$A;$A2;$E:$E;$E2)>0;S$1;"""")"
      Cells(2, 20).FormulaLocal = "=≈—À»(»(S2<>"""";U2&W2&Y2&AA2<>"""");""+"";"""")"
      Cells(2, 21).FormulaLocal = "=≈—À»(—◊®“≈—À»ÃÕ($B:$B;U$1;$A:$A;$A2;$E:$E;$E2)>0;U$1;"""")"
      Cells(2, 22).FormulaLocal = "=≈—À»(»(U2<>"""";W2&Y2&AA2<>"""");""+"";"""")"
      Cells(2, 23).FormulaLocal = "=≈—À»(—◊®“≈—À»ÃÕ($B:$B;W$1;$A:$A;$A2;$E:$E;$E2)>0;W$1;"""")"
      Cells(2, 24).FormulaLocal = "=≈—À»(»(W2<>"""";Y2&AA2<>"""");""+"";"""")"
      Cells(2, 25).FormulaLocal = "=≈—À»(—◊®“≈—À»ÃÕ($B:$B;Y$1;$A:$A;$A2;$E:$E;$E2)>0;Y$1;"""")"
      Cells(2, 26).FormulaLocal = "=≈—À»(»(Y2<>"""";AA2<>"""");""+"";"""")"
      Cells(2, 27).FormulaLocal = "=≈—À»(—◊®“≈—À»ÃÕ($B:$B;AA$1;$A:$A;$A2;$E:$E;$E2)>0;AA$1;"""")"
      Cells(2, 28).FormulaLocal = "=O2&P2&Q2&R2&S2&T2&U2&V2&W2&X2&Y2&Z2&AA2"
      Range(Cells(2, 15), Cells(Lib.lastRow, 28)).FillDown
      Application.Calculate
      Range(Cells(2, 15), Cells(Lib.lastRow, 28)).Select
      Selection.Copy
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
      SkipBlanks:=False, Transpose:=False
      
      Application.ScreenUpdating = True
      
End Sub

Sub fixSpace()
    
End Sub

Sub exportCA()
    Dim filePath As String
    Dim nameOfFile As String
    Dim colCA As Integer
    Dim colReg As Integer
    Dim colBr As Integer
    Dim colEmpl As Integer
    sheetNames = Array("DR", "ASM", "REP")
    filePath = Lib.selectFile
    Set struct = ActiveWorkbook
    Set wb = Workbooks.Open(filePath)
    For Each sheetName In sheetNames
        struct.Sheets(sheetName).Activate
        arr = Split(filePath, "\")
        nameOfFile = arr(UBound(arr))
        colCA = Lib.getCol(1, "CA")
        colReg = Lib.getCol(1, "mReg")
        colBr = Lib.getCol(1, "Brand")
        colEmpl = Lib.getCol(1, "Employee")
        If sheetName = "DR" Then
            Cells(2, colCA).FormulaLocal = "=Œ –”√À(—”ÃÃ≈—À»ÃÕ('[Top Russia Total DPP 2016.4.xlsb]PPD'!$BZ:$BZ;'[Top Russia Total DPP 2016.4.xlsb]PPD'!$D:$D;C2;'[Top Russia Total DPP 2016.4.xlsb]PPD'!$HX:$HX;B2)/4/1000;0)"

    Next sheetName
End Sub

Sub createFilesForUpdateAccess()
    Path = "\\rucorprufil2\LOREAL\DPP\Business development\MANCOM\—ÚÛÍÚÛ‡\COM\2016\temp\"
    cardsPath = "\\Rucorpruwks0665\cards\"
    comPath = "\\RUCORPRUWKS0665\For Regions Commercial Team\"
    s1 = "icacls " & Chr(34)
    s2 = Chr(34) & " /grant:r " & Chr(34)
    s3 = Chr(34) & ":(OI)(CI)F /t"
    fName = Path & returnDate & ".txt"
    Set fso = CreateObject("scripting.filesystemobject")
    Set ts = fso.createTextFile(fName, True, True)
    
    Sheets("hcData").Activate
    lRow = Lib.lastRow
    cId = Lib.getCol(1, "EmployeeId")
    cRole = Lib.getCol(1, "Role")
    cVac = Lib.getCol(1, "Vacancy")
    cStat = Lib.getCol(1, "Status")
    cMail = Lib.getCol(1, "Mail")
    cmReg = Lib.getCol(1, "mReg")
    cEmp = Lib.getCol(1, "Employee")

    For i = 2 To lRow
        If Cells(i, cVac).Value <> 1 And Cells(i, cStat).Value = "Direct" And Cells(i, cRole).Value <> "REP" Then
            If Cells(i, cRole).Value = "DR" Then
                ts.writeline s1 & cardsPath & Cells(i, cmReg).Value & s2 & Cells(i, cMail).Value & s3
                ts.writeline s1 & comPath & Cells(i, cmReg).Value & s2 & Cells(i, cMail).Value & s3
            Else
                ts.writeline s1 & cardsPath & Cells(i, cmReg).Value & "\" & Cells(i, cEmp).Value & s2 & Cells(i, cMail).Value & s3
                ts.writeline s1 & comPath & Cells(i, cmReg).Value & "\" & Cells(i, cEmp).Value & s2 & Cells(i, cMail).Value & s3
            End If
        End If
    Next i

End Sub

Function returnDate() As String
    Dim result As String
    result = ""
    
    tYear = CStr(Year(Now))
    tSecs = CStr(Second(Now))
    
    If Day(Now) > 9 Then
        tDay = CStr(Day(Now))
    Else
        tDay = "0" & CStr(Day(Now))
    End If
    
    If Month(Now) > 9 Then
        tmonth = CStr(Month(Now))
    Else
        tmonth = "0" & CStr(Month(Now))
    End If
    
    result = tYear & tmonth & tDay
    returnDate = CStr(result)
End Function

Sub deleteOldFolders()
    Sheets("hcData").Activate
    cmReg = Lib.getCol(1, "mReg")
    cEmp = Lib.getCol(1, "Employee")
    lRow = Lib.lastRow
    s1 = "rd " & Chr(34)
    s2 = Chr(34) & " " & "/" & "s " & "/" & "q"
    Path = "\\rucorprufil2\LOREAL\DPP\Business development\MANCOM\—ÚÛÍÚÛ‡\COM\2016\temp\"
    fName = Path & returnDate & ".txt"
    
    Set fso = CreateObject("scripting.filesystemobject")
    Set ts = fso.createTextFile(fName, True, True)

    Mregs = Array("Moscou GR", "Nord-Ouest", "Centre", "Sud", "Oural", "Siberie", "EO")
    
    For Each mReg In Mregs
        Path = "\\RUCORPRUWKS0665\Dropbox\For Regions Commercial Team\FLSM\" & mReg & "\"
        Set objFolder = fso.GetFolder(Path)
        For Each objFile In objFolder.SubFolders
            result = 0
            For i = 2 To lRow
                If Cells(i, cmReg).Value = mReg And Cells(i, cEmp).Value = objFile.name Then
                    result = 1
                    Exit For
                End If
            Next i
            If result = 0 Then
                ts.writeline s1 & objFile.Path & s2
            End If
        Next objFile
    Next mReg
    
End Sub

