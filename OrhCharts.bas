Attribute VB_Name = "OrhCharts"
Sub CreateBrandChart()
    Application.ScreenUpdating = False
    
    Dim oSALayout As SmartArtLayout
    Dim QNode As SmartArtNode
    Dim AsmNode As SmartArtNode
    Dim RepNode As SmartArtNode
    Dim QNodes As SmartArtNodes
    Dim PID As String
    Dim Employees As New Collection
    Dim Emp As CEmployee
    Dim drEmp As CEmployee
    Dim AsmEmp As CEmployee
    Dim RepEmp As CEmployee
    
    'set columns for attributes
    EmpC = Lib.getCol(1, "Employee")
    ChC = Lib.getCol(1, "Chief")
    StC = Lib.getCol(1, "Status")
    VacC = Lib.getCol(1, "Vacancy")
    RoleC = Lib.getCol(1, "Role")
    MixtC = Lib.getCol(1, "Specialization")
    lRow = Lib.lastRow
    Cac = Lib.getCol(1, "CA")
    
    sheetNames = Array("DR", "ASM", "REP")
    brandNames = Array("LP", "MX", "KR", "RD", "ES")
    regNames = Array("Moscou GR", "Nord-Ouest", "Centre", "Sud", "Oural", "Siberie", "EO")
    Set wb = ActiveWorkbook
    tmonth = CInt(InputBox("Insert actual month"))
    Dim curDate As Long
    
    curDate = DateSerial(2016, tmonth, 1)
    
    For Each mReg In regNames
        Set orgCharts = Workbooks.Add
        orgCharts.Sheets(1).name = "temp"
        For Each brand In brandNames
            wb.Activate
            For Each sheetName In sheetNames
                Sheets(sheetName).Activate
                Range("1:1").Select
                For Each fil In ActiveSheet.AutoFilter.Filters
                    If fil.On Then ActiveSheet.ShowAllData: Exit For
                Next fil
                Selection.AutoFilter Field:=Lib.getCol(1, "Date"), Criteria1:=curDate
                Selection.AutoFilter Field:=Lib.getCol(1, "Brand"), Criteria1:=brand
                Selection.AutoFilter Field:=Lib.getCol(1, "mReg"), Criteria1:=mReg
                lRow = Lib.lastRow
                For i = 2 To lRow
                    If Rows(i).EntireRow.Hidden = False Then
                        If Worker.checkName(Cells(i, EmpC).Value, Employees) = False Then
                            Set Emp = New CEmployee
                            With Emp
                                .name = Cells(i, EmpC).Value
                                .Chief = Cells(i, ChC).Value
                                .Dir = Cells(i, StC).Value
                                .Role = Cells(i, RoleC).Value
                                .Spec = Cells(i, MixtC).Value
                                .CA = Cells(i, Cac).Value
                            End With
                            Employees.Add Emp
                        End If
                    End If
                Next i
            Next sheetName
        Next brand
   Next mReg
    
    
End Sub




Sub CreateChartHC()
    
    Application.ScreenUpdating = False
    
    Dim oSALayout As SmartArtLayout
    Dim QNode As SmartArtNode
    Dim AsmNode As SmartArtNode
    Dim RepNode As SmartArtNode
    Dim QNodes As SmartArtNodes
    Dim PID As String
    Dim Employees As New Collection
    Dim Emp As CEmployee
    Dim drEmp As CEmployee
    Dim AsmEmp As CEmployee
    Dim RepEmp As CEmployee
    
    Set wSheet = Sheets("hcData")
    wSheet.Activate
    
    'set columns for attributes
    EmpC = Lib.getCol(1, "Employee")
    ChC = Lib.getCol(1, "Chief")
    StC = Lib.getCol(1, "Status")
    VacC = Lib.getCol(1, "Vacancy")
    RoleC = Lib.getCol(1, "Role")
    MixtC = Lib.getCol(1, "Specialization")
    lRow = Lib.lastRow
    Cac = Lib.getCol(1, "CA")
    
    'fill collection of Employees from sheet "hcData"
    For i = 2 To lRow
        Set Emp = New CEmployee
        With Emp
            .name = Cells(i, EmpC).Value
            .Chief = Cells(i, ChC).Value
            .Dir = Cells(i, StC).Value
            .Role = Cells(i, RoleC).Value
            .Spec = Cells(i, MixtC).Value
            .CA = Cells(i, Cac).Value
        End With
        Employees.Add Emp
    Next i
    
    Set orgCharts = Workbooks.Add
    orgCharts.Sheets(1).name = "temp"
    
    'Create sheet for each DR with his orgStructure from Employees, using SmartArt Objects
    For i = 1 To lRow - 1
        Set drEmp = Employees.Item(i)
        If drEmp.Chief = "" Then
            orgCharts.Worksheets.Add.name = drEmp.name
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.DisplayHeadings = False
            Set oSALayout = Application.SmartArtLayouts(88)
            Set oShp = ActiveWorkbook.ActiveSheet.Shapes.AddSmartArt(oSALayout, 1, 1, 680, 510)
            Set QNodes = oShp.SmartArt.AllNodes
            For n = 1 To 5
                oShp.SmartArt.AllNodes(1).Delete
            Next n
            Set QNode = oShp.SmartArt.AllNodes.Add
                
            setNode drEmp, QNode
                
            For j = i + 1 To lRow - 1
                Set AsmEmp = Employees.Item(j)
                If AsmEmp.Chief = drEmp.name Then
                    Set AsmNode = QNode.AddNode(msoSmartArtNodeBelow)
                    setNode AsmEmp, AsmNode
                    For p = j + 1 To lRow - 1
                        Set RepEmp = Employees.Item(p)
                        If RepEmp.Chief = AsmEmp.name Then
                            Set RepNode = AsmNode.AddNode(msoSmartArtNodeBelow)
                            setNode RepEmp, RepNode
                        End If
                    Next p
                End If
            Next j
            fixPageSettings
        End If
    Next i
      
    Sheets("temp").Delete
    Application.ScreenUpdating = True
    
End Sub


Sub setNode(employee As CEmployee, node As SmartArtNode)
    node.TextFrame2.TextRange.text = employee.name & Chr(10) & employee.Spec & Chr(10) & employee.Role & " " & employee.Dir & Chr(10) & "CA = " & employee.CA
    node.TextFrame2.TextRange.Font.name = "Times New Roman"
    If employee.name Like "*акансия*" Then
        node.Shapes(1).Fill.ForeColor.RGB = RGB(190, 190, 190)
    Else
        Select Case employee.Role
            Case "DR"
                node.Shapes(1).Fill.ForeColor.RGB = RGB(200, 40, 40)
            Case "ASM"
                Select Case employee.Dir
                    Case "Partner"
                        node.Shapes(1).Fill.ForeColor.RGB = RGB(40, 40, 200)
                    Case "Ancor"
                        node.Shapes(1).Fill.ForeColor.RGB = RGB(80, 80, 200)
                    Case "Direct"
                        node.Shapes(1).Fill.ForeColor.RGB = RGB(80, 120, 200)
                End Select
            Case "REP"
                Select Case employee.Dir
                    Case "Partner"
                        node.Shapes(1).Fill.ForeColor.RGB = RGB(49, 153, 49)
                    Case "Ancor"
                        node.Shapes(1).Fill.ForeColor.RGB = RGB(43, 172, 130)
                    Case "Direct"
                        node.Shapes(1).Fill.ForeColor.RGB = RGB(122, 213, 40)
                    Case "Intern"
                        node.Shapes(1).Fill.ForeColor.RGB = RGB(43, 172, 130)
                End Select
        End Select
    End If
    If Len(employee.Spec) > 2 Then
        node.Shapes(1).Glow.Color.RGB = RGB(150, 150, 40)
    End If
End Sub

Sub sendOrgChartToDrs()
    Set empl = ActiveWorkbook.Sheets("EMPLOYEES")
    s = Lib.selectFile
    Set wFile = Workbooks.Open(s)
    For Each wSheet In wFile.Sheets
        wSheet.Copy
        NFD = "\\rucorprufil2\LOREAL\DPP\Business development\MANCOM\Структура\COM\2016\" & wSheet.name & ".xlsx"
        ActiveWorkbook.SaveAs NFD
        ActiveWorkbook.Close
        drMail = takeMail(wSheet.name)
        
        Set OutlookApp = CreateObject("Outlook.Application")
            Set oMail = OutlookApp.CreateItem(0)
            With oMail
                .To = drMail
                .Importance = 2
                .Subject = "Организационная диаграмма"
                .Body = "Файл во вложении"
                .Attachments.Add NFD
                .Send
            End With
        Kill NFD

    Next wSheet

End Sub

Function takeMail(drName As String) As String
    Dim awb As Workbook
    Set awb = ActiveWorkbook
    Dim result As String
    result = ""
    Workbooks("Structure Template.xlsm").Sheets("EMPLOYEES").Activate
    For i = 2 To Lib.lastRow
        If Cells(i, Lib.getCol(1, "Employee")).Value = drName Then
            result = Cells(i, getCol(1, "Mail")).Value
            Exit For
        End If
    Next i
    awb.Activate
    takeMail = result
End Function

Sub fixPageSettings()
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Orientation = xlLandscape
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
    End With
End Sub
