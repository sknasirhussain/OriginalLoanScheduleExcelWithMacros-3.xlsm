Option Explicit

Private Sub cmdCalculate_Click()
    Dim ws As Worksheet
    Dim amount As Double
    Dim tenure As Long
    Dim interestRate As Double
    Dim numberOfPayments As Integer
    Dim paymentDate As Date
    Dim beginningBalance As Double
    Dim payment As Double
    Dim principal As Double
    Dim interest As Double
    Dim endingBalance As Double
    Dim pmtNo As Integer

    Application.DisplayAlerts = False
    
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set conn = New ADODB.Connection
    conn.Open "driver={sql server};server=LINKOLFEB24-092;database=db;uid=admin;pwd=admin123;"
    Debug.Print "Connected successfully"
    
    Dim years As Long
    Dim rate As Double
    Dim sql As String
    
    tenure = tbxDuration
    
    If tenure <= 1 Then
        years = 1
    ElseIf tenure >= 2 And tenure <= 5 Then
        years = 5
    ElseIf tenure >= 6 Then
        years = 10
    End If
    
    sql = "Select Rate from rates where Duration =" & years
    Set rs = conn.Execute(sql)
    Debug.Print rs.Fields("Rate").Value
    rate = rs.Fields("Rate").Value / 100
    
    Set ws = ActiveSheet
    
    ws.Range("A2").Value = "Amount"
    ws.Range("A3").Value = "Tenure"
    ws.Range("A4").Value = "Rate of Interest"
    ws.Range("A5").Value = "Number of payments"
    
    ws.Range("D2").Value = "Monthly payment"
    ws.Range("D3").Value = "Number of payments"
    ws.Range("D4").Value = "Total interest"
    ws.Range("D5").Value = "Total cost of loan"
    
    amount = tbxAmount
    interestRate = rate
    numberOfPayments = tbxNumberOfPayments
    
    ws.Range("B2").Value = amount
    ws.Range("B3").Value = tenure
    ws.Range("B4").Value = rate * 100
    ws.Range("B5").Value = numberOfPayments
    
    paymentDate = DateSerial(Year(Date), Month(Date), 1)
    beginningBalance = amount
    pmtNo = 1

    ws.Range("A8:G1000").ClearContents

    payment = Round(WorksheetFunction.Pmt((interestRate / numberOfPayments) / 100, tenure * numberOfPayments, -amount), 2)
    
    
    ' Write the header row
    ws.Cells(7, 1).Value = "Pmt No."
    ws.Cells(7, 2).Value = "Payment Date"
    ws.Cells(7, 3).Value = "Beginning Balance"
    ws.Cells(7, 4).Value = "Payment"
    ws.Cells(7, 5).Value = "Principal"
    ws.Cells(7, 6).Value = "Interest"
    ws.Cells(7, 7).Value = "Ending Balance"
    
    Do While beginningBalance > 0 And pmtNo <= tenure * numberOfPayments
        interest = Round((beginningBalance * interestRate / numberOfPayments) / 100, 2)
        
        principal = Round(payment - interest, 2)
        
        ' Calculate ending balance
        endingBalance = Round(beginningBalance - principal, 2)
        
        ' Fill in data in the table
        ws.Cells(pmtNo + 7, 1).Value = pmtNo
        ws.Cells(pmtNo + 7, 2).Value = Format(paymentDate, "dd-mm-yyyy")
        ws.Cells(pmtNo + 7, 3).Value = beginningBalance
        ws.Cells(pmtNo + 7, 4).Value = payment
        ws.Cells(pmtNo + 7, 5).Value = principal
        ws.Cells(pmtNo + 7, 6).Value = interest
        ws.Cells(pmtNo + 7, 7).Value = endingBalance
        
        ' Update variables for next iteration
        beginningBalance = endingBalance
        paymentDate = DateAdd("m", 1, paymentDate)
        pmtNo = pmtNo + 1
    Loop
    
    ' Adjust column widths
    Dim col As Integer
    For col = 1 To 7
        ws.Columns(col).AutoFit
    Next col
    
    ' Adjust row heights
    Dim row As Integer
    For row = 7 To pmtNo + 6
        ws.Rows(row).AutoFit
    Next row

    ws.Range("E2").Value = Format(payment, "0.00")
    ws.Range("E3").Value = tenure * numberOfPayments
    ws.Range("E4").Value = Format(WorksheetFunction.Sum(ws.Range("F9:F1000")), "0.00")
    ws.Range("E5").Value = Format(amount + ws.Range("E4"), "0.00")
    ws.Range("A2").EntireRow.Insert Shift:=xlUp
    
    With ws.Range("B1:d1")
        .Select
        .Merge
        .Font.Size = 21
        .Value = "Loan amortization schedule"
        .Font.Bold = True
    End With
    
    Range("A3:A6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Range("D3:D6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Range("B3:B6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    Range("E3:E6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    Range("B1:D1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    Range("A8:G8").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    Application.DisplayAlerts = True
    
    Cells.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    
    Range("J11").Select
    
    Unload Me
End Sub


'Private Sub cmdCalculate_Click(amount As Double, duration As Long, numberOfPayments As Long)
'
'    Dim conn As New ADODB.Connection
'    Dim rs As New ADODB.Recordset
'    Dim cmd As New ADODB.Command
'    Dim Query As String
'
'    Dim server_name As String
'    Dim database_name As String
'    Dim user_id As String
'    Dim password As String
'
'    server_name = "LINRNCAUG23-272\SQLEXPRESS"
'    database_name = "BankDB"
'    user_id = "vbaadmin"
'    password = "qwerty"
'
'    conn.ConnectionString = _
'        "Provider =SQLOLEDB;" & _
'        "Data Source=" & server_name & ";" & _
'        "Initial Catalog=" & database_name & ";" & _
'        "User ID=" & user_id & ";" & _
'        "Password=" & password & ";"
'    conn.Open
'
'    cmd.ActiveConnection = conn
'    cmd.CommandType = adCmdText
'    Query = "select Rate from interestRateTbl Where Tenure = " & duration
'    cmd.CommandText = Query
'    Set rs = cmd.Execute
'    Debug.Print rs.Fields(0).Value
'    With ActiveSheet
'        .Range("E5").Value = amount
'        .Range("E6").Value = rs.Fields(0).Value / 100
'        .Range("E7").Value = duration
'        .Range("E8").Value = numberOfPayments
'    End With
'
'
'End Sub