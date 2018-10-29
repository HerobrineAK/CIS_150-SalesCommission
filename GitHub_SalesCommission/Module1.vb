Module Module1

    Private EmpName As String
    Private EmpNumber As String
    Private CurrentRecord() As String

    Private EmpSales As Decimal
    Private EmpCommission As Decimal

    Private Const CommissionRate As Decimal = 0.03
    Private SalesFile As New Microsoft.VisualBasic.FileIO.TextFieldParser("COMSALES.txt")
    Sub Main()
        Call HouseKeeping()
        Do While Not (SalesFile).EndOfData
            Call ProcessRecords()
        Loop
        Call EndOfJob()
    End Sub

    Sub HouseKeeping()
        Call SetFileDelimiters()
        Call WriteHeadings()
    End Sub

    Sub SetFileDelimiters()
        SalesFile.TextFieldType = FileIO.FieldType.Delimited
        SalesFile.SetDelimiters(",")
    End Sub

    Sub WriteHeadings()
        Console.WriteLine()
        Console.WriteLine(Space(30) & "Sales Commission Report")
        Console.WriteLine()
        Console.WriteLine(Space(5) & "Emp Number" & Space(10) & "Sales Person" & Space(12) & "Sales" & Space(6) & "Commission")
        Console.WriteLine()
    End Sub

    Sub ProcessRecords()
        Call ReadFile()
        Call DetailCalculation()
        Call WriteDetailLine()
    End Sub

    Sub ReadFile()
        CurrentRecord = SalesFile.ReadFields()

        EmpName = CurrentRecord(1)
        EmpNumber = CurrentRecord(0)
        EmpSales = CurrentRecord(2)
    End Sub

    Sub DetailCalculation()
        EmpCommission = EmpSales * CommissionRate
    End Sub

    Sub WriteDetailLine()
        Console.WriteLine(Space(5) & EmpNumber.PadLeft(11) & Space(9) & EmpName.PadRight(15) & Space(5) & EmpSales.ToString("c").PadLeft(9) & Space(10) & EmpCommission.ToString("N2").PadLeft(6))
    End Sub

    Sub EndOfJob()
        Call SummaryOutput()
        Call CloseFile()
    End Sub

    Sub SummaryOutput()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine(Space(33) & "End Of Sales Report")
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine(Space(32) & "Press -ENTER- To Exit")
    End Sub

    Sub CloseFile()
        SalesFile.Close()
        Console.ReadLine()
    End Sub
End Module
