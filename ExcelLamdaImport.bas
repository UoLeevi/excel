Attribute VB_Name = "ExcelLamdaImport"
Option Explicit

Sub AddLambdaNamedRangesFromFile( _
    ByVal lambdasToImport As Variant, _
    Optional ByVal filePathOrUrl As String, _
    Optional ByVal wb As Workbook)
    
    If filePathOrUrl = Empty Then filePathOrUrl = "https://raw.githubusercontent.com/UoLeevi/excel/main/README.md"

    If wb Is Nothing Then Set wb = ActiveWorkbook

    ' Step 1. Read the file contents or get text from URL
    Dim text As String

    If filePathOrUrl Like "http*://*" Then
        Dim httpRequest As Object
        Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
        httpRequest.Open "GET", filePathOrUrl, False
        httpRequest.Send
        text = httpRequest.ResponseText
    Else
        Dim fileNum As Integer
        fileNum = FreeFile
        Open filePathOrUrl For Input As #fileNum
        text = Input$(LOF(filePathOrUrl), fileNum)
        Close #fileNum
    End If

    ' Step 2. Replace CrLf with Lf
    text = Replace(text, vbCrLf, vbLf)

    Dim nameOfLambda As Variant
    Dim lambdaFormula As String
    Dim i, j, k As Long
    Dim nm As Name

    For Each nameOfLambda In lambdasToImport

        ' Step 3. Find lambda definition by name
        i = InStr(1, text, vbLf & "# " & nameOfLambda, vbTextCompare)
        If i = 0 Then Err.Raise 1001, , "Specified lambda '" & nameOfLambda & "' was not found"

        ' Step 4. Find the start of formula definition
        i = InStr(i, text, vbLf & "=LAMBDA(", vbTextCompare)

        ' Step 5. Find the end of formula definition marked by a blank line
        j = InStr(i, text, vbLf & vbLf, vbTextCompare)
        k = InStrRev(text, "```", j, vbTextCompare)
        If k > i Then j = k

        ' Step 6. Extract the formula text
        lambdaFormula = Trim(Mid(text, i + 1, j - i - 1))

        ' Step 7. Add or update named range

        On Error Resume Next
        Set nm = wb.Names(nameOfLambda)
        On Error GoTo 0

        If nm Is Nothing Then
            wb.Names.Add Name:=nameOfLambda, RefersTo:=lambdaFormula
        Else
            nm.RefersTo = lambdaFormula
            Set nm = Nothing
        End If
    Next

End Sub

