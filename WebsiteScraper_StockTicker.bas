Option Explicit
Dim childNodes As IHTMLDOMChildrenCollection
Dim ie As InternetExplorer
Dim i As Integer
Dim n As Integer
Dim rowCount As Integer
Dim varYr As Variant
Dim sTicker As String
Dim sSection As String
Dim oITM As Object
Dim iChild As Integer
Dim iStatement As Integer
Dim vStatement As Variant

Function GetOpenIEByURL(ByVal i_URL As String) As SHDocVw.InternetExplorer
Dim objShellWindows As New SHDocVw.ShellWindows

  'ignore errors when accessing the document property
  On Error Resume Next
  'loop over all Shell-Windows
  For Each GetOpenIEByURL In objShellWindows
    'if the document is of type HTMLDocument, it is an IE window
    If TypeName(GetOpenIEByURL.Document) = "HTMLDocument" Then
      'check the URL
      If UCase(GetOpenIEByURL.Document.URL) Like UCase(i_URL) Then
        'leave, we found the right window
        Exit Function
      End If
    End If
  Next

End Function
Sub WebsiteScraper()

Sheets("WebsiteScraper").Select
Cells.Select
Selection.ClearContents
Range("A1").Select
Range("A1") = "Ticker"
Range("B1") = "Statement Type"
Range("C1") = "Section"
Range("D1") = "Period"
Range("E1") = "Value"
Range("F1") = "Insert_DT"

sTicker = Sheets("TickerList").Range("A2")
vStatement = Array("Income Statement", "Balance Sheet", "Cash Flow")
iStatement = 0
rowCount = 2
i = 0

ReDim varYr(i)
    
    
    
'create and open web page(not waiting long enough for page to fully render)
Set ie = CreateObject("InternetExplorer.Application")
ie.Visible = True
ie.Navigate "https://finance.website.com/quote/" & sTicker & "/financials?p=" & sTicker
Do While ie.ReadyState = 4: DoEvents: Loop
Do Until ie.ReadyState = 4: DoEvents: Loop
While ie.Busy
    DoEvents
Wend

'Application.Wait (Now + TimeValue("0:00:03"))

''look for web page
'Set ie = GetOpenIEByURL("https://finance.website.com/quote/" & sTicker & "/financials?p=" & sTicker)
'If ie Is Nothing Then
'    MsgBox "error"
'    End
'End If

    'loops through vStatement array which contains the list of financial statements
    While iStatement <= UBound(vStatement)
    On Error Resume Next
        
        'set i counter to 0
        i = 0
        
        'search through website
        For Each oITM In ie.Document.All
            
            'store the periods
            If InStr(LCase(oITM.ClassName), LCase("D(ib) Fw(b) Ta(end)")) > 0 Then
                ReDim Preserve varYr(i)
                varYr(i) = oITM.innerText
                i = i + 1
            End If
            
            'enter ticker, period, section, and amount from financials
            If LCase(oITM.ClassName) = LCase("D(tbr) fi-row Bgc($hoverBgColor):h") Then
                Set childNodes = oITM.childNodes
                For iChild = 0 To childNodes.Length
                    Range("A" & rowCount) = sTicker    'ticker
                    Range("B" & rowCount) = WorksheetFunction.Proper(vStatement(iStatement))   'statement type
                    Range("D" & rowCount) = varYr(iChild - 1)     'period
                    Select Case iChild
                    Case 0
                        sSection = childNodes(iChild).innerText
                    Case Else
                        Range("C" & rowCount) = sSection    'statement section
                        Range("E" & rowCount) = childNodes(iChild).innerText  'value
                    End Select
                    Range("F" & rowCount) = Now()
                    
                    rowCount = rowCount + 1
                Next
            End If
        Next oITM
        
        iStatement = iStatement + 1
    '    Debug.Print LCase(vStatement(iStatement))
        'search for balance sheet and click link
        For Each oITM In ie.Document.GetElementsByTagName("A")
    '        Debug.Print LCase(oITM.Text) = LCase(vStatement(iStatement))
            Select Case LCase(oITM.Text)
            Case LCase(vStatement(iStatement))
                oITM.Click
            End Select
        Next oITM
        
        Do While ie.ReadyState = 4: DoEvents: Loop
        Do Until ie.ReadyState = 4: DoEvents: Loop
        While ie.Busy
            DoEvents
        Wend
    Wend
    
    On Error GoTo 0
    ie.Quit
    Set ie = Nothing
    Cells.Select
    Selection.WrapText = False
    MsgBox ("done")
End Sub

Sub OpenInternetExplorer()

    sTicker = Sheets("TickerList").Range("A2")
    
    ''create and open web page(not waiting long enough for page to fully render)
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate "https://finance.website.com/quote/" & sTicker & "/financials?p=" & sTicker

End Sub
