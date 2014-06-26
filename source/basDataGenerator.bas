Attribute VB_Name = "basDataGenerator"
'------------------------------------------------------------------------
' Description  : This class contains all specific procedures
'                   required for data generation
'------------------------------------------------------------------------

'Declarations

'Constants

'options
Option Explicit
'-------------------------------------------------------------
' Description   : starts data generation
' Parameter     :
'-------------------------------------------------------------
Public Sub generateTraffic()

    Dim wshTrafficData As Worksheet
    Dim intCustomerId As Integer
    Dim intProductsPerVisit As Integer
    Dim intVisit As Integer                 'ZŠhler fŸr den Aufruf einer Produktseite
    Dim intSelectedProductType As Integer
    Dim intLastProductType As Integer
    Dim intSelectedColor As Integer
    Dim intColor As Integer                 'eine von 16 Farben
    Dim lngProductId As Long                'ProduktId = Typ * 16 + Farboffset (1-16)
    Dim rngCurrent As Range
    Dim colBasket As Collection
    Dim intBasketCount As Integer
    Dim strRequest As String
    Dim strQuantities As String
    Dim strItemIds As String
   
    On Error GoTo error_handler
    Randomize
    Set wshTrafficData = basDataGenerator.p_createSheet()
    Set rngCurrent = wshTrafficData.Range("A1")
    
    'Schleife Ÿber alle unique vistors
    For intCustomerId = 1 To 8000
        'Warenkorb leeren
        Set colBasket = New Collection
        'per Zufall bestimmen wieviel Produkte angesehen werden
        intProductsPerVisit = CInt(Rnd * 5) + 1
        intLastProductType = -1
        For intVisit = 1 To intProductsPerVisit
            'Produkt aussuchen, dabei zuletzt gewŠhlten
            intSelectedProductType = basDataGenerator.p_chooseProduct
            'Farbe aussuchen und mit vorgegebener Wahrscheinlichkeit fŸr nŠchstes Produkt beibehalten
            If intVisit = 1 Or Rnd > 0.8 Or intLastProductType = intSelectedProductType Then
                intColor = CInt(Rnd * 16) + 1
            End If
            'ProduktId berechnen = Typ * 16 + Farboffset (1-16)
            lngProductId = (intSelectedProductType * 16) + intColor
            'jedes 10.Produkt und jeweils letztes Produkt eines Besuchs in den Warekorb legen
            If Rnd > 0.9 Or intVisit = intProductsPerVisit Then
                colBasket.Add lngProductId
            End If
            'gerade gewŠhlten Prorukttyp merken
            intLastProductType = intSelectedProductType
            'Daten ausgeben
            strRequest = "http://192.168.1.30:8080/rde_server/res/19c415c38907/recomm/ADS/sid/visitor" & _
                intCustomerId & "?sku=" & lngProductId
            rngCurrent.Value = strRequest
            Set rngCurrent = rngCurrent.Offset(1)
        Next
        'jeder 30. Kunde kauft
        If intCustomerId Mod 30 = 0 Then
            strQuantities = ""
            strItemIds = ""
            For intBasketCount = 1 To colBasket.Count
                If strQuantities = "" Then
                    strQuantities = "1"
                Else
                    strQuantities = strQuantities & ",1"
                End If
                If strItemIds = "" Then
                    strItemIds = colBasket.Item(intBasketCount)
                Else
                    strItemIds = strItemIds & "," & colBasket.Item(intBasketCount)
                End If
            Next
            strRequest = "http://192.168.1.30:8080/rde_server/res/19c415c38907/event/order/sid/visitor" & _
                  intCustomerId & "?quantities=" & strQuantities & "&itemids=" & strItemIds
            rngCurrent.Value = strRequest
            Set rngCurrent = rngCurrent.Offset(1)
        End If
        'Warenkorb verwerfen
        Set colBasket = Nothing
    Next
    Exit Sub

error_handler:
    basSystem.log_error "basDataGenerator.generateTraffic"
End Sub
'-------------------------------------------------------------
' Description   : Produkt zufŠllig auswŠhlen
' Parameter     :
'-------------------------------------------------------------
Private Function p_chooseProduct() As Integer

    Dim varProducts As Variant
    Dim varProductLikeness As Variant
    Dim intProductType As Integer           'Index fŸr Produkttyp
    Dim intSelectedType As Integer
    Dim dblLikeness As Double
    Dim dblMaxLikeness As Double

    On Error GoTo error_handler
    varProducts = Array("Quadrat", "Kreis", "Dreieck", "Pentagon", "Stern")
    varProductLikeness = Array(0.9, 0.7, 0.5, 0.4, 0.3)
    dblMaxLikeness = 0
    For intProductType = 0 To 4
        dblLikeness = Rnd * varProductLikeness(intProductType)
        If dblLikeness > dblMaxLikeness Then
           dblMaxLikeness = dblLikeness
           intSelectedType = intProductType
        End If
    Next
    p_chooseProduct = intSelectedType
    Exit Function

error_handler:
    basSystem.log_error "basDataGenerator.p_chooseProduct"
End Function
'------------------------------------------------------------------------
' Description  : create a new sheet containing the data for simulation
' Parameters   :
' Returnvalue  : worksheet object interfacing the chart container
'------------------------------------------------------------------------
Private Function p_createSheet()

    Dim wshSheet As Worksheet               'new sheet
    Dim wshCurrentSheet As Worksheet        'an existing sheet in the active workbook
    Dim strSheetNumber As String            'number of the current treemap sheet as text
    Dim blnFoundNonNumber As Boolean        'flag is true when a treemap sheet name contains
                                            ' characters other then treemap and a number
    Dim intPosition As Integer              'counts characters
    Dim intNewNumber As Integer             'the number of the the new sheet
    Dim intDefaultLen As Integer            ' length of default name

    On Error GoTo error_handler
    intNewNumber = 1
    intDefaultLen = Len(cLangSheetName) + 1
    'create a new sheet in active workbook
    Set wshSheet = ActiveWorkbook.Worksheets.Add(After:=ActiveSheet)
    basSystem.log "new sheet added"
    'set the name to treemap + number, where number is the highest not existing
    ' number for treemap sheets
    For Each wshCurrentSheet In ActiveWorkbook.Worksheets
        'looking for originally named treemap sheets
        If Left(wshCurrentSheet.Name, Len(cLangSheetName)) = cLangSheetName Then
            'try to find the number from the rest of the name
            strSheetNumber = Mid(wshCurrentSheet.Name, intDefaultLen)
            'look for non number characters
            blnFoundNonNumber = False
            For intPosition = intDefaultLen To Len(wshCurrentSheet.Name)
                'wish I could use regex, instead have to check ascii codes to find
                ' non number characters
                If Asc(Mid(wshCurrentSheet.Name, intPosition, 1)) < 48 Or _
                        Asc(Mid(wshCurrentSheet.Name, intPosition, 1)) > 57 Then
                    blnFoundNonNumber = True
                End If
            Next
            'if name of the current sheet is a default name
            If Not blnFoundNonNumber And Len(strSheetNumber) > 0 Then
                'give the new sheet a higher number
                If CInt(strSheetNumber) >= intNewNumber Then
                    intNewNumber = CInt(strSheetNumber) + 1
                End If
            End If
        End If
    Next
    'set the default name for the new sheet
    wshSheet.Name = cLangSheetName & intNewNumber
    wshSheet.Activate
    basSystem.log "new sheet named to " & cLangSheetName & intNewNumber
    'return the sheets object
    Set p_createSheet = wshSheet
    Exit Function
    
error_handler:
    basSystem.log_error "basDataGenerator.p_createSheet"
End Function

