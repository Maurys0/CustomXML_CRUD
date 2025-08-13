VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomXMLCRUD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Base 0

Property Get xmlText(ByVal strRoot As String) as String
    Dim part As CustomXMLPart

    For Each part In ThisWorkbook.CustomXMLParts
        If InStr(1, part.xml, strRoot) > 0 Then
            xmlText = part.xml
            Exit Property
        Next
    Next part
    
    xmlText = ""
End Property

Property Get Existe(ByVal strRoot) As Boolean
    Dim part As CustomXMLPart
    
    For Each part In ThisWorkbook.CustomXMLParts
        If InStr(1, part.xml, strRoot) > 0 Then
            Existe = True
            Exit Property
        End If
    Next part
    
    Existe = False
End Property

Property Get Criar(ByVal strXML As String) As Boolean
    Dim strRoot As String
    
    strRoot = Split(strXML, ">")(0) & ">"
    
    If Not Existe(strRoot) Then
        ThisWorkbook.CustomXMLParts.Add strXML
        Criar = True
    Else
        Criar = False
    End If
End Property

Property Get Deletar(ByVal strRoot As String) As Boolean
    Dim part As CustomXMLPart
    Dim encontrado As Boolean
    
    For Each part In ThisWorkbook.CustomXMLParts
        If InStr(1, part.xml, strRoot) > 0 Then
            part.Delete
            Deletar = True
            Exit Property
        End If
    Next part
    
    Deletar = False
End Property



