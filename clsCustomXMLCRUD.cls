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

Public strXMLData As String

'Sub Class_Initialize(strRoot As String)
'    If Not Existe(strRoot) Then
'        Criar strRoot & Replace(strRoot, "<", "</")
'    End If
'
'    ArquivoAlvo strRoot
'End Sub

'Property Set ArquivoAlvo(strRoot As String)
'    Dim part As CustomXMLPart
'
'    For Each part In ThisWorkbook.CustomXMLParts
'        If InStr(1, part.xml, strRoot) > 0 Then
'            strXMLData = part.xml
'            Exit Property
'        Next
'    Next part
'End Property

Property Get Existe(Optional ByVal strRoot) As Boolean
    Dim part As CustomXMLPart
    
    For Each part In ThisWorkbook.CustomXMLParts
        If InStr(1, part.xml, strRoot) > 0 Then
            Existe = True
        End If
    Next part
End Property

Public Sub Criar(ByVal strXML As String)
    Dim strRoot As String
    
    strRoot = Split(strXML, ">")(0) & ">"
    
    If Not Existe(strRoot) Then
        ThisWorkbook.CustomXMLParts.Add strXML
    Else
        MsgBox "Já existe um arquivo XML " & strRoot & "!", vbOKOnly, "Arquivo existente"
    End If
End Sub

Public Sub Deletar(ByVal strRoot As String)
    Dim part As CustomXMLPart
    Dim encontrado As Boolean
    
    For Each part In ThisWorkbook.CustomXMLParts
        If InStr(1, part.xml, strRoot) > 0 Then
            part.Delete
            encontrado = True
        End If
    Next part
    
    If Not encontrado Then
        MsgBox "Não existe arquivo XML " & strRoot & "!", vbOKOnly, "Arquivo não existente"
    End If
End Sub
