Attribute VB_Name = "Módulo1"
Option Explicit

' Constantes para los tipos de operaciones
Private Const OPERATION_WRITE As String = "WRITE"
Private Const OPERATION_IMPROVE As String = "IMPROVE"
Private Const OPERATION_EXTEND As String = "EXTEND"
Private Const OPERATION_REWRITE As String = "REWRITE"
Private Const OPERATION_TRANSLATE As String = "TRANSLATE"

' Constantes para los prompts predefinidos
Private Const PROMPT_IMPROVE As String = "Por favor, mejora y refina el siguiente texto de manera concisa y efectiva, manteniendo su estructura y estilo original. "
Private Const PROMPT_EXTEND As String = "Extiende y desarrolla el siguiente texto, conservando su estructura y estilo original. Procura que la expansión esté en un solo párrafo, sin listas ni elementos adicionales. "
Private Const PROMPT_REWRITE As String = "El texto que se presenta a continuación no me convence. Reescríbelo por completo, pero manteniendo la estructura y el estilo originales. "
Private Const PROMPT_TRANSLATE As String = "Detecta el idioma del siguiente texto y tradúcelo al {0}, conservando la estructura original y presentando solo la traducción final. "
Private Const PROMPT_INSTRUCTION As String = "solo dame lo que te pido, no me des explicacicones adicionales: "
' Configuración de Ollama
Private Const OLLAMA_URL As String = "http://localhost:11434/api/generate"
Private Const OLLAMA_MODEL As String = "llama3.2:3b"

' Estructura para almacenar la configuración de la operación
Private Type OperationConfig
    operationType As String
    TargetLanguage As String
    CustomPrompt As String
End Type

'------------------------------------------------------------------------------
' Función principal para procesar texto
'------------------------------------------------------------------------------
Public Sub ProcessText(Optional operationType As String = "")
    On Error GoTo ErrorHandler
    
    Dim config As OperationConfig
    
    ' Si no se especifica operationType, mostrar menú de selección
    If operationType = "" Then
        operationType = ShowOperationMenu()
        If operationType = "" Then Exit Sub
    End If
    
    ' Configurar la operación
    config = ConfigureOperation(operationType)
    
    ' Obtener y procesar el texto seleccionado
    Dim selectedText As String
    selectedText = GetSelectedText()
    
    ' Generar y aplicar el resultado
    Dim processedText As String
    processedText = ProcessSelectedText(selectedText, config)
    
    ' Aplicar el resultado formateado
    ApplyFormattedText processedText
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error en ProcessText: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Funciones de configuración y menús
'------------------------------------------------------------------------------
Private Function ShowOperationMenu() As String
    Dim choice As Variant
    choice = InputBox( _
        prompt:="Seleccione la operación a realizar:" & vbNewLine & _
                "1. Mejorar texto" & vbNewLine & _
                "2. Alargar texto" & vbNewLine & _
                "3. Reescribir texto" & vbNewLine & _
                "4. Traducir texto" & vbNewLine & _
                "5. ¿Qué quieres que escriba la IA?", _
        Title:="Asistente IA", _
        Default:=5)
    
    Select Case choice
        Case 1: ShowOperationMenu = OPERATION_IMPROVE
        Case 2: ShowOperationMenu = OPERATION_EXTEND
        Case 3: ShowOperationMenu = OPERATION_REWRITE
        Case 4: ShowOperationMenu = OPERATION_TRANSLATE
        Case 5: ShowOperationMenu = OPERATION_WRITE
        Case Else: ShowOperationMenu = ""
    End Select
End Function

Private Function ConfigureOperation(operationType As String) As OperationConfig
    Dim config As OperationConfig
    config.operationType = operationType
    
    Select Case operationType
        Case OPERATION_TRANSLATE
            config.TargetLanguage = InputBox("¿A qué idioma desea traducir?", "Asistente IA", "Inglés")
            If config.TargetLanguage = "" Then Exit Function
            config.CustomPrompt = Replace(PROMPT_TRANSLATE, "{0}", config.TargetLanguage)
            
        Case OPERATION_IMPROVE
            config.CustomPrompt = PROMPT_IMPROVE
            
        Case OPERATION_EXTEND
            config.CustomPrompt = PROMPT_EXTEND
            
        Case OPERATION_REWRITE
            config.CustomPrompt = PROMPT_REWRITE
        
        Case OPERATION_WRITE
            config.CustomPrompt = InputBox("¿Qué quieres que escriba la IA?", "Asistente IA")
    End Select
    
    ConfigureOperation = config
End Function

'------------------------------------------------------------------------------
' Funciones de procesamiento de texto
'------------------------------------------------------------------------------
Private Function GetSelectedText() As String
    GetSelectedText = Selection.text
    
   
    If GetSelectedText = "" Then
        Err.Raise vbObjectError + 1, "GetSelectedText", "No hay texto seleccionado"
    End If
End Function

Private Function ProcessSelectedText(text As String, config As OperationConfig) As String
    ' Limpiar el texto para JSON
    Dim cleanText As String
    cleanText = CleanTextForJSON(text)
    
    ' Generar el prompt final
    Dim finalPrompt As String
    finalPrompt = config.CustomPrompt & PROMPT_INSTRUCTION & cleanText
    
    ' Generar respuesta con Ollama
    ProcessSelectedText = GenerateWithOllama(finalPrompt)
    ProcessSelectedText = Replace(ProcessSelectedText, "\n", vbNewLine)
End Function

'------------------------------------------------------------------------------
' Funciones de formato y presentación
'------------------------------------------------------------------------------
Private Sub ApplyFormattedText(text As String)
    'MsgBox (text)
    Selection.TypeText text
    
    
    
    ' Opcional: Aplicar formato adicional
    'With Selection.Range
    '    .Style = "Normal"
    '    .Font.Name = "Calibri"
    '    .Font.Size = 11
    '    .ParagraphFormat.LineSpacing = LinesToPoints(1.15)
    '    .ParagraphFormat.SpaceAfter = 6
    'End With
End Sub

'------------------------------------------------------------------------------
' Funciones de utilidad
'------------------------------------------------------------------------------
Private Function CleanTextForJSON(text As String) As String
    Dim cleanText As String
    cleanText = text
    
    ' Reemplazar caracteres especiales
    cleanText = Replace(cleanText, vbCr, " ")
    cleanText = Replace(cleanText, vbLf, " ")
    cleanText = Replace(cleanText, vbCrLf, " ")
    cleanText = Replace(cleanText, vbNewLine, " ")
    cleanText = Replace(cleanText, vbTab, " ")
    cleanText = Replace(cleanText, """", "\""")
    
    ' Eliminar caracteres de control
    Dim result As String
    Dim i As Integer
    For i = 1 To Len(cleanText)
        If Asc(Mid(cleanText, i, 1)) > 31 Then
            result = result & Mid(cleanText, i, 1)
        End If
    Next i
    
    ' Eliminar espacios múltiples
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    CleanTextForJSON = Trim(result)
End Function

Private Function GenerateWithOllama(prompt As String) As String
    On Error GoTo ErrorHandler
    
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    
    ' Preparar la solicitud
    Dim requestData As String
    requestData = "{""model"":""" & OLLAMA_MODEL & """,""prompt"":""" & prompt & """,""stream"":false}"
    
    With xhr
        .Open "POST", OLLAMA_URL, False
        .setRequestHeader "Content-Type", "application/json"
        .send requestData
        
        Do While .readyState <> 4
            DoEvents
        Loop
        
        If .Status = 200 Then
            Dim jsonResponse As String
            jsonResponse = .responseText
            
            ' Extraer el texto de la respuesta
            GenerateWithOllama = Mid(jsonResponse, InStr(jsonResponse, """response"":""") + 12, _
                               InStr(InStr(jsonResponse, """response"":""") + 12, jsonResponse, """") - (InStr(jsonResponse, """response"":""") + 12))
        Else
            Err.Raise vbObjectError + 2, "GenerateWithOllama", "Error en la llamada a la IA: " & .Status & " - " & .responseText
        End If
    End With
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "GenerateWithOllama", "Error al generar respuesta: " & Err.Description
End Function

'------------------------------------------------------------------------------
' Subrutinas de acceso público para cada operación específica
'------------------------------------------------------------------------------
Public Sub MejorarTexto()
    ProcessText OPERATION_IMPROVE
End Sub

Public Sub AlargarTexto()
    ProcessText OPERATION_EXTEND
End Sub

Public Sub ReescribirTexto()
    ProcessText OPERATION_REWRITE
End Sub

Public Sub TraducirTexto()
    ProcessText OPERATION_TRANSLATE
End Sub

Public Sub AsistenteIA()
    ProcessText OPERATION_WRITE
End Sub

