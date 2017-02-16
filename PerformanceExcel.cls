VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerformanceExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Ganhos reais de performance
Private Sub Class_Initialize()
    
    'Desabilita o botão Esc
    'Application.EnableCancelKey = xlDisabled
    
    'Valida se nenhum workbook estiver carregado
    If Workbooks.Count = 0 Then
        Exit Sub
    End If
    
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    
    'Desprotege a planilha
    'ActiveSheet.Unprotect "senha"
    
End Sub

'Volta as condições iniciais do Excel
Private Sub Class_Terminate()
    
    'Habilita o botão Esc
    'Application.EnableCancelKey = xlDisabled
    
    'Valida se nenhum workbook estiver carregado
    If Workbooks.Count = 0 Then
        Exit Sub
    End If

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
    'Protege a planilha novamente
'    ActiveSheet.Protect "senha", _
'                DrawingObjects:=True, _
'                Contents:=True, _
'                Scenarios:=True

End Sub


