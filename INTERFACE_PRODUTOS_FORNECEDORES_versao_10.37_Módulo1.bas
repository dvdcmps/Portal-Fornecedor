Attribute VB_Name = "Módulo1"
Option Explicit
        '.Color = 15773696 AZUL
        '.Color = 65535 AMARELO
        '.Color = 5296274 VERDE
        '.Color = 255 VERMELHO
        '&H00FF8080& BOTAO AZUL
        '&H0000FFFF& AMARELO
        '&H0000C000& VERDE
        '&H000000FF& VERMELHO
        Public USUARIO As String
        Public Const Key = "Manutencao"
        
        
Public Sub ESCONDE_M0STRA_COMPOSICAO(PARAM As Boolean)
    '
    Application.ScreenUpdating = False
    '
    Sheets("Composicao").Select
    Columns("U:U").Select
    Selection.EntireColumn.Hidden = PARAM
    Columns("X:X").Select
    Selection.EntireColumn.Hidden = PARAM
    Columns("AA:AA").Select
    Selection.EntireColumn.Hidden = PARAM
    Columns("AD:AD").Select
    Selection.EntireColumn.Hidden = PARAM
    Columns("AG:AG").Select
    Selection.EntireColumn.Hidden = PARAM
    '
End Sub

Public Sub ESCONDE_M0STRA_QUALIDADE(PARAM As Boolean)
    'VALIDACAO CODIGO DE CUIDADOS
    '
    Application.ScreenUpdating = False
    '
    Columns("I:I").Select 'COD.ORG.CERTIFICADOR
    Selection.EntireColumn.Hidden = PARAM
    Columns("K:K").Select 'COD.APLITUDE CERTIF.
    Selection.EntireColumn.Hidden = PARAM
    Columns("O:O").Select 'COD.LOCAL SELO
    Selection.EntireColumn.Hidden = PARAM
    Columns("U:U").Select 'COD.ORIGEM MADEIRA
    Selection.EntireColumn.Hidden = PARAM
    Columns("Z:Z").Select 'COD.ESPECIE 1
    Selection.EntireColumn.Hidden = PARAM
    Columns("AC:AC").Select 'COD.ESPECIE 2
    Selection.EntireColumn.Hidden = PARAM
    Columns("AF:AF").Select 'COD.ESPECIE 3
    Selection.EntireColumn.Hidden = PARAM
    '
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_COMPOSICAO
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_COMPOSICAO
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_COMPOSICAO()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim PERCENT_CEL As Double
Dim COD_MATERIAL As String
Dim PERCENT_TOTAL As Double
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
'
Dim PERCENTE_TOTAL As Integer
'
Dim i As Integer
'
Dim ARR(5) As String
Dim AR As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA Composicao
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
Application.Goto Reference:="VALIDA_SO_ESTA_COMPOS"

If ActiveCell.FormulaR1C1 <> "TRUE" Then
   Exit Sub
End If
' BUSCA USUARIO
If ThisWorkbook.Sheets("PARAMETROS").Visible = False Then
   ThisWorkbook.Sheets("PARAMETROS").Visible = True
End If
'
ThisWorkbook.Sheets("PARAMETROS").Select
Application.Goto Reference:="USUARIO"
USUARIO = ActiveCell.Value
'
'BUSCA CREDITOS
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value
'
Sheets("ERROS de Composicao").Select
Cells.Select
Selection.Delete Shift:=xlUp
'
Range("A1").Select
ActiveCell.FormulaR1C1 = "MENSAGEM"
Range("B1").Select
ActiveCell.FormulaR1C1 = "PLANILHA"
Range("C1").Select
ActiveCell.FormulaR1C1 = "QUADRO"
Range("D1").Select
ActiveCell.FormulaR1C1 = "INTERVALO"
Range("A2").Select
LINHAS_ERRO = 2 'SEGUNDA LINHA

Sheets("Composicao").Select

Application.Goto Reference:="INICIO_COMPOSICAO"
'
COL_CELL = ActiveCell.Column
LIN_CELL = ActiveCell.Row
CORRE_COL = ActiveCell.Column
'
'If USUARIO = "1" Then
'   ESCONDE_M0STRA_COMPOSICAO (True) ' HIDDEN =
'   CORRE_COL = CORRE_COL + 1
'Else
'   ESCONDE_M0STRA_COMPOSICAO (False) ' HIDDEN =
'End If
'
Sheets("Composicao").COMPOSICAO.BackColor = &HC000& 'VERDE
Sheets("Composicao").CUIDADOS.BackColor = &HC000& 'VERDE
Sheets("Inicio").COMPOSICAO.BackColor = &HC000& 'VERDE
Sheets("Composicao").COMPOSICAO.ForeColor = &H8000000E  'BRANCO
Sheets("Composicao").CUIDADOS.ForeColor = &H8000000E  'BRANCO
Sheets("Inicio").COMPOSICAO.ForeColor = &H8000000E  'BRANCO
ERROS = False
'
For i = 1 To QTDE_LINHAS_LIBERADAS
    ''''''''''''''''''''''''''''''''''''''''''''''
    'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
    If Cells(LIN_CELL, 2).Value = 0 Then
       Exit For
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''
    '
    If IsError(Cells(LIN_CELL, CORRE_COL).Value) Then
       CELL_ADRESS = Cells(LIN_CELL, CORRE_COL + 1).Address
       ERROS = True
       Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrado MATERIAL 1. O texto não existe em nossas tabelas."
       Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
       Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
       Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
       Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
       Sheets("Composicao").COMPOSICAO.BackColor = &HFF& 'VERMELHO
       Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
       LINHAS_ERRO = LINHAS_ERRO + 1
       '
    Else
      If IsError(Cells(LIN_CELL, CORRE_COL + 3).Value) Then
         CELL_ADRESS = Cells(LIN_CELL, CORRE_COL + 4).Address
         ERROS = True
         Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrado MATERIAL 2. O texto não existe em nossas tabelas."
         Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
         Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
         Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
         Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
         Sheets("Composicao").COMPOSICAO.BackColor = &HFF& 'VERMELHO
         Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
         LINHAS_ERRO = LINHAS_ERRO + 1
         '
      Else
         If IsError(Cells(LIN_CELL, CORRE_COL + 6).Value) Then
            CELL_ADRESS = Cells(LIN_CELL, CORRE_COL + 7).Address
            ERROS = True
            Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrado MATERIAL 3. O texto não existe em nossas tabelas."
            Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
            Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
            Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
            Sheets("Composicao").COMPOSICAO.BackColor = &HFF& 'VERMELHO
            Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
            LINHAS_ERRO = LINHAS_ERRO + 1
            '
         Else
            If IsError(Cells(LIN_CELL, CORRE_COL + 9).Value) Then
               CELL_ADRESS = Cells(LIN_CELL, CORRE_COL + 10).Address
               ERROS = True
               Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrado MATERIAL 4. O texto não existe em nossas tabelas."
               Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
               Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
               Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
               Sheets("Composicao").COMPOSICAO.BackColor = &HFF& 'VERMELHO
               Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
               LINHAS_ERRO = LINHAS_ERRO + 1
               '
            Else
               If IsError(Cells(LIN_CELL, CORRE_COL + 12).Value) Then
                  CELL_ADRESS = Cells(LIN_CELL, CORRE_COL + 13).Address
                  ERROS = True
                  Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrado MATERIAL 5. O texto não existe em nossas tabelas."
                  Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
                  Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
                  Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                  Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                  Sheets("Composicao").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                  Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                  LINHAS_ERRO = LINHAS_ERRO + 1
                  '
                Else
                    If Cells(LIN_CELL, CORRE_COL).Value = "" Or _
                       Cells(LIN_CELL, CORRE_COL).Value = 0 Then
                       CELL_ADRESS = ActiveCell.Address
                       'MsgBox "Composição Material 1 é de preenchimento obrigatorio. Veja em " & CELL_ADRESS, vbOKOnly, "Composição - CUIDADOS"
                       ERROS = True
                       Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 1).Value = "Composição Material 1 é de preenchimento obrigatorio."
                       Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 2).Value = "Composicao"
                       Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 3).Value = "Composição Material"
                       Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                       Sheets("Composicao").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                       Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                       LINHAS_ERRO = LINHAS_ERRO + 1
                    Else
            
                       CORRE_COL = COL_CELL + 2
                       Cells(LIN_CELL, CORRE_COL).Activate
                       '
                       If Cells(LIN_CELL, CORRE_COL).Value = "" And _
                          Cells(LIN_CELL, CORRE_COL).Value = 0 Then
                          CELL_ADRESS = ActiveCell.Address
                          'MsgBox "Informe os materiais por ordem decrecente de porcentagem na formula. O percentual não pode ser maior que o anterior. Veja em " & CELL_ADRESS, vbOKOnly, "Composição - CUIDADOS"
                          ERROS = True
                          Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 1).Value = "O Percentual na Composição do Material não pode ser vazio ou zero."
                          Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 2).Value = "Composicao"
                          Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 3).Value = "Composição Material"
                          Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                          Sheets("Composicao").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                          Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                          LINHAS_ERRO = LINHAS_ERRO + 1
                          '
                       Else
                         PERCENT_CEL = ActiveCell.Value ' PERCENTUAL 1
                         PERCENT_TOTAL = PERCENT_CEL
                         ARR(1) = Cells(LIN_CELL, CORRE_COL - 1).Value 'MATERIAL 1
                         '
                         CORRE_COL = CORRE_COL + 1
                         Cells(LIN_CELL, CORRE_COL).Activate
                         AR = 2
                         '
                         Do While Cells(LIN_CELL, CORRE_COL).Value <> ""
                            '
                            CORRE_COL = CORRE_COL + 2
                            Cells(LIN_CELL, CORRE_COL).Activate
                            '
                            If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then
                               CELL_ADRESS = ActiveCell.Address
                               'MsgBox "O Percentual na Composição do Material não pode ser vazio ou zero.Veja em " & CELL_ADRESS, vbOKOnly, "Composição - CUIDADOS"
                               ERROS = True
                               Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 1).Value = "O Percentual na Composição do Material não pode ser vazio ou zero."
                               Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 2).Value = "Composicao"
                               Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 3).Value = "Composição Material"
                               Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                               Sheets("Composicao").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                               Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                               LINHAS_ERRO = LINHAS_ERRO + 1
                               '
                               PERCENT_TOTAL = 0
                               Exit Do
                            Else
                               If PERCENT_CEL < ActiveCell.Value Then
                                  CELL_ADRESS = ActiveCell.Address
                                  'MsgBox "Informe os materiais por ordem decrecente de porcentagem na formula. O percentual não pode ser maior que o anterior. Veja em " & CELL_ADRESS, vbOKOnly, "Composição - CUIDADOS"
                                  ERROS = True
                                  Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 1).Value = "Informe os materiais por ordem decrecente de porcentagem na formula. O percentual não pode ser maior que o anterior."
                                  Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 2).Value = "Composicao"
                                  Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 3).Value = "Composição Material"
                                  Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                                  Sheets("Composicao").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                                  Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                                  LINHAS_ERRO = LINHAS_ERRO + 1
                                  '
                                  PERCENT_TOTAL = 0
                                  Exit Do
                               End If
                            End If
                             '
                            PERCENT_CEL = ActiveCell.Value
                            PERCENT_TOTAL = PERCENT_TOTAL + ActiveCell.Value
                            ARR(AR) = Cells(LIN_CELL, CORRE_COL - 1).Value
                            '
                            Select Case AR
                            Case 2
                                 If ARR(2) = ARR(1) Then
                                    CELL_ADRESS = ActiveCell.Address
                                    'MsgBox "Não pode haver componentes iguais na Composição do Material. Veja em linha: " & _
                                    '       ActiveCell.Row & " - Material 2", vbOKOnly, "Composição - Composição"
                                    ERROS = True
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 1).Value = "Não pode haver componentes iguais na composição do material."
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 2).Value = "Composicao"
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 3).Value = "Composição Material"
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                                    '
                                    Sheets("Composicao").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                                    Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                                    LINHAS_ERRO = LINHAS_ERRO + 1
                                 End If
                             '
                            Case 3
                                 If ARR(3) = ARR(2) Or ARR(3) = ARR(1) Then
                                    CELL_ADRESS = ActiveCell.Address
                                    'MsgBox "Não pode haver componentes iguais na Composição do Material. Veja em linha: " & _
                                    '       ActiveCell.Row & " - Material 2", vbOKOnly, "Composição - Composição"
                                    ERROS = True
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 1).Value = "Não pode haver componentes iguais na composição do material."
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 2).Value = "Composicao"
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 3).Value = "Composição Material"
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                                    '
                                    Sheets("Composicao").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                                    Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                                    LINHAS_ERRO = LINHAS_ERRO + 1
                                 End If
                
                            Case 4
                                 If ARR(4) = ARR(3) Or ARR(4) = ARR(2) Or ARR(4) = ARR(1) Then
                                    CELL_ADRESS = ActiveCell.Address
                                    'MsgBox "Não pode haver componentes iguais na Composição do Material. Veja em linha: " & _
                                    '       ActiveCell.Row & " - Material 2", vbOKOnly, "Composição - Composição"
                                    ERROS = True
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 1).Value = "Não pode haver componentes iguais na composição do material."
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 2).Value = "Composicao"
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 3).Value = "Composição Material"
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                                    '
                                    Sheets("Composicao").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                                    Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                                    LINHAS_ERRO = LINHAS_ERRO + 1
                                 End If
                            Case 5
                                 If ARR(5) = ARR(4) Or ARR(5) = ARR(3) Or ARR(5) = ARR(2) Or ARR(5) = ARR(1) Then
                                    CELL_ADRESS = ActiveCell.Address
                                    'MsgBox "Não pode haver componentes iguais na Composição do Material. Veja em linha: " & _
                                    '       ActiveCell.Row & " - Material 2", vbOKOnly, "Composição - Composição"
                                    ERROS = True
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 1).Value = "Não pode haver componentes iguais na composição do material."
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 2).Value = "Composicao"
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 3).Value = "Composição Material"
                                    Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                                    '
                                    Sheets("Composicao").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                                    Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                                    LINHAS_ERRO = LINHAS_ERRO + 1
                                 End If
                           End Select
                           CORRE_COL = CORRE_COL + 1
                           Cells(LIN_CELL, CORRE_COL).Activate
                           AR = AR + 1
                        Loop
                        '
                        If PERCENT_TOTAL > 0 Then
                           If PERCENT_TOTAL < 100 Or PERCENT_TOTAL > 100 Then
                              CELL_ADRESS = Cells(LIN_CELL, CORRE_COL - 1).Address
                              'MsgBox "A soma dos Percentuais na Composição dos Materiais deve ser 100. A soma dos Percentuais na linha " & LIN_CELL & " agora esta em " & PERCENT_TOTAL, vbOKOnly, "Composição - CUIDADOS"
                              ERROS = True
                              Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 1).Value = "A soma dos Percentuais na Composição dos Materiais deve ser 100. A soma dos Percentuais na linha " & LIN_CELL & " , agora esta em " & PERCENT_TOTAL
                              Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 2).Value = "Composicao"
                              Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 3).Value = "Composição Material"
                              Sheets("ERROS de Composicao").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                              Sheets("Composicao").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                              Sheets("Inicio").COMPOSICAO.BackColor = &HFF& 'VERMELHO
                              LINHAS_ERRO = LINHAS_ERRO + 1
                           End If
                        End If
                      End If
                   End If
                End If
            End If
          End If
       End If
    End If
    '
    LIN_CELL = LIN_CELL + 1
    CORRE_COL = COL_CELL
    '
    Cells(LIN_CELL, CORRE_COL).Activate
    PERCENT_TOTAL = 0
    PERCENT_CEL = 0
Next i

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_QUALIDADE()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_QUALIDADE()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim ULT_LIN_ERRO As Integer
'
Dim LINHAS_ERRO As Integer
'
Dim i As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA DADOS GERAIS 2
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
Sheets("ERROS em Dados Qualidade").Select
'
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
ActiveCell.Activate
QTDE_LINHAS_LIBERADAS = ActiveCell.Value
'
Sheets("ERROS em Dados Qualidade").Select
Cells.Select
Selection.Delete Shift:=xlUp
'
Range("A1").Select
ActiveCell.FormulaR1C1 = "MENSAGEM"
Range("B1").Select
ActiveCell.FormulaR1C1 = "PLANILHA"
Range("C1").Select
ActiveCell.FormulaR1C1 = "QUADRO"
Range("D1").Select
ActiveCell.FormulaR1C1 = "INTERVALO"
Range("A2").Select
LINHAS_ERRO = 2 'SEGUNDA LINHA
'
Application.Goto Reference:="VALIDA_SO_ESTA_QUALI"
If ActiveCell.FormulaR1C1 = "TRUE" Then
    Sheets("Dados Qualidade").Select
    Application.Goto Reference:="INICIO_QUALIDADE"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Dados Qualidade").QUALIDADE.BackColor = &HC000& 'VERDE
    Sheets("Inicio").QUALIDADE.BackColor = &HC000& 'VERDE
    Sheets("Dados Qualidade").QUALIDADE.ForeColor = &H8000000E  'BRANCO
    Sheets("Inicio").QUALIDADE.ForeColor = &H8000000E  'BRANCO
    ERROS = False
    '
    For i = 1 To QTDE_LINHAS_LIBERADAS
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
        If ActiveCell.Value = "SIM" Or ActiveCell.Value = "NÃO" Then 'POSSUI CERTIFADO QUALIDADE?
            '
            If ActiveCell.Value = "SIM" Then 'POSSUI CERTIFADO?
               '
'               If USUARIO = "1" Then
'                  CORRE_COL = CORRE_COL + 1  'ORGAÃO CERTIFICADOR
'               Else
                  CORRE_COL = CORRE_COL + 2  'ORGAÃO CERTIFICADOR
'               End If
               '
               Cells(LIN_CELL, CORRE_COL).Activate
               '
               If ActiveCell.Value = "" Then 'ORGAÃO CERTIFICADOR
                  '
                  CELL_ADRESS = ActiveCell.Address
                  'MsgBox "O campo ORGAÃO CERTIFICADOR deve ser preenchido, por que o campo POSSUI CERTIFICADO?(QUALIDADE) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Qualidade - QUALIDADE"
                  ERROS = True
                  Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 1).Value = "O campo ORGÃO CERTIFICADOR deve ser preenchido por que o campo POSSUI CERTIFICADO?(QUALIDADE) = 'SIM'."
                  Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 2).Value = "Dados Qualidade"
                  Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 3).Value = "Certificado de Qualidade"
                  Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                  Sheets("Dados Qualidade").QUALIDADE.BackColor = &HFF& 'VERMELHO
                  Sheets("Inicio").QUALIDADE.BackColor = &HFF& 'VERMELHO
                  LINHAS_ERRO = LINHAS_ERRO + 1
                  '
               End If
               '
'               If USUARIO = "1" Then
'                  CORRE_COL = CORRE_COL + 1  'AMPLITUDE CERTIFICACAO
'               Else
                  CORRE_COL = CORRE_COL + 2  'AMPLITUDE CERTIFICACAO
'               End If
               '
               Cells(LIN_CELL, CORRE_COL).Activate
               '
               If ActiveCell.Value = "" Then 'AMPLITUDE CERTIFICACAO
                  '
                  CELL_ADRESS = ActiveCell.Address
                  'MsgBox "O campo AMPLITUDE CERTIFICACAO deve ser preenchido, por que o campo POSSUI CERTIFICADO?(QUALIDADE) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Qualidade - QUALIDADE"
                  ERROS = True
                  Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 1).Value = "O campo AMPLITUDE CERTIFICACAO deve ser preenchido por que o campo POSSUI CERTIFICADO?(QUALIDADE) = 'SIM'."
                  Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 2).Value = "Dados Qualidade"
                  Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 3).Value = "Certificado de Qualidade"
                  Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                  Sheets("Dados Qualidade").QUALIDADE.BackColor = &HFF& 'VERMELHO
                  Sheets("Inicio").QUALIDADE.BackColor = &HFF& 'VERMELHO
                  LINHAS_ERRO = LINHAS_ERRO + 1
               End If
               '
               CORRE_COL = CORRE_COL + 1 'REG.ORGAO CERTIFICADOR
               Cells(LIN_CELL, CORRE_COL).Activate
               '
               If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then  'REG.ORGAO CERTIFICADOR
                  CELL_ADRESS = ActiveCell.Address
                  'MsgBox "O campo REG.ORGAO CERTIFICADOR deve ser preenchido, por que o campo POSSUI CERTIFICADO?(QUALIDADE) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Qualidade - QUALIDADE"
                  ERROS = True
                  Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 1).Value = "O campo REG.ORGAO CERTIFICADOR deve ser preenchido por que o campo POSSUI CERTIFICADO?(QUALIDADE) = 'SIM'."
                  Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 2).Value = "Dados Qualidade"
                  Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 3).Value = "Certificado de Qualidade"
                  Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                  Sheets("Dados Qualidade").QUALIDADE.BackColor = &HFF& 'VERMELHO
                  Sheets("Inicio").QUALIDADE.BackColor = &HFF& 'VERMELHO
                  LINHAS_ERRO = LINHAS_ERRO + 1
               End If
               
               CORRE_COL = CORRE_COL + 1 'LOCAL SELO/INSCRIÇÃO
               Cells(LIN_CELL, CORRE_COL).Activate
               '
               If ActiveCell.Value = "SIM" Then 'POSSUI SELO CERTIFICACAO?
                  '
                  CORRE_COL = CORRE_COL + 1 'LOCAL SELO/INSCRIÇÃO
                  Cells(LIN_CELL, CORRE_COL).Activate
                  If ActiveCell.Value = "" Then 'LOCAL SELO/INSCRIÇÃO
                     '
                     CELL_ADRESS = ActiveCell.Address
                     'MsgBox "O campo LOCAL SELO/INSCRIÇÃO deve ser preenchido, por que o campo POSSUI SELO CERTIFICACAO?(QUALIDADE) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Qualidade - QUALIDADE"
                     ERROS = True
                     Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 1).Value = "O campo LOCAL SELO/INSCRIÇÃO deve ser preenchido por que o campo POSSUI SELO CERTIFICACAO?(QUALIDADE) = 'SIM'."
                     Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 2).Value = "Dados Qualidade"
                     Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 3).Value = "Selo Certificação"
                     Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                     Sheets("Dados Qualidade").QUALIDADE.BackColor = &HFF& 'VERMELHO
                     Sheets("Inicio").QUALIDADE.BackColor = &HFF& 'VERMELHO
                     LINHAS_ERRO = LINHAS_ERRO + 1
                  End If
                  '
'                  If USUARIO = "1" Then
'                     CORRE_COL = CORRE_COL + 1  'SEL0 DA CERTIFICACAO
'                  Else
                     CORRE_COL = CORRE_COL + 2  'SEL0 DA CERTIFICACAO
'                  End If
                  '
                  Cells(LIN_CELL, CORRE_COL).Activate
                  '
                  If ActiveCell.Value = "" Then 'SEL0 DA CERTIFICACAO
                     '
                     CELL_ADRESS = ActiveCell.Address
                     'MsgBox "O campo SEL0 DA CERTIFICACAO deve ser preenchido, por que o campo POSSUI SELO CERTIFICACAO?(QUALIDADE) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Qualidade - QUALIDADE"
                     ERROS = True
                     Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 1).Value = "O campo SEL0 DA CERTIFICACAO deve ser preenchido por que o campo POSSUI SELO CERTIFICACAO(QUALIDADE) = 'SIM'."
                     Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 2).Value = "Dados Qualidade"
                     Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 3).Value = "Selo Certificação"
                     Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                     Sheets("Dados Qualidade").QUALIDADE.BackColor = &HFF& 'VERMELHO
                     Sheets("Inicio").QUALIDADE.BackColor = &HFF& 'VERMELHO
                     LINHAS_ERRO = LINHAS_ERRO + 1
                  End If
               End If
            End If
        Else
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "O campo POSSUI CERTIFADO?(QUALIDADE) deve ser preenchido, 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - ANVISA"
           ERROS = True
           Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 1).Value = "O campo POSSUI CERTIFADO?(QUALIDADE) deve ser preenchido, 'SIM' ou 'NÃO'."
           Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
           Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 3).Value = "ANVISA"
           Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Dados Qualidade").QUALIDADE.BackColor = &HFF& 'VERMELHO
           Sheets("Inicio").QUALIDADE.BackColor = &HFF& 'VERMELHO
           LINHAS_ERRO = LINHAS_ERRO + 1
        End If
        '
        LIN_CELL = LIN_CELL + 1
        CORRE_COL = COL_CELL
        Cells(LIN_CELL, CORRE_COL).Activate
       
    Next i
    '
'    Sheets("ERROS em Dados Qualidade").Select
'    Cells.Select
'    Selection.ColumnWidth = 8.29
'    Cells.EntireColumn.AutoFit
'    If ERROS Then
'       MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Dados Qualidade'.", vbOKOnly, "Dados Qualidade"
'       Sheets("ERROS em Embalagens").Select
'       Cells.Select
'       Selection.ColumnWidth = 8.29
'       Cells.EntireColumn.AutoFit
'    Else
'      Sheets("Dados Qualidade").Select
'      MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Qualidade"
'    End If
End If
Sheets("Dados Qualidade").Select
Application.Goto Reference:="VALIDA_SO_ESTA_QUALI"
'
''Application.ScreenUpdating = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_MADEIRA()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim ULT_LIN_ERRO As Integer
'
Dim LINHAS_ERRO As Integer
'
Dim i As Integer
'
Dim SIMNAO As String
Dim resultado
'
Dim PERCENTE_TOTAL As Integer
Dim CONTA_ESPECIE_MADEIRA As Integer
'
'PLANILHA DADOS GERAIS 2
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
Sheets("ERROS em Dados Qualidade").Select
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value
'
Sheets("ERROS em Dados Regulatorios").Select
Cells(1, 1).Select ' CELULA A1
ULT_LIN_ERRO = 1
Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
   ULT_LIN_ERRO = ULT_LIN_ERRO + 1
Loop
'
LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Range("A1").Select
ActiveCell.FormulaR1C1 = "MENSAGEM"
Range("B1").Select
ActiveCell.FormulaR1C1 = "PLANILHA"
Range("C1").Select
ActiveCell.FormulaR1C1 = "QUADRO"
Range("D1").Select
ActiveCell.FormulaR1C1 = "INTERVALO"
Range("A2").Select
LINHAS_ERRO = 2 'SEGUNDA LINHA
'
Application.Goto Reference:="VALIDA_SO_ESTA_QUALI"
'
If ActiveCell.FormulaR1C1 = "TRUE" Then
    Sheets("Dados Qualidade").Select
    Application.Goto Reference:="INICIO_MADEIRA"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Dados Qualidade").MADEIRA.BackColor = &HC000& 'VERDE
    Sheets("Dados Qualidade").MADEIRA.ForeColor = &H8000000E  'BRANCO
    'Sheets("Inicio").QUALIDADE.BackColor = &HC000& 'VERDE
    'Sheets("Inicio").QUALIDADE.ForeColor = &H8000000E  'BRANCO
    
    ERROS = False
    For i = 1 To QTDE_LINHAS_LIBERADAS
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
        If ActiveCell.Value = "SIM" Or ActiveCell.Value = "NÃO" Then 'POSSUI MADEIRA?
           If ActiveCell.Value = "SIM" Then 'POSSUI MADEIRA?
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate 'POSSUI CERTIFICADO?
              If ActiveCell.Value = "SIM" Or ActiveCell.Value = "NÃO" Then 'POSSUI CERTIFADO?
                 If ActiveCell.Value = "SIM" Then 'POSSUI CERTIFADO?
                    '
                    CORRE_COL = CORRE_COL + 1  'origem madeira
                    '
                    Cells(LIN_CELL, CORRE_COL).Activate
                    '
                    If ActiveCell.Value = "" Then 'origem madeira
                       '
                       CELL_ADRESS = ActiveCell.Address
                       'MsgBox "O campo ORIGEM MADEIRA deve ser preenchido, por que o campo POSSUI MADEIRA = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Qualidade - MADEIRA"
                       ERROS = True
                       Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 1).Value = "O campo ORIGEM MADEIRA deve ser preenchido por que o campo POSSUI MADEIRA = 'SIM'."
                       Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 2).Value = "Dados Qualidade"
                       Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 3).Value = "Madeira"
                       Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                       Sheets("Dados Qualidade").MADEIRA.BackColor = &HFF& 'VERMELHO
                       Sheets("Inicio").QUALIDADE.BackColor = &HFF& 'VERMELHO
                       LINHAS_ERRO = LINHAS_ERRO + 1
                       '
                    End If
                    '
                    If ActiveCell.Value = "" Then 'origem madeira
                       '
                       CELL_ADRESS = ActiveCell.Address
                       'MsgBox "O campo ORGAÃO CERTIFICADOR deve ser preenchido, por que o campo POSSUI CERTIFICADO?(MADEIRA) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Qualidade - MADEIRA"
                       ERROS = True
                       Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 1).Value = "O campo ORGÃO CERTIFICADOR deve ser preenchido por que o campo POSSUI CERTIFICADO?(MADEIRA) = 'SIM'."
                       Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 2).Value = "Dados Qualidade"
                       Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 3).Value = "Madeira"
                       Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                       Sheets("Dados Qualidade").MADEIRA.BackColor = &HFF& 'VERMELHO
                       Sheets("Inicio").QUALIDADE.BackColor = &HFF& 'VERMELHO
                       LINHAS_ERRO = LINHAS_ERRO + 1
                       '
                    End If
                    '
                    '
                    CORRE_COL = CORRE_COL + 1  'REG.ORGAÃO CERTIFICADOR
                    '
                    Cells(LIN_CELL, CORRE_COL).Activate
                    '
                    If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then  'REG.ORGAO CERTIFICADOR
                       CELL_ADRESS = ActiveCell.Address
                       'MsgBox "O campo REG.ORGAO CERTIFICADOR deve ser preenchido, por que o campo POSSUI CERTIFICADO?(MADEIRA) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Qualidade - MADEIRA"
                       ERROS = True
                       Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 1).Value = "O campo REG.ORGAO CERTIFICADOR deve ser preenchido por que o campo POSSUI CERTIFICADO?(MADEIRA) = 'SIM'."
                       Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 2).Value = "Dados Qualidade"
                       Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 3).Value = "Madeira"
                       Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                       Sheets("Dados Qualidade").MADEIRA.BackColor = &HFF& 'VERMELHO
                       Sheets("Inicio").QUALIDADE.BackColor = &HFF& 'VERMELHO
                       LINHAS_ERRO = LINHAS_ERRO + 1
                    End If
                    '
                 End If 'POSSUI CERTIFADO = NÃO
              Else
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo POSSUI CERTIFICADO?(MADEIRA) deve ser preenchido, 'SIM' ou 'NÃO', porque POSSUI MADEIRA? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - MADEIRA"
                 ERROS = True
                 Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 1).Value = "O campo POSSUI MADEIRA? deve ser preenchido, 'SIM' ou 'NÃO'. , porque POSSUI MADEIRA? = 'SIM'."
                 Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 2).Value = "Dados Qualidade"
                 Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 3).Value = "Madeira"
                 Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Dados Qualidade").MADEIRA.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").QUALIDADE.BackColor = &HFF& 'VERMELHO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              Range("AB4").Select 'PORCENTAGEM MADEIRA 1
              CORRE_COL = ActiveCell.Column
              Cells(LIN_CELL, CORRE_COL).Activate
              PERCENTE_TOTAL = 0
              CONTA_ESPECIE_MADEIRA = 1
              Do While CONTA_ESPECIE_MADEIRA < 4
                 '
                 PERCENTE_TOTAL = PERCENTE_TOTAL + ActiveCell.Value
                 '
'                 If USUARIO = "1" Then
'                    CORRE_COL = CORRE_COL + 3  'PORCENTAGEM
'                 Else
                    CORRE_COL = CORRE_COL + 3
'                 End If
                 '
                 Cells(LIN_CELL, CORRE_COL).Activate
                 '
                 CONTA_ESPECIE_MADEIRA = CONTA_ESPECIE_MADEIRA + 1
              Loop
              '
              If PERCENTE_TOTAL > 0 Then
                 If PERCENTE_TOTAL < 100 Or PERCENTE_TOTAL > 100 Then
                    CELL_ADRESS = Cells(LIN_CELL, CORRE_COL).Address
                    'MsgBox "A soma dos Percentuais na Composição da Madeira deve ser 100. A soma dos Percentuais na linha " & LIN_CELL & " agora esta em " & PERCENTE_TOTAL, vbOKOnly, "Madeira - COMPOSICAO"
                    ERROS = True
                    Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 1).Value = "A soma dos Percentuais na Composição dos Materiais deve ser 100. A soma dos Percentuais na linha " & LIN_CELL & " , agora esta em " & PERCENTE_TOTAL
                    Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 2).Value = "Madeira"
                    Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 3).Value = "Composição Madeira"
                    Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                    Sheets("Dados Qualidade").MADEIRA.BackColor = &HFF& 'VERMELHO
                    Sheets("Inicio").QUALIDADE.BackColor = &HFF& 'VERMELHO
                    LINHAS_ERRO = LINHAS_ERRO + 1
                 End If
              End If
           End If
        Else
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "O campo POSSUI MADEIRA? deve ser preenchido, 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - MADEIRA"
           ERROS = True
           Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 1).Value = "O campo POSSUI MADEIRA? deve ser preenchido, 'SIM' ou 'NÃO'."
           Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 2).Value = "Dados Qualidade"
           Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 3).Value = "Madeira"
           Sheets("ERROS em Dados Qualidade").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Dados Qualidade").MADEIRA.BackColor = &HFF& 'VERMELHO
           Sheets("Inicio").QUALIDADE.BackColor = &HFF& 'VERMELHO
           LINHAS_ERRO = LINHAS_ERRO + 1
        End If
        '
        LIN_CELL = LIN_CELL + 1
        CORRE_COL = COL_CELL
        Cells(LIN_CELL, CORRE_COL).Activate
       
    Next i
    '
'    Sheets("ERROS em Dados Qualidade").Select
'    Cells.Select
'    Selection.ColumnWidth = 8.29
'    Cells.EntireColumn.AutoFit
'    If ERROS Then
'       MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Dados Qualidade'.", vbOKOnly, "Dados Qualidade - MADEIRA"
'       Sheets("ERROS em Embalagens").Select
'       Cells.Select
'       Selection.ColumnWidth = 8.29
'       Cells.EntireColumn.AutoFit
'    Else
'      Sheets("Dados Qualidade").Select
'      MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Qualidade - MADEIRA"
'    End If
End If
Sheets("Dados Qualidade").Select
Application.Goto Reference:="VALIDA_SO_ESTA_QUALI"
'
''Application.ScreenUpdating = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_ANVISA()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_ANVISA()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
'
Dim i As Integer
'
Dim ULT_LIN_ERRO As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA DADOS GERAIS 3
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
Sheets("PARAMETROS").Select
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BUSCA ULTIMA LINHA DA TABELA DE "ERROS em Dados Regulatorios" POR QUE JÁ
'FOI ABASTECIDA PELOS ERROS DO QUADRO PERIGOSO
'
'
Sheets("ERROS em Dados Regulatorios").Select
Cells(1, 1).Select ' CELULA A1
ULT_LIN_ERRO = 1
Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
   ULT_LIN_ERRO = ULT_LIN_ERRO + 1
Loop
'
LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Application.Goto Reference:="VALIDA_SO_ESTA_REGULA"
If ActiveCell.FormulaR1C1 = "TRUE" Then
'
    Sheets("Dados Regulatorios").Select
    Application.Goto Reference:="INICIO_ANVISA"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Dados Regulatorios").ANVISA.BackColor = &HC000& 'VERDE
    'Sheets("Inicio").REGULATORIOS.BackColor = &HC000& 'VERDE
    Sheets("Dados Regulatorios").ANVISA.ForeColor = &H8000000E  'BRANCO
    ERROS = False
    For i = 1 To QTDE_LINHAS_LIBERADAS
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
        SIMNAO = ActiveCell.Value
        If SIMNAO = "SIM" Or SIMNAO = "NÃO" Then
           CORRE_COL = CORRE_COL + 3
           Cells(LIN_CELL, CORRE_COL).Activate
           '
           If ActiveCell.Value = "" And SIMNAO = "SIM" Then 'REGISTRO/LICENÇA/CADASTRO
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo REGISTRO/LICENÇA/CADASTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(ANVISA) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - ANVISA"
                 ERROS = True
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo REGISTRO/LICENÇA/CADASTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(ANVISA) = 'SIM'."
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "ANVISA"
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Dados Regulatorios").ANVISA.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                 LINHAS_ERRO = LINHAS_ERRO + 1
           '  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           Else
               If ActiveCell.Value <> "" And _
                  ActiveCell.Value <> 0 And SIMNAO = "NÃO" Then 'REGISTRO/LICENÇA/CADASTRO
'                  RESULTADO = MsgBox("A Pergunta Inicial 'PRODUTO CONTROLADO' desta linha, foi colocada como 'NÃO'. Portanto " & _
'                                     "O campo REGISTRO/LICENÇA/CADASTRO não poderia estar preenchido. Posso LIMPA-LO?", vbYesNo, "VALIDACAO ANVISA.")
                  If resultado = vbYes Then
                     ActiveCell.ClearContents
                  Else
                      ERROS = True
                      CELL_ADRESS = ActiveCell.Address
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo REGISTRO/LICENÇA/CADASTRO  NÃO deve ser preenchido por que o campo PRODUTO CONTROLADO?(ANVISA) = 'NÃO'."
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "ANVISA"
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                      Sheets("Dados Regulatorios").ANVISA.BackColor = &HFF& 'VERMELHO
                      Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                      LINHAS_ERRO = LINHAS_ERRO + 1
                  End If
               End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Else
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo PRODUTO CONTROLADO?(ANVISA) deve ser preenchido, 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - IBAMA"
             ERROS = True
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRODUTO CONTROLADO?(IBAMA) deve ser preenchido, 'SIM' ou 'NÃO'."
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "ANVISA"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Dados Regulatorios").IBAMA.BackColor = &HFF&
             Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
             LINHAS_ERRO = LINHAS_ERRO + 1
        End If
        '
        LIN_CELL = LIN_CELL + 1
        CORRE_COL = COL_CELL
        Cells(LIN_CELL, CORRE_COL).Activate
    Next i
End If
'
'Sheets("ERROS em Dados Regulatorios").Select
'Cells.Select
'Selection.ColumnWidth = 8.29
'Cells.EntireColumn.AutoFit
'Sheets("Dados Regulatorios").Select
'If ERROS Then
'   MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Dados Regulatorios'.", vbOKOnly, "Dados Regulatorios - ANVISA"
'   Sheets("ERROS em Dados Regulatorios").Select
'   Cells.Select
'   Selection.ColumnWidth = 8.29
'   Cells.EntireColumn.AutoFit
'Else
'  Sheets("Dados Qualidade").Select
'  MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Regulatorios - ANVISA"
'End If
'
''Application.ScreenUpdating = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_IBAMA()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_IBAMA()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
'
Dim i As Integer
'
Dim ULT_LIN_ERRO As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA DADOS GERAIS 3
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
Sheets("PARAMETROS").Select
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BUSCA ULTIMA LINHA DA TABELA DE "ERROS em Dados Regulatorios" POR QUE JÁ
'FOI ABASTECIDA PELOS ERROS DO QUADRO PERIGOSO
'
'
Sheets("ERROS em Dados Regulatorios").Select
Cells(1, 1).Select ' CELULA A1
ULT_LIN_ERRO = 1
Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
   ULT_LIN_ERRO = ULT_LIN_ERRO + 1
Loop
'
LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Application.Goto Reference:="VALIDA_SO_ESTA_REGULA"
If ActiveCell.FormulaR1C1 = "TRUE" Then
'
    Sheets("Dados Regulatorios").Select
    Application.Goto Reference:="INICIO_IBAMA"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Dados Regulatorios").IBAMA.BackColor = &HC000& 'VERDE
    'Sheets("Inicio").REGULATORIOS.BackColor = &HC000& 'VERDE
    Sheets("Dados Regulatorios").IBAMA.ForeColor = &H8000000E  'BRANCO
    ERROS = False
    For i = 1 To QTDE_LINHAS_LIBERADAS
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
        '
        SIMNAO = ActiveCell.Value
        If SIMNAO = "SIM" Or SIMNAO = "NÃO" Then
           CORRE_COL = CORRE_COL + 5
           Cells(LIN_CELL, CORRE_COL).Activate
           '
           If ActiveCell.Value = "" And SIMNAO = "SIM" Then 'REGISTRO/LICENÇA/CADASTRO
              CELL_ADRESS = ActiveCell.Address
              'MsgBox "O campo REGISTRO/LICENÇA/CADASTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(IBAMA) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - IBAMA"
              ERROS = True
              Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo REGISTRO/LICENÇA/CADASTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(IBAMA) = 'SIM'."
              Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
              Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "IBAMA"
              Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
              Sheets("Dados Regulatorios").IBAMA.BackColor = &HFF& 'VERMELHO
              Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
              LINHAS_ERRO = LINHAS_ERRO + 1
           Else
              If ActiveCell.Value <> "" And _
                 ActiveCell.Value <> 0 And SIMNAO = "NÃO" Then 'REGISTRO/LICENÇA/CADASTRO
'                 RESULTADO = MsgBox("A Pergunta Inicial 'PRODUTO CONTROLADO' desta linha, foi colocada como 'NÃO'. Portanto " & _
'                                    "O campo REGISTRO/LICENÇA/CADASTRO não poderia estar preenchido.", vbOKOnly, "VALIDACAO IBAMA.")
                 If resultado = vbYes Then
                    ActiveCell.ClearContents
                 Else
                     ERROS = True
                     CELL_ADRESS = ActiveCell.Address
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo REGISTRO/LICENÇA/CADASTRO  NÃO deve ser preenchido por que o campo PRODUTO CONTROLADO?(IBAMA) = 'NÃO'."
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "IBAMA"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                     Sheets("Dados Regulatorios").IBAMA.BackColor = &HFF& 'VERMELHO
                     Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                     LINHAS_ERRO = LINHAS_ERRO + 1
                 End If
              End If
           End If
           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           '
           CORRE_COL = CORRE_COL + 1
           Cells(LIN_CELL, CORRE_COL).Activate
           '
           If ActiveCell.Value = "" And SIMNAO = "SIM" Then 'VENCIMENTO REGISTRO
              CELL_ADRESS = ActiveCell.Address
              'MsgBox "O campo VENCIMENTO REGISTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(IBAMA) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - IBAMA"
              ERROS = True
              Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo VENCIMENTO REGISTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(IBAMA) = 'SIM'."
              Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
              Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "IBAMA"
              Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
              Sheets("Dados Regulatorios").IBAMA.BackColor = &HFF& 'VERMELHO
              Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
              LINHAS_ERRO = LINHAS_ERRO + 1
           Else
              If ActiveCell.Value <> "" And _
                 ActiveCell.Value <> 0 And SIMNAO = "NÃO" Then 'REGISTRO/LICENÇA/CADASTRO
                 'RESULTADO = MsgBox("A Pergunta Inicial 'PRODUTO CONTROLADO' desta linha, foi colocada como 'NÃO'. Portanto " & _
                 '                   "O campo VENCIMENTO REGISTRO não poderia estar preenchido.", vbOKOnly, "VALIDACAO IBAMA.")
                 If resultado = vbYes Then
                    ActiveCell.ClearContents
                 Else
                     ERROS = True
                     CELL_ADRESS = ActiveCell.Address
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo VENCIMENTO REGISTRO  NÃO deve ser preenchido por que o campo PRODUTO CONTROLADO?(IBAMA) = 'NÃO'."
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "IBAMA"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                     Sheets("Dados Regulatorios").IBAMA.BackColor = &HFF& 'VERMELHO
                     Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                     LINHAS_ERRO = LINHAS_ERRO + 1
                 End If
              End If
           End If
           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           '
           CORRE_COL = CORRE_COL + 1
           Cells(LIN_CELL, CORRE_COL).Activate
           '
           If ActiveCell.Value = "" And SIMNAO = "SIM" Then 'VENCIMENTO REGISTRO
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo PRNCIPIO ATIVO SUJEITO A CONTROLE deve ser preenchido, escolha em DESCR PRINCIPIO ATIVO por que o campo PRODUTO CONTROLADO?(IBAMA) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "IBAMA"
             ERROS = True
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRNCIPIO ATIVO SUJEITO A CONTROLE deve ser preenchido por que o campo PRODUTO CONTROLADO?(IBAMA) = 'SIM'."
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "IBAMA"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Dados Regulatorios").IBAMA.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
             LINHAS_ERRO = LINHAS_ERRO + 1
           Else
               If ActiveCell.Value <> "" And _
                  ActiveCell.Value <> 0 And SIMNAO = "NÃO" Then 'REGISTRO/LICENÇA/CADASTRO
'                  RESULTADO = MsgBox("A Pergunta Inicial 'PRODUTO CONTROLADO' desta linha " & ActiveCell.Row & _
'                                     ", foi colocada como 'NÃO'. Portanto " & _
'                                     "O campo PRNCIPIO ATIVO SUJEITO A CONTROLE não poderia estar preenchido. Posso LIMPA-LO?", vbYesNo, "VALIDACAO IBAMA.")
                  If resultado = vbYes Then
                     Cells(ActiveCell.Row, ActiveCell.Column + 1).Select
                     ActiveCell.ClearContents
                  Else
                      ERROS = True
                      CELL_ADRESS = ActiveCell.Address
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRNCIPIO ATIVO SUJEITO A CONTROLE  NÃO deve ser preenchido por que o campo PRODUTO CONTROLADO?(IBAMA) = 'NÃO'."
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "IBAMA"
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                      Sheets("Dados Regulatorios").IBAMA.BackColor = &HFF& 'VERMELHO
                      Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                      LINHAS_ERRO = LINHAS_ERRO + 1
                  End If
               End If
           End If
           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Else
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo PRODUTO CONTROLADO?(IBAMA) deve ser preenchido, 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - IBAMA"
             ERROS = True
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRODUTO CONTROLADO?(IBAMA) deve ser preenchido, 'SIM' ou 'NÃO'."
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "IBAMA"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Dados Regulatorios").IBAMA.BackColor = &HFF&
             Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
             LINHAS_ERRO = LINHAS_ERRO + 1
       End If
       '
       LIN_CELL = LIN_CELL + 1
       CORRE_COL = COL_CELL
       Cells(LIN_CELL, CORRE_COL).Activate
    Next i
End If
'
'Sheets("ERROS em Dados Regulatorios").Select
'Cells.Select
'Selection.ColumnWidth = 8.29
'Cells.EntireColumn.AutoFit
'Sheets("Dados Regulatorios").Select
'If ERROS Then
'   MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Dados Regulatorios'.", vbOKOnly, "Dados Regulatorios - IBAMA"
'   Sheets("ERROS em Dados Regulatorios").Select
'   Cells.Select
'   Selection.ColumnWidth = 8.29
'   Cells.EntireColumn.AutoFit
'Else
'  Sheets("Dados Qualidade").Select
'  MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Regulatorios - IBAMA"
'End If
'
''Application.ScreenUpdating = True
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_INMETRO
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_INMETRO()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
'
Dim i As Integer
'
Dim ULT_LIN_ERRO As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA DADOS GERAIS 3
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
Sheets("PARAMETROS").Select
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BUSCA ULTIMA LINHA DA TABELA DE "ERROS em Dados Regulatorios" POR QUE JÁ
'FOI ABASTECIDA PELOS ERROS DO QUADRO PERIGOSO
'
'
Sheets("ERROS em Embalagens").Select
Cells(1, 1).Select ' CELULA A1
ULT_LIN_ERRO = 1
Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
   ULT_LIN_ERRO = ULT_LIN_ERRO + 1
Loop
'
LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Application.Goto Reference:="VALIDA_SO_ESTA_REGULA"
If ActiveCell.FormulaR1C1 = "TRUE" Then
'
    Sheets("Dados Regulatorios").Select
    Application.Goto Reference:="INICIO_INMETRO"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Dados Regulatorios").INMETRO.BackColor = &HC000& 'VERD
    'Sheets("Inicio").REGULATORIOS.BackColor = &HC000& 'VERD
    Sheets("Dados Regulatorios").INMETRO.ForeColor = &H8000000E  'BRANCO
    ERROS = False
    For i = 1 To QTDE_LINHAS_LIBERADAS
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
        SIMNAO = ActiveCell.Value
        If SIMNAO = "SIM" Or SIMNAO = "NÃO" Then
           CORRE_COL = CORRE_COL + 2
           Cells(LIN_CELL, CORRE_COL).Activate
           '
           If ActiveCell.Value = "" And SIMNAO = "SIM" Then 'REGISTRO/LICENÇA/CADASTRO
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo REGISTRO/LICENÇA/CADASTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(INMETRO) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios INMETRO"
             ERROS = True
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo REGISTRO/LICENÇA/CADASTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(INMETRO) = 'SIM'."
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "INMETRO"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Dados Regulatorios").INMETRO.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
             LINHAS_ERRO = LINHAS_ERRO + 1
           Else
              If ActiveCell.Value <> "" And _
                 ActiveCell.Value <> 0 And SIMNAO = "NÃO" Then 'REGISTRO/LICENÇA/CADASTRO
'                 RESULTADO = MsgBox("A Pergunta Inicial 'PRODUTO CONTROLADO' desta linha, foi colocada como 'NÃO'. Portanto " & _
'                                    "O campo REGISTRO/LICENÇA/CADASTRO não poderia estar preenchido. Posso LIMPA-LO?", vbYesNo, "VALIDACAO INMETRO.")
                 If resultado = vbYes Then
                    ActiveCell.ClearContents
                 Else
                     ERROS = True
                     CELL_ADRESS = ActiveCell.Address
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo REGISTRO/LICENÇA/CADASTRO  NÃO deve ser preenchido por que o campo PRODUTO CONTROLADO?(ANVISA) = 'NÃO'."
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "INMETRO"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                     Sheets("Dados Regulatorios").ANVISA.BackColor = &HFF& 'VERMELHO
                     Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                     LINHAS_ERRO = LINHAS_ERRO + 1
                 End If
              End If
           End If
           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           '
           CORRE_COL = CORRE_COL + 1
           Cells(LIN_CELL, CORRE_COL).Activate
           '
           If ActiveCell.Value = "" And SIMNAO = "SIM" Then 'VENCIMENTO REGISTRO
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo VENCIMENTO REGISTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(INMETRO) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios INMETRO"
                 ERROS = True
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo VENCIMENTO REGISTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(INMETRO) = 'SIM'."
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "INMETRO"
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Dados Regulatorios").INMETRO.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                 LINHAS_ERRO = LINHAS_ERRO + 1
           Else
              If ActiveCell.Value <> "" And _
                 ActiveCell.Value <> 0 And SIMNAO = "NÃO" Then 'REGISTRO/LICENÇA/CADASTRO
'                 RESULTADO = MsgBox("A Pergunta Inicial 'PRODUTO CONTROLADO' desta linha, foi colocada como 'NÃO'. Portanto " & _
'                                    "O campo VENCIMENTO REGISTRO não poderia estar preenchido. Posso LIMPA-LO?", vbYesNo, "VALIDACAO INMETRO.")
                 If resultado = vbYes Then
                    ActiveCell.ClearContents
                 Else
                     ERROS = True
                     CELL_ADRESS = ActiveCell.Address
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo VENCIMENTO REGISTRO  NÃO deve ser preenchido por que o campo PRODUTO CONTROLADO?(INMETRO) = 'NÃO'."
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "INMETRO"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                     Sheets("Dados Regulatorios").ANVISA.BackColor = &HFF& 'VERMELHO
                     Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                     LINHAS_ERRO = LINHAS_ERRO + 1
                 End If
              End If
           End If
           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Else
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo PRODUTO CONTROLADO?(INMETRO) deve ser preenchido, 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - INMETRO"
             ERROS = True
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRODUTO CONTROLADO?(INMETRO) deve ser preenchido, 'SIM' ou 'NÃO'."
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "INMETRO"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Dados Regulatorios").INMETRO.BackColor = &HFF&
             Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
             LINHAS_ERRO = LINHAS_ERRO + 1
       End If
       '
       LIN_CELL = LIN_CELL + 1
       CORRE_COL = COL_CELL
       Cells(LIN_CELL, CORRE_COL).Activate
    Next i
End If
'
'Sheets("ERROS em Dados Regulatorios").Select
'Cells.Select
'Selection.ColumnWidth = 8.29
'Cells.EntireColumn.AutoFit
'Sheets("Dados Regulatorios").Select
'If ERROS Then
'   MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Dados Regulatorios'.", vbOKOnly, "Dados Regulatorios - INMETRO"
'   Sheets("ERROS em Dados Regulatorios").Select
'   Cells.Select
'   Selection.ColumnWidth = 8.29
'   Cells.EntireColumn.AutoFit
'Else
'  Sheets("Dados Qualidade").Select
'  MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Regulatorios - INMETRO"
'End If
'
''Application.ScreenUpdating = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_ANATEL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_ANATEL()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
'
Dim i As Integer
'
Dim ULT_LIN_ERRO As Integer
'
Dim SIMNAO As String
Dim resultado
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
Sheets("PARAMETROS").Select
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BUSCA ULTIMA LINHA DA TABELA DE "ERROS em Dados Regulatorios" POR QUE JÁ
'FOI ABASTECIDA PELOS ERROS DO QUADRO PERIGOSO
'
'
Sheets("ERROS em Dados Regulatorios").Select
Cells(1, 1).Select ' CELULA A1
ULT_LIN_ERRO = 1
Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
   ULT_LIN_ERRO = ULT_LIN_ERRO + 1
Loop
'
LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Application.Goto Reference:="VALIDA_SO_ESTA_REGULA"
If ActiveCell.FormulaR1C1 = "TRUE" Then
'
    Sheets("Dados Regulatorios").Select
    Application.Goto Reference:="INICIO_ANATEL"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Dados Regulatorios").ANATEL.BackColor = &HC000& 'VERDE
    'Sheets("Inicio").REGULATORIOS.BackColor = &HC000& 'VERDE
    Sheets("Dados Regulatorios").ANATEL.ForeColor = &H8000000E  'BRANCO
    For i = 1 To QTDE_LINHAS_LIBERADAS
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
        SIMNAO = ActiveCell.Value
        If SIMNAO = "SIM" Or SIMNAO = "NÃO" Then
           CORRE_COL = CORRE_COL + 1
           Cells(LIN_CELL, CORRE_COL).Activate
           '
           If ActiveCell.Value = "" And SIMNAO = "SIM" Then 'REGISTRO/LICENÇA/CADASTRO
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo REGISTRO/LICENÇA/CADASTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(ANATEL) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - ANATEL"
                 ERROS = True
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo REGISTRO/LICENÇA/CADASTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(ANATEL) = 'SIM'."
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "ANATEL"
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Dados Regulatorios").ANATEL.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                 LINHAS_ERRO = LINHAS_ERRO + 1
           Else
               If ActiveCell.Value <> "" And _
                  ActiveCell.Value <> 0 And SIMNAO = "NÃO" Then 'REGISTRO/LICENÇA/CADASTRO
'                  RESULTADO = MsgBox("A Pergunta Inicial 'PRODUTO CONTROLADO' desta linha, foi colocada como 'NÃO'. Portanto " & _
'                                     "O campo REGISTRO/LICENÇA/CADASTRO não poderia estar preenchido. Posso LIMPA-LO?", vbYesNo, "VALIDACAO ANATEL.")
                  If resultado = vbYes Then
                     ActiveCell.ClearContents
                  Else
                      CELL_ADRESS = ActiveCell.Address
                      ERROS = True
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo REGISTRO/LICENÇA/CADASTRO  NÃO deve ser preenchido por que o campo PRODUTO CONTROLADO?(ANATEL) = 'NÃO'."
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "ANATEL"
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                      Sheets("Dados Regulatorios").ANVISA.BackColor = &HFF& 'VERMELHO
                      Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                      LINHAS_ERRO = LINHAS_ERRO + 1
                  End If
               End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            CORRE_COL = CORRE_COL + 1
            Cells(LIN_CELL, CORRE_COL).Activate
            '
            If ActiveCell.Value = "" And SIMNAO = "SIM" Then 'VENCIMENTO REGISTRO
               CELL_ADRESS = ActiveCell.Address
               'MsgBox "O campo VENCIMENTO REGISTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(ANATEL) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - ANATEL"
               ERROS = True
               Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo VENCIMENTO REGISTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(ANATEL) = 'SIM'."
               Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
               Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "ANATEL"
               Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Dados Regulatorios").ANATEL.BackColor = &HFF& 'VERMELHO
               Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
               LINHAS_ERRO = LINHAS_ERRO + 1
            Else
              If ActiveCell.Value <> "" And _
                 ActiveCell.Value <> 0 And SIMNAO = "NÃO" Then 'REGISTRO/LICENÇA/CADASTRO
'                 RESULTADO = MsgBox("A Pergunta Inicial 'PRODUTO CONTROLADO' desta linha, foi colocada como 'NÃO'. Portanto " & _
'                                    "O campo VENCIMENTO REGISTRO não poderia estar preenchido. Posso LIMPA-LO?", vbYesNo, "VALIDACAO ANATEL.")
                 If resultado = vbYes Then
                    ActiveCell.ClearContents
                 Else
                     CELL_ADRESS = ActiveCell.Address
                     ERROS = True
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo VENCIMENTO REGISTRO  NÃO deve ser preenchido por que o campo PRODUTO CONTROLADO?(ANATEL) = 'NÃO'."
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "ANATEL"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                     Sheets("Dados Regulatorios").ANVISA.BackColor = &HFF& 'VERMELHO
                     Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                     LINHAS_ERRO = LINHAS_ERRO + 1
                 End If
              End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Else
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo PRODUTO CONTROLADO?(ANATEL) deve ser preenchido, 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - ANATEL"
             ERROS = True
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRODUTO CONTROLADO?(ANATEL) deve ser preenchido, 'SIM' ou 'NÃO'."
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "ANATEL"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Dados Regulatorios").ANATEL.BackColor = &HFF&
             Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
             LINHAS_ERRO = LINHAS_ERRO + 1
       End If
       '
       LIN_CELL = LIN_CELL + 1
       CORRE_COL = COL_CELL
       Cells(LIN_CELL, CORRE_COL).Activate
    Next i
End If
'
'If ERROS Then
'   MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Dados Regulatorios'.", vbOKOnly, "Dados Regulatorios - ANATEL"
'   Sheets("ERROS em Dados Regulatorios").Select
'   Cells.Select
'   Selection.ColumnWidth = 8.29
'   Cells.EntireColumn.AutoFit
'Else
'  Sheets("Dados Regulatorios").Select
'  MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Regulatorios - ANATEL"
'End If
'
''Application.ScreenUpdating = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_MAPA
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_MAPA()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
'
Dim i As Integer
'
Dim ULT_LIN_ERRO As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA DADOS GERAIS 3
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
Sheets("PARAMETROS").Select
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BUSCA ULTIMA LINHA DA TABELA DE "ERROS em Dados Regulatorios" POR QUE JÁ
'FOI ABASTECIDA PELOS ERROS DO QUADRO PERIGOSO
'
'
Sheets("ERROS em Dados Regulatorios").Select
Cells(1, 1).Select ' CELULA A1
ULT_LIN_ERRO = 1
Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
   ULT_LIN_ERRO = ULT_LIN_ERRO + 1
Loop
'
LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Application.Goto Reference:="VALIDA_SO_ESTA_REGULA"
If ActiveCell.FormulaR1C1 = "TRUE" Then
'
    Sheets("Dados Regulatorios").Select
    Application.Goto Reference:="INICIO_MAPA"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Dados Regulatorios").MAPA.BackColor = &HC000& 'VERDE
    'Sheets("Inicio").REGULATORIOS.BackColor = &HC000&  'VERDE
    Sheets("Dados Regulatorios").MAPA.ForeColor = &H8000000E  'BRANCO
    ERROS = False
    '
    For i = 1 To QTDE_LINHAS_LIBERADAS
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
        SIMNAO = ActiveCell.Value
        If SIMNAO = "SIM" Or SIMNAO = "NÃO" Then
           '
           CORRE_COL = CORRE_COL + 1
           Cells(LIN_CELL, CORRE_COL).Activate
           '
           If ActiveCell.Value = "" And SIMNAO = "SIM" Then 'REGISTRO/LICENÇA/CADASTRO
              CELL_ADRESS = ActiveCell.Address
              'MsgBox "O campo REGISTRO/LICENÇA/CADASTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(MAPA) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - MAPA"
              ERROS = True
              Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo REGISTRO/LICENÇA/CADASTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(MAPA) = 'SIM'."
              Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
              Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "MAPA"
              Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
              Sheets("Dados Regulatorios").MAPA.BackColor = &HFF& 'VERMELHO
              Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
              LINHAS_ERRO = LINHAS_ERRO + 1
           Else
              If ActiveCell.Value <> "" And _
                 ActiveCell.Value <> 0 And SIMNAO = "NÃO" Then 'REGISTRO/LICENÇA/CADASTRO
'                 RESULTADO = MsgBox("A Pergunta Inicial 'PRODUTO CONTROLADO' desta linha, foi colocada como 'NÃO'. Portanto " & _
'                                    "O campo REGISTRO/LICENÇA/CADASTRO não poderia estar preenchido. Posso LIMPA-LO?", vbYesNo, "VALIDACAO INMETRO.")
                 If resultado = vbYes Then
                    ActiveCell.ClearContents
                 Else
                     ERROS = True
                     CELL_ADRESS = ActiveCell.Address
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo REGISTRO/LICENÇA/CADASTRO NÃO deve ser preenchido por que o campo PRODUTO CONTROLADO?(MAPA) = 'NÃO'."
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "MAPA"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                     Sheets("Dados Regulatorios").ANVISA.BackColor = &HFF& 'VERMELHO
                     Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                     LINHAS_ERRO = LINHAS_ERRO + 1
                 End If
              End If
           End If
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           '
           CORRE_COL = CORRE_COL + 1
           Cells(LIN_CELL, CORRE_COL).Activate
           '
           If ActiveCell.Value = "" And SIMNAO = "SIM" Then  'VENCIMENTO REGISTRO
                      CELL_ADRESS = ActiveCell.Address
                      'MsgBox "O campo VENCIMENTO REGISTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(MAPA) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - MAPA"
                      ERROS = True
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo VENCIMENTO REGISTRO deve ser preenchido por que o campo PRODUTO CONTROLADO?(MAPA) = 'SIM'."
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "MAPA"
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                      Sheets("Dados Regulatorios").MAPA.BackColor = &HFF& 'VERMELHO
                      Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                      LINHAS_ERRO = LINHAS_ERRO + 1
           Else
              If ActiveCell.Value <> "" And _
                 ActiveCell.Value <> 0 And SIMNAO = "NÃO" Then 'REGISTRO/LICENÇA/CADASTRO
'                 RESULTADO = MsgBox("A Pergunta Inicial 'PRODUTO CONTROLADO' desta linha, foi colocada como 'NÃO'. Portanto " & _
'                                    "O campo VENCIMENTO REGISTRO não poderia estar preenchido. Posso LIMPA-LO?", vbYesNo, "VALIDACAO INMETRO.")
                 If resultado = vbYes Then
                    ActiveCell.ClearContents
                 Else
                     ERROS = True
                     CELL_ADRESS = ActiveCell.Address
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo VENCIMENTO REGISTR NÃO deve ser preenchido por que o campo PRODUTO CONTROLADO?(MAPA) = 'NÃO'."
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "MAPA"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                     Sheets("Dados Regulatorios").ANVISA.BackColor = &HFF& 'VERMELHO
                     Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                     LINHAS_ERRO = LINHAS_ERRO + 1
                 End If
              End If
           End If
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Else
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo PRODUTO CONTROLADO?(MAPA) deve ser preenchido, 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - MAPA"
             ERROS = True
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRODUTO CONTROLADO?(MAPA) deve ser preenchido, 'SIM' ou 'NÃO'."
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "MAPA"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Dados Regulatorios").MAPA.BackColor = &HFF&
             Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
             LINHAS_ERRO = LINHAS_ERRO + 1
       End If
       '
       LIN_CELL = LIN_CELL + 1
       CORRE_COL = COL_CELL
       Cells(LIN_CELL, CORRE_COL).Activate
    Next i
End If
'
'If ERROS Then
'   MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Dados Regulatorios'.", vbOKOnly, "Dados Regulatorios - MAPA"
'   Sheets("ERROS em Dados Regulatorios").Select
'   Cells.Select
'   Selection.ColumnWidth = 8.29
'   Cells.EntireColumn.AutoFit
'Else
'  Sheets("Dados Regulatorios").Select
'  MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Regulatorios - MAPA"
'End If
'
''Application.ScreenUpdating = True
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_CIVIL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_CIVIL()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
'
Dim i As Integer
'
Dim ULT_LIN_ERRO As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA DADOS GERAIS 3
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
Sheets("PARAMETROS").Select
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BUSCA ULTIMA LINHA DA TABELA DE "ERROS em Dados Regulatorios" POR QUE JÁ
'FOI ABASTECIDA PELOS ERROS DO QUADRO PERIGOSO
'
'
Sheets("ERROS em Dados Regulatorios").Select
Cells(1, 1).Select ' CELULA A1
ULT_LIN_ERRO = 1
Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
   ULT_LIN_ERRO = ULT_LIN_ERRO + 1
Loop
'
LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Application.Goto Reference:="VALIDA_SO_ESTA_REGULA"
If ActiveCell.FormulaR1C1 = "TRUE" Then
'
    Sheets("Dados Regulatorios").Select
    Application.Goto Reference:="INICIO_POLICIACIVIL"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Dados Regulatorios").CIVIL.BackColor = &HC000& 'VERDE
    'Sheets("Inicio").REGULATORIOS.BackColor = &HC000& 'VERDE
    Sheets("Dados Regulatorios").CIVIL.ForeColor = &H8000000E  'BRANCO
    ERROS = False
    For i = 1 To QTDE_LINHAS_LIBERADAS
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
        SIMNAO = ActiveCell.Value
        If SIMNAO = "SIM" Or SIMNAO = "NÃO" Then
           CORRE_COL = CORRE_COL + 3
           Cells(LIN_CELL, CORRE_COL).Activate
           '
           If ActiveCell.Value = "" And SIMNAO = "SIM" Then 'PRINCIPIO ATIVO SUJEITO A CONTROLE
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo PRINCIPIO ATIVO CONTROLE deve ser preenchido por que o campo PRODUTO CONTROLADO?(POLICIA_CIVIL) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - POLICIA CIVIL"
                 ERROS = True
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRINCIPIO ATIVO CONTROLE deve ser preenchido por que o campo PRODUTO CONTROLADO?(POLICIA_CIVIL) = 'SIM'."
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "POLICIA CIVIL"
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Dados Regulatorios").CIVIL.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                 LINHAS_ERRO = LINHAS_ERRO + 1
           Else
              If ActiveCell.Value <> "" And _
                 ActiveCell.Value <> 0 And SIMNAO = "NÃO" Then 'REGISTRO/LICENÇA/CADASTRO
'                 RESULTADO = MsgBox("A Pergunta Inicial 'PRODUTO CONTROLADO' desta linha, foi colocada como 'NÃO'. Portanto " & _
'                                    "O campo PRINCIPIO ATIVO CONTROLE não poderia estar preenchido. Posso LIMPA-LO?", vbYesNo, "VALIDACAO IBAMA.")
                 If resultado = vbYes Then
                    ActiveCell.ClearContents
                 Else
                     ERROS = True
                     CELL_ADRESS = ActiveCell.Address
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRINCIPIO ATIVO CONTROLE  NÃO deve ser preenchido por que o campo PRODUTO CONTROLADO?(CIVIL) = 'NÃO'."
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "CIVIL"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                     Sheets("Dados Regulatorios").CIVIL.BackColor = &HFF& 'VERMELHO
                     Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                     LINHAS_ERRO = LINHAS_ERRO + 1
                 End If
              End If
           End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Else
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo PRODUTO CONTROLADO?(POLICIA_CIVIL) deve ser preenchido, 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - POLICIA CIVIL"
             ERROS = True
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRODUTO CONTROLADO?(POLICIA_CIVIL) deve ser preenchido, 'SIM' ou 'NÃO'."
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "POLICIA CIVIL"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Dados Regulatorios").CIVIL.BackColor = &HFF&
             Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
             LINHAS_ERRO = LINHAS_ERRO + 1
       End If
       '
       LIN_CELL = LIN_CELL + 1
       CORRE_COL = COL_CELL
       Cells(LIN_CELL, CORRE_COL).Activate
    Next i
End If
'
'If ERROS Then
'   MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Dados Regulatorios'.", vbOKOnly, "Dados Regulatorios - ANATEL"
'   Sheets("ERROS em Dados Regulatorios").Select
'   Cells.Select
'   Selection.ColumnWidth = 8.29
'   Cells.EntireColumn.AutoFit
'Else
'  Sheets("Dados Regulatorios").Select
'  MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Regulatorios - ANATEL"
'End If
'
''Application.ScreenUpdating = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_FEDERAL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_FEDERAL()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
'
Dim i As Integer
'
Dim ULT_LIN_ERRO As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA DADOS GERAIS 3
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
Sheets("PARAMETROS").Select
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BUSCA ULTIMA LINHA DA TABELA DE "ERROS em Dados Regulatorios" POR QUE JÁ
'FOI ABASTECIDA PELOS ERROS DO QUADRO PERIGOSO
'
'
Sheets("ERROS em Dados Regulatorios").Select
Cells(1, 1).Select ' CELULA A1
ULT_LIN_ERRO = 1
Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
   ULT_LIN_ERRO = ULT_LIN_ERRO + 1
Loop
'
LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Application.Goto Reference:="VALIDA_SO_ESTA_REGULA"
If ActiveCell.FormulaR1C1 = "TRUE" Then
'
    Sheets("Dados Regulatorios").Select
    Application.Goto Reference:="INICIO_POLICIAFEDERAL"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Dados Regulatorios").FEDERAL.BackColor = &HC000& 'VERDE
    'Sheets("Inicio").REGULATORIOS.BackColor = &HC000& 'VERDE
    Sheets("Dados Regulatorios").FEDERAL.ForeColor = &H8000000E  'BRANCO
    ERROS = False
    '
    For i = 1 To QTDE_LINHAS_LIBERADAS
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
        SIMNAO = ActiveCell.Value
        If SIMNAO = "SIM" Or SIMNAO = "NÃO" Then
           CORRE_COL = CORRE_COL + 2
           Cells(LIN_CELL, CORRE_COL).Activate
           '
           If ActiveCell.Value = "" And SIMNAO = "SIM" Then 'PRINCIPIO ATIVO SUJEITO A CONTROLE
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo PRINCIPIO ATIVO SUJEITO A CONTROLE deve ser preenchido , escolha em DESCR.PRINCIPIO ATIVO, por que o campo PRODUTO CONTROLADO?(POLICIA FEDERAL) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - POLICIA FEDERAL"
                 ERROS = True
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRINCIPIO ATIVO SUJEITO A CONTROLE deve ser preenchido por que o campo PRODUTO CONTROLADO?(POLICIA FEDERAL) = 'SIM'."
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "P0LICIA FEDERAL"
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Dados Regulatorios").FEDERAL.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                 LINHAS_ERRO = LINHAS_ERRO + 1
           Else
              If ActiveCell.Value <> "" And _
                 ActiveCell.Value <> 0 And SIMNAO = "NÃO" Then 'REGISTRO/LICENÇA/CADASTRO
'                 RESULTADO = MsgBox("A Pergunta Inicial 'PRODUTO CONTROLADO' desta linha, foi colocada como 'NÃO'. Portanto " & _
'                                    "O campo PRINCIPIO ATIVO SUJEITO A CONTROLE não poderia estar preenchido. Posso LIMPA-LO?", vbYesNo, "VALIDACAO IBAMA.")
                 If resultado = vbYes Then
                    ActiveCell.ClearContents
                 Else
                     ERROS = True
                     CELL_ADRESS = ActiveCell.Address
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRINCIPIO ATIVO SUJEITO A CONTROLE  NÃO deve ser preenchido por que o campo PRODUTO CONTROLADO?(FEDERAL) = 'NÃO'."
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "FEDERAL"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                     Sheets("Dados Regulatorios").FEDERAL.BackColor = &HFF& 'VERMELHO
                     Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                     LINHAS_ERRO = LINHAS_ERRO + 1
                 End If
              End If
           End If
        Else
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo PRODUTO CONTROLADO?(POLICIA FEDERAL) deve ser preenchido, 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - POLICIA FEDERAL"
             ERROS = True
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRODUTO CONTROLADO?(POLICIA FEDERAL) deve ser preenchido, 'SIM' ou 'NÃO'."
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "POLICIA FEDERAL"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Dados Regulatorios").FEDERAL.BackColor = &HFF&
             Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
             LINHAS_ERRO = LINHAS_ERRO + 1

       End If
       '
       LIN_CELL = LIN_CELL + 1
       CORRE_COL = COL_CELL
       Cells(LIN_CELL, CORRE_COL).Activate
    Next i
End If
'
'If ERROS Then
'   MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Dados Regulatorios'.", vbOKOnly, "Dados Regulatorios - FEDERAL"
'   Sheets("ERROS em Dados Regulatorios").Select
'   Cells.Select
'   Selection.ColumnWidth = 8.29
'   Cells.EntireColumn.AutoFit
'Else
'  Sheets("Dados Regulatorios").Select
'  MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Regulatorios - FEDERAL"
'End If
'
''Application.ScreenUpdating = True
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_EXERCITO
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_EXERCITO()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
'
Dim i As Integer
'
Dim ULT_LIN_ERRO As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA DADOS GERAIS 3
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
Sheets("PARAMETROS").Select
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BUSCA ULTIMA LINHA DA TABELA DE "ERROS em Dados Regulatorios" POR QUE JÁ
'FOI ABASTECIDA PELOS ERROS DO QUADRO PERIGOSO
'
'
Sheets("ERROS em Dados Regulatorios").Select
Cells(1, 1).Select ' CELULA A1
ULT_LIN_ERRO = 1
Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
   ULT_LIN_ERRO = ULT_LIN_ERRO + 1
Loop
'
LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Application.Goto Reference:="VALIDA_SO_ESTA_REGULA"
If ActiveCell.FormulaR1C1 = "TRUE" Then
'
    Sheets("Dados Regulatorios").Select
    Application.Goto Reference:="INICIO_EXERCITO"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Dados Regulatorios").EXERCITO.BackColor = &HC000& 'VERDE
    'Sheets("Inicio").REGULATORIOS.BackColor = &HC000& 'VERDE
    Sheets("Dados Regulatorios").EXERCITO.ForeColor = &H8000000E  'BRANCO
    ERROS = False
    '
    For i = 1 To QTDE_LINHAS_LIBERADAS
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
        SIMNAO = ActiveCell.Value
        If SIMNAO = "SIM" Or SIMNAO = "NÃO" Then
           CORRE_COL = CORRE_COL + 1
           Cells(LIN_CELL, CORRE_COL).Activate
           '
           If ActiveCell.Value = "" And SIMNAO = "SIM" Then 'PRINCIPIO ATIVO SUJEITO A CONTROLE
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo PRINCIPIO ATIVO SUJEITO A CONTROLE deve ser preenchido por que o campo PRODUTO CONTROLADO?(EXERCITO) = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - EXERCITO"
                 ERROS = True
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRINCIPIO ATIVO SUJEITO A CONTROLE deve ser preenchido por que o campo PRODUTO CONTROLADO?(EXERCITO) = 'SIM'."
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "EXERCITO"
                 Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Dados Regulatorios").EXERCITO.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                 LINHAS_ERRO = LINHAS_ERRO + 1
           Else
              If ActiveCell.Value <> "" And _
                 ActiveCell.Value <> 0 And SIMNAO = "NÃO" Then 'REGISTRO/LICENÇA/CADASTRO
'                 RESULTADO = MsgBox("A Pergunta Inicial 'PRODUTO CONTROLADO' desta linha, foi colocada como 'NÃO'. Portanto " & _
'                                    "O campo PRINCIPIO ATIVO SUJEITO A CONTROLE não poderia estar preenchido. Posso LIMPA-LO?", vbYesNo, "VALIDACAO EXERCITO.")
                 If resultado = vbYes Then
                    ActiveCell.ClearContents
                 Else
                     ERROS = True
                     CELL_ADRESS = ActiveCell.Address
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRINCIPIO ATIVO SUJEITO A CONTROLE  NÃO deve ser preenchido por que o campo PRODUTO CONTROLADO?(EXERCITO) = 'NÃO'."
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "EXERCITO"
                     Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                     Sheets("Dados Regulatorios").EXERCITO.BackColor = &HFF& 'VERMELHO
                     Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                     LINHAS_ERRO = LINHAS_ERRO + 1
                 End If
              End If
           End If
        Else
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo PRODUTO CONTROLADO?(EXERCITO) deve ser preenchido, 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - EXERCITO"
             ERROS = True
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo PRODUTO CONTROLADO? deve ser preenchido, 'SIM' ou 'NÃO'."
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "EXERCITO"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Dados Regulatorios").EXERCITO.BackColor = &HFF&
             Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
             LINHAS_ERRO = LINHAS_ERRO + 1

       End If
       '
       LIN_CELL = LIN_CELL + 1
       CORRE_COL = COL_CELL
       Cells(LIN_CELL, CORRE_COL).Activate
    Next i
End If
'
'If ERROS Then
'   MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Dados Regulatorios'.", vbOKOnly, "Dados Regulatorios - EXERCITO"
'   Sheets("ERROS em Dados Regulatorios").Select
'   Cells.Select
'   Selection.ColumnWidth = 8.29
'   Cells.EntireColumn.AutoFit
'Else
'  Sheets("Dados Regulatorios").Select
'  MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Regulatorios - EXERCITO"
'End If
'
''Application.ScreenUpdating = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_EMBALAGEBS - UNIDADE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_UNIDADE()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim PERCENT_CEL As Double
Dim COD_MATERIAL As String
Dim PERCENT_TOTAL As Double
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
'
Dim i As Integer
'
Dim USUARIO As Integer
'
'PLANILHA Composicao
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
' BUSCA USUARIO
Sheets("PARAMETROS").Select
Application.Goto Reference:="USUARIO"
USUARIO = ActiveCell.Value
'
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value
'
Sheets("ERROS em Embalagens").Select
Cells.Select
Selection.Delete Shift:=xlUp
'
Range("A1").Select
ActiveCell.FormulaR1C1 = "MENSAGEM"
Range("B1").Select
ActiveCell.FormulaR1C1 = "PLANILHA"
Range("C1").Select
ActiveCell.FormulaR1C1 = "QUADRO"
Range("D1").Select
ActiveCell.FormulaR1C1 = "INTERVALO"
Range("A2").Select
LINHAS_ERRO = 2 'SEGUNDA LINHA
'
COL_CELL = ActiveCell.Column
LIN_CELL = ActiveCell.Row
CORRE_COL = ActiveCell.Column
'
Application.Goto Reference:="VALIDA_SO_ESTA_EMBAL"
If ActiveCell.FormulaR1C1 = "TRUE" Then
'
    Sheets("Embalagens").Select
    Application.Goto Reference:="INICIO_UNIDADE"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Embalagens").BT_UNIDADE.BackColor = &HC000& 'VERDE
    Sheets("Embalagens").BT_UNIDADE.ForeColor = &H8000000E  'BRANCO
    Sheets("Inicio").EMBALAGENS.BackColor = &HC000&  'VERDE
    Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
    'Sheets("Embalagens").MONTADO.BackColor = &HC000& 'VERDE
    'Sheets("Embalagens").MONTADO.ForeColor = &H8000000E  'BRANCO
    ERROS = False
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'APENAS PARA CHECAGEM DE ESTOQUE, COMPRA E VENDA NOS CAMPOS TIPO_MOVIMENTACAO.
    Sheets("Embalagens").INICIO.BackColor = &H808080    'cinza
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To QTDE_LINHAS_LIBERADAS
        Conta_fiscais.NUM_REGISTRO = LIN_CELL
        Conta_fiscais.Repaint
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
        If ActiveCell.Value = "UN" Or ActiveCell.Value = "M" Or _
           ActiveCell.Value = "M2" Or ActiveCell.Value = "UND" Then
           '
           CORRE_COL = CORRE_COL + 1 'PULA COD.EAN
           Cells(LIN_CELL, CORRE_COL).Activate
           If ActiveCell.Value = "" Then
              ActiveCell.Value = "INTERNO"
              CELL_ADRESS = ActiveCell.Address
              'MsgBox "CODIGO EAN é de preenchimento obrigatorio. Se não disponível digite 'INTERNO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
              ERROS = True
              Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "CODIGO EAN é de preenchimento obrigatorio. Foi preenchido com a palavra 'INTERNO'."
              Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
              Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
              Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
              Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF& 'VERMELHO
              Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
              Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
              LINHAS_ERRO = LINHAS_ERRO + 1
           Else
              If Not IsNumeric(ActiveCell.Value) Then
                 ActiveCell.Value = "INTERNO"
              End If
           End If
          '
          CORRE_COL = CORRE_COL + 2 'PULA CCATEG.OD.EAN
          Cells(LIN_CELL, CORRE_COL).Activate
          '
          If ActiveCell.Value = "" Then 'TIPO DE MOVIMENACAO COMPRA
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo TIPO DE MOVIMENACAO COMPRA deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
             ERROS = True
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo TIPO DE MOVIMENACAO COMPRA deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'."
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
             LINHAS_ERRO = LINHAS_ERRO + 1
          End If
          '
          CORRE_COL = CORRE_COL + 1
          Cells(LIN_CELL, CORRE_COL).Activate
          '
          If ActiveCell.Value = "" Then 'TIPO DE MOVIMENACAO VENDA
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo TIPO DE MOVIMENACAO VENDA deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
             ERROS = True
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo TIPO DE MOVIMENACAO VENDA deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'."
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
             LINHAS_ERRO = LINHAS_ERRO + 1
          End If
          '
          CORRE_COL = CORRE_COL + 1
          Cells(LIN_CELL, CORRE_COL).Activate
          '
          If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'MEDIDA COMPR./LARG.
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo MEDIDA COMPR./LARG. deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
             ERROS = True
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo  MEDIDA COMPR./LARG. deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'."
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
             LINHAS_ERRO = LINHAS_ERRO + 1
          End If
          '
          CORRE_COL = CORRE_COL + 1
          Cells(LIN_CELL, CORRE_COL).Activate
          '
          If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'COMPRIMENTO
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo COMPRIMENTO deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
             ERROS = True
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo  COMPRIMENTO deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'."
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
             LINHAS_ERRO = LINHAS_ERRO + 1
          End If
          '
          CORRE_COL = CORRE_COL + 1
          Cells(LIN_CELL, CORRE_COL).Activate
          '
          If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'LARGURA
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo LARGURA deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
             ERROS = True
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo LARGURA deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'."
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
             LINHAS_ERRO = LINHAS_ERRO + 1
          End If
          '
          CORRE_COL = CORRE_COL + 1
          Cells(LIN_CELL, CORRE_COL).Activate
          '
          If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'ALTURA
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo ALTURA deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
             ERROS = True
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo ALTURA deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'."
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
             LINHAS_ERRO = LINHAS_ERRO + 1
          End If
          '
        
          CORRE_COL = CORRE_COL + 2
          Cells(LIN_CELL, CORRE_COL).Activate
          '
          If ActiveCell.Value = "" Then 'COND.ESTOCAGEM
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo CONDICAO ESTOCAGEM deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
             ERROS = True
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo CONDICAO ESTOCAGEM deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'."
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
             LINHAS_ERRO = LINHAS_ERRO + 1
          End If
          '
          CORRE_COL = CORRE_COL + 1
          Cells(LIN_CELL, CORRE_COL).Activate
          '
          If ActiveCell.Value = "" Then 'VALIDADE
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo VALIDADE? deve ser preenchido com 'SIM' ou 'NÃO', por que o campo UNDADE DE MEDIDA = 'UND'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
             ERROS = True
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo VALIDADE? deve ser preenchido com 'SIM' ou 'NÃO', por que o campo UNDADE DE MEDIDA = 'UND'."
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
             Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
             Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
             LINHAS_ERRO = LINHAS_ERRO + 1
          Else
            If ActiveCell.Value = "SIM" Then 'QTDE DIAS
               '
               CORRE_COL = CORRE_COL + 1
               Cells(LIN_CELL, CORRE_COL).Activate
               '
               If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'QTDE DIAS
                  CELL_ADRESS = ActiveCell.Address
                  'MsgBox "O campo QTDE DIAS deve ser preenchido por que os campos UNDADE DE MEDIDA = 'UND' e VALIDADE? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
                  ERROS = True
                  Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo QTDE DIAS deve ser preenchido por que os campos UNDADE DE MEDIDA = 'UND' e VALIDADE? = 'SIM'."
                  Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                  Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
                  Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                  Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF& 'VERMELHO
                  Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                  Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                  LINHAS_ERRO = LINHAS_ERRO + 1
               End If
            Else
              CORRE_COL = CORRE_COL + 1 'VAI EM DIRECAO A ALTO RISCO.
            End If
          End If
         '
         CORRE_COL = CORRE_COL + 1
         Cells(LIN_CELL, CORRE_COL).Activate
         '
         If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'MEDIDA PESO
            CELL_ADRESS = ActiveCell.Address
            'MsgBox "O campo MEDIDA PESO deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
            ERROS = True
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo  MEDIDA PESO deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'."
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF&
            Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
            Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
            LINHAS_ERRO = LINHAS_ERRO + 1
         End If
         '
         CORRE_COL = CORRE_COL + 1
         Cells(LIN_CELL, CORRE_COL).Activate
         '
         '
         If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'PESO BRUTO
            CELL_ADRESS = ActiveCell.Address
            'MsgBox "O campo PESO BRUTO deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
            ERROS = True
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO BRUTO deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'."
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF&
            Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
            Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
            LINHAS_ERRO = LINHAS_ERRO + 1
         Else
           If ActiveCell.Value < Cells(ActiveCell.Row, ActiveCell.Column + 1) Then
               CELL_ADRESS = ActiveCell.Address
               'MsgBox "O campo PESO BRUTO deve ser maior que Peso Liquido. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
               ERROS = True
               Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO BRUTO deve ser Maior ou igual a PESO LIQUIDO."
               Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
               Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
               Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF&
               Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
               Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
               LINHAS_ERRO = LINHAS_ERRO + 1
           End If
        End If
          '
         CORRE_COL = CORRE_COL + 1
         Cells(LIN_CELL, CORRE_COL).Activate

         If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'PESO LIQUIDO
            CELL_ADRESS = ActiveCell.Address
            'MsgBox "O campo PESO LIQUIDO deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
            ERROS = True
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO LIQUIDO deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'."
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF&
            Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
            Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
            LINHAS_ERRO = LINHAS_ERRO + 1
          Else
             If ActiveCell.Value > Cells(ActiveCell.Row, ActiveCell.Column - 1) Then
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo PESO BRUTO deve ser maior que Peso Liquido. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO LIQUIDO deve ser Menor ou igual a PESO BRUTO."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF&
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
             End If
         End If
         '
         CORRE_COL = CORRE_COL + 1
         Cells(LIN_CELL, CORRE_COL).Activate
         '
         'CONTADOR
         'CONVERSAO UMB
         '
         If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'CONVERSAO UMB
            CELL_ADRESS = ActiveCell.Address
            'MsgBox "O campo CONTADOR CONVERSAO UMB deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
            ERROS = True
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo CONTADOR CONVERSAO UMB deve ser preenchido por que o campo UNDADE DE MEDIDA = 'UND'."
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF&
            Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
            Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
            LINHAS_ERRO = LINHAS_ERRO + 1
         End If

       Else
         CELL_ADRESS = ActiveCell.Address
         'MsgBox "O campo UNIDADE DE MEDIDA deve ser preenchido com 'UN' (Unidade), 'M' (Metro) ou 'M2'(Metro Quadrado). Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - UNIDADE"
         ERROS = True
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo UNIDADE DE MEDIDA deve ser preenchido com 'UN' (Unidade), 'M' (Metro) ou 'M2'(Metro Quadrado)."
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Unidade"
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
         Sheets("Embalagens").BT_UNIDADE.BackColor = &HFF& 'VERMELHO
         Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
         Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
         LINHAS_ERRO = LINHAS_ERRO + 1
         '
      End If
      '
      LIN_CELL = LIN_CELL + 1
      CORRE_COL = COL_CELL
      Cells(LIN_CELL, CORRE_COL).Activate
    Next i
End If
'
'Sheets("ERROS em Embalagens").Select
'Cells.Select
'Selection.ColumnWidth = 8.29
'Cells.EntireColumn.AutoFit
'Sheets("Embalagens").Select
'If ERROS Then
'   MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Embalagens'.", vbOKOnly, "Dados Embalagens - UNIDADE"
'   Sheets("ERROS em Embalagens").Select
'   Cells.Select
'   Selection.ColumnWidth = 8.29
'   Cells.EntireColumn.AutoFit
'Else
'  Sheets("Dados Regulatorios").Select
'  MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Embalagens - UNIDADE"
'End If
'
''Application.ScreenUpdating = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_EMBALAGEBS - CAIXA INNER
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_CAIXA_INNER()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim PERCENT_CEL As Double
Dim COD_MATERIAL As String
Dim PERCENT_TOTAL As Double
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
Dim ULT_LIN_ERRO As Integer
'
Dim i As Integer
'
Dim PESO_BRT_UNID As Integer
Dim PESO_LIQ_UNID As Integer
Dim Qtde_peso     As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA Composicao
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
Sheets("PARAMETROS").Select
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BUSCA ULTIMA LINHA DA TABELA DE "ERROS em Dados Regulatorios" POR QUE JÁ
'FOI ABASTECIDA PELOS ERROS DO QUADRO PERIGOSO
'
'
Sheets("ERROS em Embalagens").Select
Cells(1, 1).Select ' CELULA A1
ULT_LIN_ERRO = 1
Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
   ULT_LIN_ERRO = ULT_LIN_ERRO + 1
Loop
'
LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
'
COL_CELL = ActiveCell.Column
LIN_CELL = ActiveCell.Row
CORRE_COL = ActiveCell.Column
'
Application.Goto Reference:="VALIDA_SO_ESTA_EMBAL"
If ActiveCell.FormulaR1C1 = "TRUE" Then
'
    Sheets("Embalagens").Select
    Application.Goto Reference:="INICIO_CAIXA_INNER"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Embalagens").CAIXA_INNER.BackColor = &HC000& 'VERDE
    Sheets("Embalagens").CAIXA_INNER.ForeColor = &H8000000E  'BRANCO
    ERROS = False
    For i = 1 To QTDE_LINHAS_LIBERADAS
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
        If ActiveCell.Value = "SIM" Or ActiveCell.Value = "NÃO" Then 'CAIXA_INNER?
           If ActiveCell.Value = "SIM" Then 'CAIXA_INNER?
              '
              CORRE_COL = CORRE_COL + 2
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Then 'TIPO DE MOVIMENACAO COMPRA
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo TIPO DE MOVIMENACAO COMPRA deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA INNER"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo TIPO DE MOVIMENACAO COMPRA deve ser preenchido por que o campo CAIXA INNER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Inner"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              '
              If ActiveCell.Value = "" Then 'TIPO DE MOVIMENACAO VENDA
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo TIPO DE MOVIMENACAO VENDA deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA INNER"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo TIPO DE MOVIMENACAO VENDA deve ser preenchido por que o campo CAIXA INNER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Inner"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'QUANTIDADE
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo QUANTIDADE deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA_INNER"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo QUANTIDADE deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Inner"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              'MEDIDA COMPR./LARG.
              '
              If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'MEDIDA COMPR./LARG.
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo MEDIDA COMPR./LARG. deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA_INNER"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo MEDIDA COMPR./LARG. deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "CAIXA_INNER"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'COMPRIMENTO
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo COMPRIMENTO deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA_INNER"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo COMPRIMENTO deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "CAIXA_INNER"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'LARGURA
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo LARGURA deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA_INNER"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo LARGURA deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "CAIXA_INNER"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'ALTURA
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo ALTURA deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA_INNER"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo ALTURA deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "CAIXA_INNER"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
             CORRE_COL = CORRE_COL + 1
             Cells(LIN_CELL, CORRE_COL).Activate
             '
             'MEDIDA PESO
             '
             
             If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'MEDIDA PESO
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "O campo MEDIDA PESO deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA INNER"
                ERROS = True
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo MEDIDA PESO deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'."
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "CAIXA_INNER"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF&
                Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                LINHAS_ERRO = LINHAS_ERRO + 1
             End If
             '
             CORRE_COL = CORRE_COL + 1
             Cells(LIN_CELL, CORRE_COL).Activate
             '
             If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'PESO BRUTO
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "O campo PESO BRUTO deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA INNER"
                ERROS = True
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO BRUTO deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'."
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "CAIXA_INNER"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF&
                Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                LINHAS_ERRO = LINHAS_ERRO + 1
             Else
               If ActiveCell.Value < Cells(ActiveCell.Row, ActiveCell.Column + 1) Then
                    CELL_ADRESS = ActiveCell.Address
                    ERROS = True
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO BRUTO deve ser Maior ou Igual a PESO LIQUIDO."
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "CAIXA_INNER"
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                    Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF&
                    Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                    Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                    LINHAS_ERRO = LINHAS_ERRO + 1
               End If
             End If
             '
             CORRE_COL = CORRE_COL + 1
             Cells(LIN_CELL, CORRE_COL).Activate
             '
             If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'PESO LIQUIDO
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "O campo PESO LIQUIDO deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA INNER"
                ERROS = True
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO LIQUIDO deve ser preenchido por que o campo CAIXA_INNER? = 'SIM'."
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "CAIXA_INNER"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF&
                Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                LINHAS_ERRO = LINHAS_ERRO + 1
             Else
               If ActiveCell.Value > Cells(ActiveCell.Row, ActiveCell.Column - 1) Then
                    CELL_ADRESS = ActiveCell.Address
                    ERROS = True
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO LIQUIDO deve ser Menor ou Igual a PESO BRUTO."
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "CAIXA_INNER"
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                    Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF&
                    Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                    Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                    LINHAS_ERRO = LINHAS_ERRO + 1
               End If
             End If
          End If
       Else
         CELL_ADRESS = ActiveCell.Address
         'MsgBox "O campo CAIXA_INNER? deve ser preenchido com 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA_INNER"
         ERROS = True
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo CAIXA_INNER? deve ser preenchido com 'SIM' ou 'NÃO'."
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "CAIXA_INNER"
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
         Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF& 'VERMELHO
         Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
         Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
         LINHAS_ERRO = LINHAS_ERRO + 1
         '
      End If
      '
      LIN_CELL = LIN_CELL + 1
      CORRE_COL = COL_CELL
      Cells(LIN_CELL, CORRE_COL).Activate
    Next i
End If
'
'If ERROS Then
'   MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Embalagens'.", vbOKOnly, "Dados Embalagens - CAIXA INNER"
'   Sheets("ERROS em Embalagens").Select
'   Cells.Select
'   Selection.ColumnWidth = 8.29
'   Cells.EntireColumn.AutoFit
'Else
'  Sheets("Dados Regulatorios").Select
'  MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Embalagens - CAIXA INNER"
'End If
'
''Application.ScreenUpdating = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_EMBALAGEBS - CAIXA MASTER
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_CAIXA_MASTER()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim PERCENT_CEL As Double
Dim COD_MATERIAL As String
Dim PERCENT_TOTAL As Double
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
Dim ULT_LIN_ERRO As Integer
'
Dim i As Integer
'
Dim USUARIO As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA Composicao
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
' BUSCA USUARIO
Sheets("PARAMETROS").Select
Application.Goto Reference:="USUARIO"
USUARIO = ActiveCell.Value
'
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BUSCA ULTIMA LINHA DA TABELA DE "ERROS em Dados Regulatorios" POR QUE JÁ
'FOI ABASTECIDA PELOS ERROS DO QUADRO PERIGOSO
'
'
Sheets("ERROS em Embalagens").Select
Cells(1, 1).Select ' CELULA A1
ULT_LIN_ERRO = 1
Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
   ULT_LIN_ERRO = ULT_LIN_ERRO + 1
Loop
'
LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
'
COL_CELL = ActiveCell.Column
LIN_CELL = ActiveCell.Row
CORRE_COL = ActiveCell.Column
'
Application.Goto Reference:="VALIDA_SO_ESTA_EMBAL"
If ActiveCell.FormulaR1C1 = "TRUE" Then
'
    Sheets("Embalagens").Select
    Application.Goto Reference:="INICIO_CAIXA_MASTER"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Embalagens").CAIXA_MASTER.BackColor = &HC000& 'VERDE
    Sheets("Embalagens").CAIXA_MASTER.ForeColor = &H8000000E  'BRANCO
    ERROS = False
    '
    '
    For i = 1 To QTDE_LINHAS_LIBERADAS
        Conta_fiscais.NUM_REGISTRO = LIN_CELL
        Conta_fiscais.Repaint
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
        If ActiveCell.Value = "SIM" Or ActiveCell.Value = "NÃO" Then 'CAIXA_MASTER?
           If ActiveCell.Value = "SIM" Then 'CAIXA_INNER?
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Then 'TIPO DE CAIXA
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo TIPO DE CAIXA deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA MASTER"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo TIPO DE CAIXA deve ser preenchido por que o campo CAIXA MASTER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Then
                 ActiveCell.Value = "INTERNO"
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "CODIGO EAN é de preenchimento obrigatorio. Se não disponível digite 'INTERNO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "CODIGO EAN é de preenchimento obrigatorio. Foi preenchido com a palavra 'INTERNO'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              Else
                 If Not IsNumeric(ActiveCell.Value) Then
                    ActiveCell.Value = "INTERNO"
                 Else
                    If Cells(LIN_CELL, CORRE_COL - 16).Value = ActiveCell.Value Or _
                       Cells(LIN_CELL, CORRE_COL - 34).Value = ActiveCell.Value Or _
                       Cells(LIN_CELL, CORRE_COL + 15).Value = ActiveCell.Value Then
                            CELL_ADRESS = ActiveCell.Address
                            ERROS = True
                            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "Encontrado CODIGOS EAN iguais na Unidade, Caixa Inner ou Palleté."
                            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                            Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF& 'VERMELHO
                            Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                            Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                            LINHAS_ERRO = LINHAS_ERRO + 1
                    End If
                 End If
              End If
              '
              CORRE_COL = CORRE_COL + 2
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Then 'TIPO DE MOVIMENACAO COMPRA
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo TIPO DE MOVIMENACAO COMPRA deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA MASTER"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo TIPO DE MOVIMENACAO COMPRA deve ser preenchido por que o campo CAIXA MASTER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
                            '
              If ActiveCell.Value = "" Then 'TIPO DE MOVIMENACAO VENDA
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo TIPO DE MOVIMENACAO VENDA deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA MASTER"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo TIPO DE MOVIMENACAO VENDA deve ser preenchido por que o campo CAIXA MASTER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'QUANTIDADE
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo QUANTIDADE deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA_MASTER"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo QUANTIDADE deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              ' MEDIDA COMPRIM´./LARG.

              If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'MEDIDA COMPRIM´./LARG.
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo MEDIDA COMPRIM´./LARG. deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA_MASTER"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo MEDIDA COMPRIM´./LARG. deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'COMPRIMENTO
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo COMPRIMENTO deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA_MASTER"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo COMPRIMENTO deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'LARGURA
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo LARGURA deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA_MASTER"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo LARGURA deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'ALTURA
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo ALTURA deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA_MASTER"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo ALTURA deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              
             CORRE_COL = CORRE_COL + 1
             Cells(LIN_CELL, CORRE_COL).Activate
             '
             'MEDIDA PESO
             '
             If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'MEDIDA PESO
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "O campo MEDIDA PESO deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA MASTER"
                ERROS = True
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo MEDIDA PESO deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'."
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF&
                Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                LINHAS_ERRO = LINHAS_ERRO + 1
             End If
             '
             CORRE_COL = CORRE_COL + 1
             Cells(LIN_CELL, CORRE_COL).Activate
             '
             If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'PESO BRUTO
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "O campo PESO BRUTO deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA MASTER"
                ERROS = True
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO BRUTO deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'."
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF&
                Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                LINHAS_ERRO = LINHAS_ERRO + 1
             Else
                If ActiveCell.Value < Cells(ActiveCell.Row, ActiveCell.Column + 1) Then
                   CELL_ADRESS = ActiveCell.Address
                   ERROS = True
                   Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO BRUTO deve ser Maior ou Igual a PESO LIQUIDO"
                   Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                   Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                   Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                   Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF&
                   Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                   Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                   LINHAS_ERRO = LINHAS_ERRO + 1
                End If
             End If
             '
             CORRE_COL = CORRE_COL + 1
             Cells(LIN_CELL, CORRE_COL).Activate
             '
             If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'PESO LIQUIDO
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "O campo PESO LIQUIDO deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA MASTER"
                ERROS = True
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO LIQUIDO deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'."
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF&
                Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                LINHAS_ERRO = LINHAS_ERRO + 1
             Else
                If ActiveCell.Value > Cells(ActiveCell.Row, ActiveCell.Column - 1) Then
                   CELL_ADRESS = ActiveCell.Address
                   ERROS = True
                   Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO LIQUIDO deve ser Menor ou Igual a PESO BRUTO"
                   Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                   Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                   Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                   Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF&
                   Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                   Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                   LINHAS_ERRO = LINHAS_ERRO + 1
                End If
             End If
             CORRE_COL = CORRE_COL + 1
             Cells(LIN_CELL, CORRE_COL).Activate
             '
             'SOMENTE CENTRAL DE COMPRAS
             If USUARIO = 2 And (ActiveCell.Value = "" Or ActiveCell.Value = 0) Then 'PESO LIQUIDO
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "O campo CONTADOR CONVERSAO UMB deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA MASTER"
                ERROS = True
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo CONTADOR CONVERSAO UMB deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'."
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF&
                Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                LINHAS_ERRO = LINHAS_ERRO + 1
             End If
          End If
       Else
         CELL_ADRESS = ActiveCell.Address
         'MsgBox "O campo CAIXA_MASTER? deve ser preenchido com 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - CAIXA_MASTER"
         ERROS = True
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo CAIXA_MASTER? deve ser preenchido com 'SIM' ou 'NÃO'."
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
         Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF& 'VERMELHO
         Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
         Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
         LINHAS_ERRO = LINHAS_ERRO + 1
         '
      End If
      '
      LIN_CELL = LIN_CELL + 1
      CORRE_COL = COL_CELL
      Cells(LIN_CELL, CORRE_COL).Activate
    Next i
End If
'
'If ERROS Then
'   MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Embalagens'.", vbOKOnly, "Dados Embalagens - CAIXA MASTER"
'   Sheets("ERROS em Embalagens").Select
'   Cells.Select
'   Selection.ColumnWidth = 8.29
'   Cells.EntireColumn.AutoFit
'Else
'  Sheets("Dados Regulatorios").Select
'  MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Embalagens - CAIXA MASTER"
'End If
'
''Application.ScreenUpdating = True
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_EMBALAGEBS - PALLET
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_PALLET()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim PERCENT_CEL As Double
Dim COD_MATERIAL As String
Dim PERCENT_TOTAL As Double
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
Dim ULT_LIN_ERRO As Integer
'
Dim i As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA Composicao
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
'BUSCA USUARIO
Sheets("PARAMETROS").Select
Application.Goto Reference:="USUARIO"
USUARIO = ActiveCell.Value
'
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BUSCA ULTIMA LINHA DA TABELA DE "ERROS em Dados Regulatorios" POR QUE JÁ
'FOI ABASTECIDA PELOS ERROS DO QUADRO PERIGOSO
'
'
Sheets("ERROS em Embalagens").Select
Cells(1, 1).Select ' CELULA A1
ULT_LIN_ERRO = 1
Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
   ULT_LIN_ERRO = ULT_LIN_ERRO + 1
Loop
'
LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
'
COL_CELL = ActiveCell.Column
LIN_CELL = ActiveCell.Row
CORRE_COL = ActiveCell.Column
'
Application.Goto Reference:="VALIDA_SO_ESTA_EMBAL"
If ActiveCell.FormulaR1C1 = "TRUE" Then
'
    Sheets("Embalagens").Select
    Application.Goto Reference:="INICIO_PALLET"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Embalagens").PALLET.BackColor = &HC000& 'VERDE
    Sheets("Embalagens").PALLET.ForeColor = &H8000000E  'BRANCO
    ERROS = False
    '
    For i = 1 To QTDE_LINHAS_LIBERADAS
        Conta_fiscais.NUM_REGISTRO = LIN_CELL
        Conta_fiscais.Repaint
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
        If ActiveCell.Value = "SIM" Or ActiveCell.Value = "NÃO" Then 'CAIXA_MASTER?
           If ActiveCell.Value = "SIM" Then 'CAIXA_INNER?
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Then
                 ActiveCell.Value = "INTERNO"
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "CODIGO EAN é de preenchimento obrigatorio. Se não disponível digite 'INTERNO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "CODIGO EAN é de preenchimento obrigatorio. Foi preenchido com a palavra 'INTERNO'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").PALLET.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              Else
                 If Not IsNumeric(ActiveCell.Value) Then
                    ActiveCell.Value = "INTERNO"
                 Else
                    If Cells(LIN_CELL, CORRE_COL - 15).Value = ActiveCell.Value Or _
                       Cells(LIN_CELL, CORRE_COL - 31).Value = ActiveCell.Value Or _
                       Cells(LIN_CELL, CORRE_COL - 49).Value = ActiveCell.Value Then
                            CELL_ADRESS = ActiveCell.Address
                            ERROS = True
                            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "Encontrado CODIGOS EAN iguais na Unidade, Caixa Inner ou Caixa Master."
                            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Caixa Master"
                            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                            Sheets("Embalagens").CAIXA_MASTER.BackColor = &HFF& 'VERMELHO
                            Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                            Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                            LINHAS_ERRO = LINHAS_ERRO + 1
                    End If
                 End If
              End If
              '
              CORRE_COL = CORRE_COL + 2
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Then 'UNIDADE COMPOE PALLET
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo UNIDADE COMPOE PALLET deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo UNIDADE COMPOE PALLET deve ser preenchido por que o campo PALLET? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").PALLET.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Then 'TIPO DE MOVIMENACAO COMPRA
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo TIPO DE MOVIMENACAO COMPRA deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo TIPO DE MOVIMENACAO COMPRA deve ser preenchido por que o campo CAIXA INNER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF& 'VERMELHO
                 'SE ALGUM BOTAO AINDA É BRANCO INICO CORRESPONDENTE BRANCO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              '
              If ActiveCell.Value = "" Then 'TIPO DE MOVIMENACAO VENDA
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo TIPO DE MOVIMENACAO VENDA deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo TIPO DE MOVIMENACAO VENDA deve ser preenchido por que o campo CAIXA INNER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              'MEDIDA COMPRIM./LARG.
              '
              If ActiveCell.Value = "" Then 'MEDIDA COMPRIM./LARG.
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo MEDIDA COMPRIM./LARG. VENDA deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo MEDIDA COMPRIM./LARG. deve ser preenchido por que o campo CAIXA INNER? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").CAIXA_INNER.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'COMPRIMENTO
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo COMPRIMENTO deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo COMPRIMENTO deve ser preenchido por que o campo PALLET? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").PALLET.BackColor = &HFF& 'VERMELHO
                 'SE ALGUM BOTAO AINDA É BRANCO INICO CORRESPONDENTE BRANCO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'LARGURA
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo LARGURA deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo LARGURA deve ser preenchido por que o campo PALLET? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              '
              CORRE_COL = CORRE_COL + 1
              Cells(LIN_CELL, CORRE_COL).Activate
              '
              If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'ALTURA
                 CELL_ADRESS = ActiveCell.Address
                 'MsgBox "O campo ALTURA deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                 ERROS = True
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo ALTURA deve ser preenchido por que o campo PALLET? = 'SIM'."
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                 Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                 Sheets("Embalagens").PALLET.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                 Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                 LINHAS_ERRO = LINHAS_ERRO + 1
              End If
              
             CORRE_COL = CORRE_COL + 1
             Cells(LIN_CELL, CORRE_COL).Activate
             '
             'MEDIDA PESO
             '
             If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'MEDIDA PESO
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "O campo MEDIDA PESO deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                ERROS = True
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo MEDIDA PESO deve ser preenchido por que o campo PALLET? = 'SIM'."
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Embalagens").PALLET.BackColor = &HFF&
                Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                LINHAS_ERRO = LINHAS_ERRO + 1
             End If
             '
             CORRE_COL = CORRE_COL + 1
             Cells(LIN_CELL, CORRE_COL).Activate
             '
             If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'PESO BRUTO
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "O campo PESO BRUTO deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                ERROS = True
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO BRUTO deve ser preenchido por que o campo PALLET? = 'SIM'."
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Embalagens").PALLET.BackColor = &HFF&
                Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                LINHAS_ERRO = LINHAS_ERRO + 1
             Else
                If ActiveCell.Value < Cells(ActiveCell.Row, ActiveCell.Column + 1) Then
                    CELL_ADRESS = ActiveCell.Address
                    'MsgBox "O campo PESO LIQUIDO deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                    ERROS = True
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO BRUTO deve ser Maior ou Igual a PESO LIQUIDO."
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                    Sheets("Embalagens").PALLET.BackColor = &HFF&
                    Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                    Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                    LINHAS_ERRO = LINHAS_ERRO + 1
                End If
             End If
             '
             CORRE_COL = CORRE_COL + 1
             Cells(LIN_CELL, CORRE_COL).Activate
             '
             If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'PESO LIQUIDO
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "O campo PESO LIQUIDO deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                ERROS = True
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO LIQUIDO deve ser preenchido por que o campo PALLET? = 'SIM'."
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Embalagens").PALLET.BackColor = &HFF&
                Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                LINHAS_ERRO = LINHAS_ERRO + 1
             Else
                If ActiveCell.Value > Cells(ActiveCell.Row, ActiveCell.Column - 1) Then
                    CELL_ADRESS = ActiveCell.Address
                    'MsgBox "O campo PESO LIQUIDO deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                    ERROS = True
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO LIQUIDO deve ser Menor ou Igual a PESO BRUTO."
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                    Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                    Sheets("Embalagens").PALLET.BackColor = &HFF&
                    Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                    Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                    LINHAS_ERRO = LINHAS_ERRO + 1
                End If
             End If
             '
             CORRE_COL = CORRE_COL + 1
             Cells(LIN_CELL, CORRE_COL).Activate
             '
             If USUARIO = 2 And (ActiveCell.Value = "" Or ActiveCell.Value = 0) Then 'CAMADA
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "O campo CAMADA deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                ERROS = True
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo CAMADA deve ser preenchido por que o campo PALLET? = 'SIM'."
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Embalagens").PALLET.BackColor = &HFF&
                Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                LINHAS_ERRO = LINHAS_ERRO + 1
             End If
             '
             CORRE_COL = CORRE_COL + 1
             Cells(LIN_CELL, CORRE_COL).Activate
             '
             If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'LASTRO
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "O campo LASTRO deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                ERROS = True
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo LASTRO deve ser preenchido por que o campo PALLET? = 'SIM'."
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Embalagens").PALLET.BackColor = &HFF&
                Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                LINHAS_ERRO = LINHAS_ERRO + 1
             End If
             '
             CORRE_COL = CORRE_COL + 1
             Cells(LIN_CELL, CORRE_COL).Activate
             '
             If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'EMPILHAMENTO MAXIMO
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "O campo EMPILHAMENTO MAXIMO deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                ERROS = True
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo EMPILHAMENTO MAXIMO deve ser preenchido por que o campo PALLET? = 'SIM'."
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Embalagens").PALLET.BackColor = &HFF&
                Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                LINHAS_ERRO = LINHAS_ERRO + 1
             End If
             '
             CORRE_COL = CORRE_COL + 1
             Cells(LIN_CELL, CORRE_COL).Activate
             '
             If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'TIPO PALLET
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "O campo TIPO PALLET deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                ERROS = True
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo TIPO PALLET deve ser preenchido por que o campo PALLET? = 'SIM'."
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Embalagens").PALLET.BackColor = &HFF&
                Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                LINHAS_ERRO = LINHAS_ERRO + 1
             End If
             '
             CORRE_COL = CORRE_COL + 1
             Cells(LIN_CELL, CORRE_COL).Activate
             '
             'SOMENTE CENTRAL DE COMPRAS
             If USUARIO = 2 And (ActiveCell.Value = "" Or ActiveCell.Value = 0) Then 'TIPO PALLET
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "O campo CONTADOR CONVERSAO UMB deve ser preenchido por que o campo PALLET? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
                ERROS = True
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo CONTADOR CONVERSAO UMB deve ser preenchido por que o campo PALLET? = 'SIM'."
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
                Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Embalagens").PALLET.BackColor = &HFF&
                Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
                Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
                LINHAS_ERRO = LINHAS_ERRO + 1
             End If
         End If
       Else
         CELL_ADRESS = ActiveCell.Address
         'MsgBox "O campo PALLET? deve ser preenchido com 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - PALLET"
         ERROS = True
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PALLET? deve ser preenchido com 'SIM' ou 'NÃO'."
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Pallet"
         Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
         Sheets("Embalagens").PALLET.BackColor = &HFF& 'VERMELHO
         Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
         Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
         LINHAS_ERRO = LINHAS_ERRO + 1
         '
      End If
      '
      LIN_CELL = LIN_CELL + 1
      CORRE_COL = COL_CELL
      Cells(LIN_CELL, CORRE_COL).Activate
    Next i
End If
'
'If ERROS Then
'   MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Embalagens'.", vbOKOnly, "Dados Embalagens - PALLET"
'   Sheets("ERROS em Embalagens").Select
'   Cells.Select
'   Selection.ColumnWidth = 8.29
'   Cells.EntireColumn.AutoFit
'Else
'  Sheets("Dados Regulatorios").Select
'  MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Embalagens - PALLET"
'End If
'
''Application.ScreenUpdating = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_EMBALAGEBS - MONTADO
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_MONTADO()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim PERCENT_CEL As Double
Dim COD_MATERIAL As String
Dim PERCENT_TOTAL As Double
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
Dim ULT_LIN_ERRO As Integer
'
Dim i As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA Composicao
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BUSCA ULTIMA LINHA DA TABELA DE "ERROS em Embalagens" POR QUE JÁ
'FOI ABASTECIDA PELOS ERROS DO QUADRO EMBALAGENS
'
'
Sheets("ERROS em Embalagens").Select
Cells(1, 1).Select ' CELULA A1
ULT_LIN_ERRO = 1
Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
   ULT_LIN_ERRO = ULT_LIN_ERRO + 1
Loop
'
LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
'
COL_CELL = ActiveCell.Column
LIN_CELL = ActiveCell.Row
CORRE_COL = ActiveCell.Column
'
Application.Goto Reference:="VALIDA_SO_ESTA_EMBAL"
If ActiveCell.FormulaR1C1 = "TRUE" Then
'
    Sheets("Embalagens").Select
    Application.Goto Reference:="INICIO_MONTADO"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Embalagens").MONTADO.BackColor = &HC000& 'VERDE
    Sheets("Embalagens").MONTADO.ForeColor = &H8000000E  'BRANCO
    ERROS = False
    '
    For i = 1 To QTDE_LINHAS_LIBERADAS
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
        If ActiveCell.Value = "" Then 'MEDIDA COMPR./LARG.
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "O campo MEDIDA COMPRIMENTO LARGURA deve ser preenchido. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - MONTADO"
           ERROS = True
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo MEDIDA COMPRIMENTO LARGURA deve ser preenchido."
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Montado"
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Embalagens").MONTADO.BackColor = &HFF& 'VERMELHO
           Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
           Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
           LINHAS_ERRO = LINHAS_ERRO + 1
        End If
         '
         CORRE_COL = CORRE_COL + 1
         Cells(LIN_CELL, CORRE_COL).Activate
         '
         If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'COMPRIMENTO
            CELL_ADRESS = ActiveCell.Address
            'MsgBox "O campo COMPRIMENTO deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - MONTADO"
            ERROS = True
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo COMPRIMENTO deve ser preenchido."
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Montado"
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Embalagens").MONTADO.BackColor = &HFF& 'VERMELHO
            Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
            Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
            LINHAS_ERRO = LINHAS_ERRO + 1
         End If
         '
         CORRE_COL = CORRE_COL + 1
         Cells(LIN_CELL, CORRE_COL).Activate
         '
         If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'LARGURA
            CELL_ADRESS = ActiveCell.Address
            'MsgBox "O campo LARGURA deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - MONTADO"
            ERROS = True
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo LARGURA deve ser preenchido."
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Montado"
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Embalagens").MONTADO.BackColor = &HFF& 'VERMELHO
            Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
            Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
            LINHAS_ERRO = LINHAS_ERRO + 1
         End If
         '
         CORRE_COL = CORRE_COL + 1
         Cells(LIN_CELL, CORRE_COL).Activate
         '
         If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'ALTURA
            CELL_ADRESS = ActiveCell.Address
            'MsgBox "O campo ALTURA deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - MONTADO"
            ERROS = True
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo ALTURA deve ser preenchido."
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Montado"
            Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Embalagens").MONTADO.BackColor = &HFF& 'VERMELHO
            Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
            Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
            LINHAS_ERRO = LINHAS_ERRO + 1
         End If
         
        CORRE_COL = CORRE_COL + 1
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        'MEDIDA PESO
        '
        If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'MEDIDA PESO
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "O campo MEDIDA PESO deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - MONTADO"
           ERROS = True
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo MEDIDA PESO deve ser preenchido."
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Montado"
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Embalagens").MONTADO.BackColor = &HFF&
           Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
           Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
           LINHAS_ERRO = LINHAS_ERRO + 1
        End If
        '
        CORRE_COL = CORRE_COL + 1
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'PESO BRUTO
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "O campo PESO BRUTO deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - MONTADO"
           ERROS = True
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO BRUTO deve ser preenchido."
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Montado"
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Embalagens").MONTADO.BackColor = &HFF&
           Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
           Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
           LINHAS_ERRO = LINHAS_ERRO + 1
        Else
           If ActiveCell.Value < Cells(ActiveCell.Row, ActiveCell.Column + 1) Then
               CELL_ADRESS = ActiveCell.Address
               'MsgBox "O campo PESO BRUTO deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - MONTADO"
               ERROS = True
               Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO BRUTO deve ser Maior ou Igual a PESO LIQUIDO"
               Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
               Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Montado"
               Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Embalagens").MONTADO.BackColor = &HFF&
               Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
               Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
               LINHAS_ERRO = LINHAS_ERRO + 1
           End If
           '
        End If
            
        '
        CORRE_COL = CORRE_COL + 1
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'PESO LIQUIDO
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "O campo PESO LIQUIDO deve ser preenchido por que o campo CAIXA_MASTER? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Embalagens - MONTADO"
           ERROS = True
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO LIQUIDO deve ser preenchido."
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Montado"
           Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Embalagens").MONTADO.BackColor = &HFF&
           Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
           Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
           LINHAS_ERRO = LINHAS_ERRO + 1
        Else
           If ActiveCell.Value > Cells(ActiveCell.Row, ActiveCell.Column - 1) Then
               CELL_ADRESS = ActiveCell.Address
               'MsgBox "O campo PESO LIQUIDO deve ser Menor" & CELL_ADRESS, vbOKOnly, "Embalagens - MONTADO"
               ERROS = True
               Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "O campo PESO LIQUIDO deve ser Menor ou Igual a PESO BRUTO"
               Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
               Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Montado"
               Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Embalagens").MONTADO.BackColor = &HFF&
               Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
               Sheets("Inicio").EMBALAGENS.ForeColor = &H8000000E  'BRANCO
               LINHAS_ERRO = LINHAS_ERRO + 1
           End If
           '
        End If
      '
      LIN_CELL = LIN_CELL + 1
      CORRE_COL = COL_CELL
      Cells(LIN_CELL, CORRE_COL).Activate
    Next i
End If
'
'If ERROS Then
'   MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Embalagens'.", vbOKOnly, "Dados Embalagens - MONTADO"
'   Sheets("ERROS em Embalagens").Select
'   Cells.Select
'   Selection.ColumnWidth = 8.29
'   Cells.EntireColumn.AutoFit
'Else
'  Sheets("Dados Regulatorios").Select
'  MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Embalagens - MONTADO"
'End If
'
''Application.ScreenUpdating = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_INICIAIS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function VALIDACAO_INICIAIS() As Boolean
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim COD_MATERIAL As String
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
'
Dim i As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA INICIAIS
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
ActiveWorkbook.Unprotect Key
If Sheets("PARAMETROS").Visible = False Then
   Sheets("PARAMETROS").Visible = True
End If
'
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value
'BUSCA USUARIO ULTIMO DOWLOAD
Sheets("PARAMETROS").Select
Application.Goto Reference:="USUARIO"
USUARIO = ActiveCell.Value
'
Sheets("ERROS Cadastrais").Select
Cells.Select
Selection.Delete Shift:=xlUp
'
Range("A1").Select
ActiveCell.FormulaR1C1 = "MENSAGEM"
Range("B1").Select
ActiveCell.FormulaR1C1 = "PLANILHA"
Range("C1").Select
ActiveCell.FormulaR1C1 = "QUADRO"
Range("D1").Select
ActiveCell.FormulaR1C1 = "INTERVALO"
Range("A2").Select
LINHAS_ERRO = 2 'SEGUNDA LINHA
'
Sheets("Dados Cadastrais").Select
Application.Goto Reference:="DADOS_INICIAIS"
'
COL_CELL = ActiveCell.Column
LIN_CELL = ActiveCell.Row
CORRE_COL = ActiveCell.Column
'
Sheets("Inicio").CADASTRAIS.BackColor = &HC000& 'VERDE
Sheets("Inicio").CADASTRAIS.ForeColor = &H8000000E  'BRANCO
'
Sheets("Dados Cadastrais").INICIO.BackColor = &H808080    'CINZA
ERROS = False
'
For i = 1 To QTDE_LINHAS_LIBERADAS
    '
    If ActiveCell.Value = "" Then
       If Cells(LIN_CELL + 1, COL_CELL).Value <> 0 And _
          Cells(LIN_CELL + 1, COL_CELL).Value <> "" Then
          CELL_ADRESS = ActiveCell.Address
          'MsgBox "COD. MATERIAL DO FORNECEDOR é de preenchimento obrigatorio. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
          ERROS = True
          Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "COD. MATERIAL DO FORNECEDOR é de preenchimento obrigatorio."
          Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
          Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
          Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
          Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
          Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
          Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
          VALIDACAO_INICIAIS = True
          LINHAS_ERRO = LINHAS_ERRO + 1
          Exit For
       Else
         Exit For
       End If
    End If
    '
    CORRE_COL = CORRE_COL + 1
    Cells(LIN_CELL, CORRE_COL).Activate
    '
    If ActiveCell.Value = "" Then
       CELL_ADRESS = ActiveCell.Address
       ERROS = True
       Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "MATER.FORNEC. é de preenchimento obrigatorio. Foi preenchido com a palavra 'INTERNO'."
       Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
       Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
       Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
       Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
       Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
       Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
       VALIDACAO_INICIAIS = True
       LINHAS_ERRO = LINHAS_ERRO + 1

    End If
    '
    CORRE_COL = CORRE_COL + 1
    Cells(LIN_CELL, CORRE_COL).Activate
    '
    If ActiveCell.Value = "" Then
       ActiveCell.Value = "INTERNO"
       CELL_ADRESS = ActiveCell.Address
       'MsgBox "CODIGO EAN é de preenchimento obrigatorio. Se não disponível digite 'INTERNO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
       ERROS = True
       Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "CODIGO EAN é de preenchimento obrigatorio. Foi preenchido com a palavra 'INTERNO'."
       Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
       Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
       Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
       Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
       Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
       Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
       VALIDACAO_INICIAIS = True
       LINHAS_ERRO = LINHAS_ERRO + 1
    Else
      If Not IsNumeric(ActiveCell.Value) Then
         ActiveCell.Value = "INTERNO"
      End If
    End If
    '
    If USUARIO = 2 Then
        On Error Resume Next
        CORRE_COL = CORRE_COL + 2
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HC000& 'VERDE
        Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
        '
        If ActiveCell.Value = "" Then
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "DESCR.COMERCIAL  é de preenchimento obrigatorio. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
           ERROS = True
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "DESCR.COMERCIAL é de preenchimento obrigatorio."
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
           '
           Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
           '
           VALIDACAO_INICIAIS = True
           LINHAS_ERRO = LINHAS_ERRO + 1
        End If
        '
        CORRE_COL = CORRE_COL + 1
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If ActiveCell.Value = "" Then
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "DESCR.COMERCIAL RESUMIDA é de preenchimento obrigatorio. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
           ERROS = True
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "DESCR.COMERC.RESUM é de preenchimento obrigatorio."
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHODA
           Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
           '
           Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
           '
           VALIDACAO_INICIAIS = True
           LINHAS_ERRO = LINHAS_ERRO + 1
        End If
        '
        CORRE_COL = CORRE_COL + 3
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If IsError(ActiveCell.Value) Then
           CELL_ADRESS = ActiveCell.Address
           ERROS = True
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrada DESCR.GRUPO MERCADORIAS. O texto não existe em nossas tabelas."
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
           '
           Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
           '
           VALIDACAO_INICIAIS = True
           LINHAS_ERRO = LINHAS_ERRO + 1
        Else
            If ActiveCell.Value = "" Then
               CELL_ADRESS = ActiveCell.Address
               'MsgBox "GRUPO MERCADORIAS  é de preenchimento obrigatorio. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
               ERROS = True
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "GRUPO MERCADORIAS é de preenchimento obrigatorio."
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
               '
               Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
               '
               VALIDACAO_INICIAIS = True
               LINHAS_ERRO = LINHAS_ERRO + 1
              End If
        End If
        '
        CORRE_COL = CORRE_COL + 2
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If IsError(ActiveCell.Value) Then
           CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column + 1).Address
           ERROS = True
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrada NÓ DE HIERARQUIA. O texto não existe em nossas tabelas."
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
           '
           Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
           '
           VALIDACAO_INICIAIS = True
           LINHAS_ERRO = LINHAS_ERRO + 1
        Else
            If ActiveCell.Value = "" Then
               CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column + 1).Address
               'MsgBox "NÓ DE HIERARQUIAS  é de preenchimento obrigatorio. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
               ERROS = True
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "NÓ DE HIERARQUIA é de preenchimento obrigatorio."
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
               '
               Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
               '
               VALIDACAO_INICIAIS = True
               LINHAS_ERRO = LINHAS_ERRO + 1
          End If
        End If
        '
        CORRE_COL = CORRE_COL + 2
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        'NIVEL SORTIMENTO(GAMA)
        If IsError(ActiveCell.Value) Then
           CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column + 1).Address
           ERROS = True
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrada NIVEL SORTIMENTO(GAMA). O texto não existe em nossas tabelas."
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
           '
           Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
           '
           VALIDACAO_INICIAIS = True
           LINHAS_ERRO = LINHAS_ERRO + 1
        Else
            If ActiveCell.Value = "" Then
               CELL_ADRESS = ActiveCell.Address
               'MsgBox "NIVEL SORTIMENTO (GAMA)  é de preenchimento obrigatorio. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
               ERROS = True
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "NIVEL SORTIMENTO (GAMA) é de preenchimento obrigatorio."
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
               '
               Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
               '
               VALIDACAO_INICIAIS = True
               LINHAS_ERRO = LINHAS_ERRO + 1
           End If
        End If
        '
        CORRE_COL = CORRE_COL + 2
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If IsError(ActiveCell.Value) Then
           CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column + 1).Address
           ERROS = True
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrada GARANTIA. O texto não existe em nossas tabelas."
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
           '
           Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
           '
           VALIDACAO_INICIAIS = True
           LINHAS_ERRO = LINHAS_ERRO + 1
        End If
        '
        CORRE_COL = CORRE_COL + 4
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        'MATERIAL EXCLUSIVO
        '
        If Cells(LIN_CELL, CORRE_COL).Value = "SIM" Then
           '
           CORRE_COL = CORRE_COL + 1
           Cells(LIN_CELL, CORRE_COL).Activate
            '
           If Cells(LIN_CELL, CORRE_COL).Value = "" Or _
              Cells(LIN_CELL, CORRE_COL + 1).Value = "" Then
               CELL_ADRESS = ActiveCell.Address
               'MsgBox "Se MATERIAL EXCLUSIVO = 'SIM' Data INICIO e Data FIM são de preenchimento obrigatório. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"""
               ERROS = True
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "Se MATERIAL EXCLUSIVO = 'SIM' Data INICIO e Data FIM são de preenchimento obrigatório. ."
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
               '
               Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
               '
               VALIDACAO_INICIAIS = True
               LINHAS_ERRO = LINHAS_ERRO + 1
           Else
              If Cells(LIN_CELL, CORRE_COL).Value > Cells(LIN_CELL, CORRE_COL + 1).Value Then
                CELL_ADRESS = ActiveCell.Address
                'MsgBox "Data FIM é menor que Data INICIO. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"""
                ERROS = True
                Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "Data FIM é menor que Data INICIO."
                Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
                Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
                Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
                Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
                Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
                '
                Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
                Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
                '
                VALIDACAO_INICIAIS = True
                LINHAS_ERRO = LINHAS_ERRO + 1
             End If
           End If
        End If
        '
        '
        CORRE_COL = CORRE_COL + 4
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If ActiveCell.Value = "" Then
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "UNIDADE MEDIDA BASICA é de preenchimento obrigatorio. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
           ERROS = True
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "UNIDADE MEDIDA BASICA é de preenchimento obrigatorio."
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
           '
           Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
           '
           VALIDACAO_INICIAIS = True
           LINHAS_ERRO = LINHAS_ERRO + 1
        End If
        '
        CORRE_COL = CORRE_COL + 1
        Cells(LIN_CELL, CORRE_COL).Activate
        '
          If IsError(ActiveCell.Value) Then
             CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column + 1).Address
             ERROS = True
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrada DESCR.FAMILIA INTERNACIONAL. O texto não existe em nossas tabelas."
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
             '
             Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
             '
             VALIDACAO_INICIAIS = True
             LINHAS_ERRO = LINHAS_ERRO + 1
        Else
            If ActiveCell.Value = "" Then
               CELL_ADRESS = ActiveCell.Address
               'MsgBox "DESCR.FAMILIA INTERNACIONAL é de preenchimento obrigatorio. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
               ERROS = True
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "DESCR.FAMILIA INTERNACIONAL é de preenchimento obrigatorio."
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
               '
               Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
               '
               VALIDACAO_INICIAIS = True
               LINHAS_ERRO = LINHAS_ERRO + 1
          End If
        End If
        '
        CORRE_COL = CORRE_COL + 3
        Cells(LIN_CELL, CORRE_COL).Activate
        '
          If IsError(ActiveCell.Value) Then
             CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column + 1).Address
             ERROS = True
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrada DESCR.NIVEL RENOVACAO. O texto não existe em nossas tabelas."
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
             '
             Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
             '
             VALIDACAO_INICIAIS = True
             LINHAS_ERRO = LINHAS_ERRO + 1
          Else
            If ActiveCell.Value = "" Then
               CELL_ADRESS = ActiveCell.Address
               'MsgBox "DESCR.NIVEL RENOVACAO é de preenchimento obrigatorio. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
               ERROS = True
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "DESCR.NIVEL RENOVACAO é de preenchimento obrigatorio."
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
               Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
               '
               Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
               '
               VALIDACAO_INICIAIS = True
               LINHAS_ERRO = LINHAS_ERRO + 1
          End If
        End If
        '
        CORRE_COL = CORRE_COL + 2
        Cells(LIN_CELL, CORRE_COL).Activate
        '
          If IsError(ActiveCell.Value) Then
             CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column + 1).Address
             ERROS = True
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrada DESCR.ECO SUSTENTAVEL. O texto não existe em nossas tabelas."
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
             '
             Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
             '
             VALIDACAO_INICIAIS = True
             LINHAS_ERRO = LINHAS_ERRO + 1
        Else
          If ActiveCell.Value = "" Then
              CELL_ADRESS = ActiveCell.Address
              'MsgBox "DESCR.ECO SUSTENTAVEL é de preenchimento obrigatorio. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
              ERROS = True
              Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "DESCR.ECO SUSTENTAVEL é de preenchimento obrigatorio."
              Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
              Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
              Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
              Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
              Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
              Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
              '
              Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
              Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
              '
              VALIDACAO_INICIAIS = True
              LINHAS_ERRO = LINHAS_ERRO + 1
          End If
        End If
        '
        CORRE_COL = CORRE_COL + 2
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If ActiveCell.Value = "" Then
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "ALTO RISCO? é de preenchimento obrigatorio. 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
           ERROS = True
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "ALTO RISCO? é de preenchimento obrigatorio. 'SIM' ou 'NÃO'."
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
           '
           Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
           '
           VALIDACAO_INICIAIS = True
           LINHAS_ERRO = LINHAS_ERRO + 1
        End If
        '
        CORRE_COL = CORRE_COL + 1
        Cells(LIN_CELL, CORRE_COL).Activate
        '
          If IsError(ActiveCell.Value) Then
             CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column + 1).Address
             ERROS = True
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrada ESTILO. O texto não existe em nossas tabelas."
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
             '
             Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
             '
             VALIDACAO_INICIAIS = True
             LINHAS_ERRO = LINHAS_ERRO + 1
        Else
          If ActiveCell.Value = "" Then
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "ESTILO é de preenchimento obrigatorio. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
             ERROS = True
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "ESTILO é de preenchimento obrigatorio."
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
             '
             Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
             '
             VALIDACAO_INICIAIS = True
             LINHAS_ERRO = LINHAS_ERRO + 1
          End If
        End If
        '
        CORRE_COL = CORRE_COL + 4
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If ActiveCell.Value = "" Then
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "PROJETO CLIENTE é de preenchimento obrigatorio. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
           ERROS = True
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "PROJETO CLIENTE é de preenchimento obrigatorio."
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
           '
           Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
           '
           VALIDACAO_INICIAIS = True
           LINHAS_ERRO = LINHAS_ERRO + 1
        Else
          If Not IsNumeric(ActiveCell.Value) Then
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "PROJETO CLIENTE deve ser NUMERICO. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
             ERROS = True
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "PROJETO CLIENTE deve ser NUMERICO."
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
             Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
             '
             Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
             '
             VALIDACAO_INICIAIS = True
             LINHAS_ERRO = LINHAS_ERRO + 1
          End If
        End If
        '
        CORRE_COL = CORRE_COL + 3
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If ActiveCell.Value = "" Then
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "SAZONALIDADEE é de preenchimento obrigatorio. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Iniciais"
           ERROS = True
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "SAZONALIDADE é de preenchimento obrigatorio."
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
           Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
           '
           Sheets("Dados Cadastrais").VALIDACAO_CC.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Cadastrais").VALIDACAO_CC.ForeColor = &H8000000E  'BRANCO
           '
           VALIDACAO_INICIAIS = True
           LINHAS_ERRO = LINHAS_ERRO + 1
        End If
        '
        On Error GoTo 0
    End If
    LIN_CELL = LIN_CELL + 1
    CORRE_COL = COL_CELL
    Cells(LIN_CELL, CORRE_COL).Activate
    '
Next i
''
If ERROS Or Not VALIDACAO_COD_EAN_UNICO Then
   MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS Cadastrais'.", vbOKOnly, "Dados Fiscais"
   VALIDACAO_INICIAIS = True
   ActiveWindow.ScrollWorkbookTabs Sheets:=1
   ActiveWindow.ScrollWorkbookTabs Sheets:=1
   ActiveWindow.ScrollWorkbookTabs Sheets:=1
   ActiveWindow.ScrollWorkbookTabs Sheets:=1
   ActiveWindow.ScrollWorkbookTabs Sheets:=1
   Sheets("ERROS Cadastrais").Select
   Cells.Select
   Selection.ColumnWidth = 8.29
   Cells.EntireColumn.AutoFit
Else
  Sheets("Dados Fiscais").Select
  MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Fiscais"
End If
'
Sheets("ERROS Cadastrais").Select
Cells.Select
Selection.ColumnWidth = 8.29
Cells.EntireColumn.AutoFit
'
'ActiveWorkbook.Protect Key
'
End Function


Sub CARREGA_FISCAIS()
'On Error GoTo ERRO
'
Dim COL_CADASTRAIS As Integer
Dim LIN_CADASTRAIS As Integer
Dim COL_PAR As Integer
Dim LIN_PAR As Integer
Dim COL_FISCAIS As Integer
Dim LIN_FISCAIS As Integer
Dim COD_MATERIAL As String
Dim CELL_FIND As String
Dim QTDE_LINHAS_LIBERADAS As Integer
Dim i As Integer

'
Dim localizador As Range
'
Dim ENDER As String
'
Dim DESCR_MATERIAL As String
Dim DESCR_MATERIAL_FISCAIS As String
Dim COD_EAN As String
Dim COD_EAN_FISCAIS As String
'
Application.ScreenUpdating = False
Application.EnableEvents = False
'
ActiveWorkbook.Sheets("PARAMETROS").Select
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value
'
'ULTIMA LINHA DADOS FISCAIS
LIN_FISCAIS = Sheets("Dados Fiscais").Cells(Rows.Count, "C").End(xlUp).Offset(1, 0).Row
'
ActiveWorkbook.Sheets("PARAMETROS").Select
Application.Goto Reference:="INICIO_PRACAS"
'
COL_PAR = ActiveCell.Column
LIN_PAR = ActiveCell.Row
'
Sheets("Dados Cadastrais").Select
Application.Goto Reference:="COD_MATERIAL" 'DADOS CADASTRAIS PRIMEIRO CODIGO DO FORNECEDOR

COL_CADASTRAIS = ActiveCell.Column
LIN_CADASTRAIS = ActiveCell.Row
'
Conta_fiscais.Show vbModeless
'
ThisWorkbook.Sheets("Dados Fiscais").Unprotect Key 'DESPROTEGE PLANILHA
'
For i = 1 To QTDE_LINHAS_LIBERADAS
    ''''''''''''''''''''''''''''''''''''''''''''''
    'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
    If ThisWorkbook.Sheets("Dados Cadastrais").Cells(LIN_CADASTRAIS, 2).Value = 0 Then
       Exit For
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''
    '
    If ThisWorkbook.Sheets("Dados Cadastrais").Cells(LIN_CADASTRAIS, COL_CADASTRAIS).Value <> "" Then
       Do While ActiveCell.Value <> ""
          Conta_fiscais.NUM_REGISTRO = LIN_FISCAIS
          Conta_fiscais.Repaint
          '
          Sheets("Dados Fiscais").Cells(LIN_FISCAIS, "B") = Format(Sheets("Dados Cadastrais").Cells(LIN_CADASTRAIS, COL_CADASTRAIS + 2).Value, "###############") 'COD.EAN
          Sheets("Dados Fiscais").Cells(LIN_FISCAIS, "C") = Sheets("Dados Cadastrais").Cells(LIN_CADASTRAIS, COL_CADASTRAIS).Value 'COD.MATERIAL
          Sheets("Dados Fiscais").Cells(LIN_FISCAIS, "D") = Sheets("Dados Cadastrais").Cells(LIN_CADASTRAIS, COL_CADASTRAIS + 1).Value 'DESCR.MATERIAIS
          Sheets("Dados Fiscais").Cells(LIN_FISCAIS, "E") = _
          Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR) 'CENTRO
          Sheets("Dados Fiscais").Cells(LIN_FISCAIS, "F") = _
                Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR + 1) 'NOME CENTRO
          Sheets("Dados Fiscais").Cells(LIN_FISCAIS, "G") = _
                Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR + 2) 'REGIÃO DE ENTRADA
          Sheets("Dados Fiscais").Cells(LIN_FISCAIS, "H") = _
                Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR + 3) 'COD.PERFILDISTRIBUICAO
          Sheets("Dados Fiscais").Cells(LIN_FISCAIS, "I") = _
                Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR + 4) 'DESCR.PERFILDISTRIBUICAO
          Sheets("Dados Fiscais").Cells(LIN_FISCAIS, "J") = _
                Right("'0" & Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR + 5), 10) 'COD.FORNEC SAP
          Sheets("Dados Fiscais").Cells(LIN_FISCAIS, "Y") = 100
          '
          LIN_PAR = LIN_PAR + 1
          LIN_FISCAIS = LIN_FISCAIS + 1
          Sheets("PARAMETROS").Select
          Cells(LIN_PAR, COL_PAR).Activate
        Loop
        ThisWorkbook.Sheets("PARAMETROS").Select
        Application.Goto Reference:="INICIO_PRACAS"
        '
        COL_PAR = ActiveCell.Column
        LIN_PAR = ActiveCell.Row
        Cells(LIN_PAR, COL_PAR).Activate
        '
        LIN_CADASTRAIS = LIN_CADASTRAIS + 1
        ThisWorkbook.Sheets("Dados Cadastrais").Select
        Cells(LIN_CADASTRAIS, COL_CADASTRAIS).Activate 'COD.MATERIAL
        '
    End If
Next i
'
'
        '''''''''''''''''''''''''''''''''''
        Unload Conta_fiscais
        '''''''''''''''''''''''''''''''''''

Sheets("Dados Fiscais").INICIO.Top = 127.5
'
'SALVA PLANILHA
'
'
'marca como gerado dados fiscais na planilha trabalho
If Sheets("TRABALHO").Visible = False Then
   ActiveWorkbook.Unprotect Key
   Sheets("TRABALHO").Visible = True
End If
ActiveWorkbook.Sheets("TRABALHO").Select
Application.Goto Reference:="DADOS_FISCAIS_FIRSTONE"
ActiveCell.Value = "X"
'ActiveWorkbook.Save
Sheets("Dados Fiscais").Select
Range("A1").Select
If Sheets("TRABALHO").Visible = True Then
   Sheets("TRABALHO").Visible = False
   'ActiveWorkbook.Protect Key
End If
'
Application.Goto Reference:="INICIO_ESCOLHA_SUBSORTIMENTO"
If ActiveCell.Value <> "" Then
    Sheets("Inicio").PERFIL.BackColor = &HC000& 'VERDE
    Sheets("Inicio").PERFIL.ForeColor = &H8000000E  'BRANCO
Else
    Sheets("Inicio").PERFIL.BackColor = &HFF& 'vermelho
    Sheets("Inicio").PERFIL.ForeColor = &H8000000E  'BRANCO
End If
'
Application.EnableEvents = True
'
ThisWorkbook.Sheets("Inicio").Select
'ThisWorkbook.Sheets("Dados Fiscais").Protect Key 'PROTEGE PLANILHA
ThisWorkbook.Sheets("Dados Fiscais").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
    False
Exit Sub
'ERRO:
'sgBox ("ERRO OCORRIDO")
End Sub

Sub VALIDACAO_FISCAIS()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim PERCENT_CEL As Double
Dim COD_MATERIAL As String
Dim PERCENT_TOTAL As Double
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
Dim ULT_LIN_ERRO As Integer
'
Dim i As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA Composicao
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
If Sheets("PARAMETROS").Visible = False Then
   ActiveWorkbook.Unprotect Key
   Sheets("PARAMETROS").Visible = True
End If
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BUSCA ULTIMA LINHA DA TABELA DE "ERROS em Dados Regulatorios" POR QUE JÁ
'FOI ABASTECIDA PELOS ERROS DO QUADRO PERIGOSO
'
'
Sheets("ERROS Dados Fiscais").Select
'ActiveSheet.Unprotect Key
Cells.Select
Selection.ClearContents
'
Range("A1").Select
ActiveCell.FormulaR1C1 = "MENSAGEM"
Range("B1").Select
ActiveCell.FormulaR1C1 = "PLANILHA"
Range("C1").Select
ActiveCell.FormulaR1C1 = "QUADRO"
Range("D1").Select
ActiveCell.FormulaR1C1 = "INTERVALO"
Range("A2").Select
LINHAS_ERRO = 2 'SEGUNDA LINHA
'
Conta_fiscais.Show vbModeless
'
COL_CELL = ActiveCell.Column
LIN_CELL = ActiveCell.Row
CORRE_COL = ActiveCell.Column
'
Application.Goto Reference:="VALIDA_SO_ESTA_FISCAIS"
If ActiveCell.FormulaR1C1 = "TRUE" Then
'
    Sheets("Dados Fiscais").Select
    Application.Goto Reference:="INICIO_FISCAIS" 'POSICIONADO NO CABECALHO EM CODIGO EAN
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row + 1 'POR ISSO PULA LINHA
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Inicio").FISCAIS.BackColor = &HC000& 'VERDE
    Sheets("Inicio").FISCAIS.ForeColor = &H8000000E  'BRANCO
    Sheets("Dados Fiscais").NAO_VALIDADO.Visible = False 'ESCONDE BOTAO
    ERROS = False
    '
    Do While ActiveCell.Value <> ""
        '
         CORRE_COL = CORRE_COL + 8
         Cells(LIN_CELL, CORRE_COL).Activate
         '
         If ActiveCell.Value = "" Then 'COD.FORNEC.REGIÃO
            CELL_ADRESS = ActiveCell.Address
            'MsgBox "O campo COD.FORNEC.REGIÃO deve ser preenchido. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
            ERROS = True
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo COD.FORNEC.REGIÃO deve ser preenchido."
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
            Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
            Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
            LINHAS_ERRO = LINHAS_ERRO + 1
         End If
         '
         '
         CORRE_COL = CORRE_COL + 1
         Cells(LIN_CELL, CORRE_COL).Activate
         '
         If IsError(ActiveCell.Value) Then
            CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column).Address
            ERROS = True
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo UNIDADE COMPRAS contem um erro. Verifique."
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
            Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
            Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
            '
            LINHAS_ERRO = LINHAS_ERRO + 1
        Else
         
            If ActiveCell.Value = "" Then 'UNIDADE COMPRAS
               CELL_ADRESS = ActiveCell.Address
               'MsgBox "O campo UNIDADE COMPRAS deve ser preenchido. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
               ERROS = True
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo UNIDADE COMPRAS deve ser preenchido."
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
               Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
               Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
               LINHAS_ERRO = LINHAS_ERRO + 1
            End If
         End If
         '
         CORRE_COL = CORRE_COL + 1
         Cells(LIN_CELL, CORRE_COL).Activate
         '
        If IsError(ActiveCell.Value) Then
           CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column + 1).Address
           ERROS = True
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrada NCM. O texto não existe em nossas tabelas."
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
           Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
           Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
           '
           LINHAS_ERRO = LINHAS_ERRO + 1
        Else
          '
          If ActiveCell.Value = "" Then 'NCM
             CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column + 1).Address
             'MsgBox "O campo NCM deve ser preenchido. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
             ERROS = True
             Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo NCM deve ser preenchido."
             Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
             Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
             Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
             Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
             Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
             LINHAS_ERRO = LINHAS_ERRO + 1
          End If
        End If
         '
        CORRE_COL = CORRE_COL + 3
        Cells(LIN_CELL, CORRE_COL).Activate
         '
        If IsError(ActiveCell.Value) Then
           CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column + 1).Address
           ERROS = True
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrada CEST. O texto não existe em nossas tabelas."
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
           Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
           Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
           LINHAS_ERRO = LINHAS_ERRO + 1
        End If
        '
         CORRE_COL = CORRE_COL + 2
         Cells(LIN_CELL, CORRE_COL).Activate
'        '
         If IsError(ActiveCell.Value) Then
            CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column).Address
            ERROS = True
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo PRECO UNITARIO ESTOQUE contem um erro. Verifique."
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
            Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
            Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
            '
            LINHAS_ERRO = LINHAS_ERRO + 1
        Else
            If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'PRECO UNITARIO ESTOQUE
               CELL_ADRESS = ActiveCell.Address
               'MsgBox "O campo PRECO UNITARIO ESTOQUE deve ser preenchido. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
               ERROS = True
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo PRECO UNITARIO ESTOQUE deve ser preenchido."
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
               Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
               Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
               LINHAS_ERRO = LINHAS_ERRO + 1
            Else
               If InStr(ActiveCell.Value, ",") > 0 Then
                    If Len(Split(ActiveCell.Value, ",")(1)) > 2 Then
                       ERROS = True
                       Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo PRECO UNITARIO ESTOQUE deve ter apenas DUAS casas decimais."
                       Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
                       Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
                       Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                       Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
                       Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
                       Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
                       Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
                       Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
                       LINHAS_ERRO = LINHAS_ERRO + 1
                    End If
               End If
            End If
        End If
        '
        CORRE_COL = CORRE_COL + 1
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If IsError(ActiveCell.Value) Then
            CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column).Address
            ERROS = True
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo PRECO TOTAL COMPRAS contem um erro. Verifique."
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
            Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
            Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
            '
            LINHAS_ERRO = LINHAS_ERRO + 1
        Else
             If IsError(ActiveCell.Value) Then
                CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column).Address
                ERROS = True
                Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo PRECO  TOTAL COMPRAS contem um erro. Verifique."
                Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
                Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
                Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
                Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
                Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
                Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
                Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
                '
                LINHAS_ERRO = LINHAS_ERRO + 1
            Else
                If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'PRECO  TOTAL COMPRAS
                   CELL_ADRESS = ActiveCell.Address
                   'MsgBox "O campo PRECO TOTAL COMPRAS deve ser preenchido. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
                   ERROS = True
                   Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo PRECO TOTAL COMPRAS deve ser preenchido."
                   Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
                   Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
                   Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                   Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
                   Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
                   Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
                   Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
                   Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
                   LINHAS_ERRO = LINHAS_ERRO + 1
                End If
            End If
        End If
        '
         '
         CORRE_COL = CORRE_COL + 1
         Cells(LIN_CELL, CORRE_COL).Activate
         '
         If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'FABRICADO OU REVENDIDO
            CELL_ADRESS = ActiveCell.Address
            'MsgBox "O campo FABRICADO OU REVENDIDO deve ser preenchido. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
            ERROS = True
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo FABRICADO OU REVENDIDO deve ser preenchido."
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
            Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
            Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
            LINHAS_ERRO = LINHAS_ERRO + 1
         End If
         '
         CORRE_COL = CORRE_COL + 1
         Cells(LIN_CELL, CORRE_COL).Activate
         '
         If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'SIMPLES NACIONAL
            CELL_ADRESS = ActiveCell.Address
            'MsgBox "O campo SIMPLES NACIONAL deve ser preenchido com 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
            ERROS = True
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo SIMPLES NACIONAL deve ser preenchido com 'SIM' ou 'NÃO'."
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
            Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
            Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
            LINHAS_ERRO = LINHAS_ERRO + 1
         End If
         '
         CORRE_COL = CORRE_COL + 1
         Cells(LIN_CELL, CORRE_COL).Activate
         '
         If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'DIFERIMENTO ICMS
            CELL_ADRESS = ActiveCell.Address
            'MsgBox "O campo DIFERIMENTO ICMS deve ser preenchido com 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
            ERROS = True
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo DIFERIMENTO ICMS deve ser preenchido com 'SIM' ou 'NÃO'."
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
            Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
            Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
            LINHAS_ERRO = LINHAS_ERRO + 1
         End If
         
        CORRE_COL = CORRE_COL + 1
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        'PAUTA DE PRECO
        '
         If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'PAUTA DE PRECO
            CELL_ADRESS = ActiveCell.Address
            'MsgBox "O campo PAUTA DE PRECO deve ser preenchido com 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
            ERROS = True
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo PAUTA DE PRECO deve ser preenchido com 'SIM' ou 'NÃO'."
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
            Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
            Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
            LINHAS_ERRO = LINHAS_ERRO + 1
         End If
        '
        CORRE_COL = CORRE_COL + 1
        Cells(LIN_CELL, CORRE_COL).Activate
        '
         If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'ISENTO OU IMUNE ICMS
            CELL_ADRESS = ActiveCell.Address
            'MsgBox "O campo ISENTO OU IMUNE ICMS deve ser preenchido com 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
            ERROS = True
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo ISENTO OU IMUNE ICMS? deve ser preenchido com 'SIM' ou 'NÃO'."
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
            Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
            Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
            Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
            Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
            Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
            LINHAS_ERRO = LINHAS_ERRO + 1
         End If
        '
        CORRE_COL = CORRE_COL + 1
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If ActiveCell.Value = "" Or ActiveCell.Value = 0 Then 'ISENTO PIS E COFINS
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "O campo ISENTO PIS E COFINS deve ser preenchido com 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
           ERROS = True
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo ISENTO PIS E COFINS? deve ser preenchido com 'SIM' ou 'NÃO'. "
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
           Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
           Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
           LINHAS_ERRO = LINHAS_ERRO + 1
        End If
        '
        CORRE_COL = CORRE_COL + 1
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If ActiveCell.Value = "" Then '%REDUCAO ICMS
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "O campo %REDUCAO ICMS deve ser preenchido. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
           ERROS = True
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo %REDUCAO ICMS deve ser preenchido."
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
           Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
           Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
           LINHAS_ERRO = LINHAS_ERRO + 1
        End If
        '
        CORRE_COL = CORRE_COL + 1
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If ActiveCell.Value = "" Or ActiveCell.Value = 0 And _
           Cells(ActiveCell.Row, 23) = "NÃO" Then 'ALIQUOTA ICMS e IMUNE ICMS?
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "O campo ALIQUOTA ICMS deve ser preenchido. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
           ERROS = True
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo ALIQUOTA ICMS deve ser preenchido."
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
           Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
           Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
           LINHAS_ERRO = LINHAS_ERRO + 1
        End If
        '
        CORRE_COL = CORRE_COL + 1
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If ActiveCell.Value = "" Then  '%ALIQUOTA IPI
           CELL_ADRESS = ActiveCell.Address
           'MsgBox "O campo %ALIQUOTA IPI deve ser preenchido. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
           ERROS = True
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo %ALIQUOTA IPI deve ser preenchido. Com zeros se não incidir."
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
           Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
           Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
           LINHAS_ERRO = LINHAS_ERRO + 1
        End If
        '
        CORRE_COL = CORRE_COL + 1
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If IsError(ActiveCell.Value) Then
           CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column).Address
           ERROS = True
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrada SUBSTITUICAO TRIBUTARIA. O texto não existe em nossas tabelas."
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
           Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
           Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
           LINHAS_ERRO = LINHAS_ERRO + 1
        Else
            If ActiveCell.Value = "" Then 'SUBSTITUICAO TRIBUTARIA
               CELL_ADRESS = ActiveCell.Address
               'MsgBox "O campo SUBSTITUICAO TRIBUTARIA deve ser preenchido. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
               ERROS = True
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo SUBSTITUICAO TRIBUTARIA deve ser preenchido."
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
               Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
               Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
               LINHAS_ERRO = LINHAS_ERRO + 1
            End If
        End If
        '
        CORRE_COL = CORRE_COL + 4
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        If IsError(ActiveCell.Value) Then
           CELL_ADRESS = Cells(ActiveCell.Row, ActiveCell.Column + 1).Address
           ERROS = True
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "Não foi encontrada ORIGEM MATERIAL. O texto não existe em nossas tabelas."
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
           Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
           Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
           Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
           Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
           Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
           LINHAS_ERRO = LINHAS_ERRO + 1
        Else
            If ActiveCell.Value = "" Then 'ORIGEM MATERIAL
               CELL_ADRESS = ActiveCell.Address
               'MsgBox "O campo ORIGEM MATERIAL deve ser preenchido. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Fiscais"
               ERROS = True
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "O campo ORIGEM MATERIAL deve ser preenchido."
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
               Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
               Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
               Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
               Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
               LINHAS_ERRO = LINHAS_ERRO + 1
            End If
        End If
        '
        'BUSCA CODIGO IVA
        SUGERE_IVA (LIN_CELL)
        '
        LIN_CELL = LIN_CELL + 1
        CORRE_COL = COL_CELL
        Cells(LIN_CELL, CORRE_COL).Activate
        '
        Conta_fiscais.NUM_REGISTRO = LIN_CELL
        Conta_fiscais.Repaint
        '
    Loop
End If
'
'
'''''''''''''''''''''''''''''''''''
Unload Conta_fiscais
'''''''''''''''''''''''''''''''''''
'
If ERROS Or Not VALIDACAO_SUBSORT_UNICO Then
   MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS Dados Fiscais'.", vbOKOnly, "Dados Fiscais"
   ActiveWindow.ScrollWorkbookTabs Sheets:=1
   ActiveWindow.ScrollWorkbookTabs Sheets:=1
   ActiveWindow.ScrollWorkbookTabs Sheets:=1
   ActiveWindow.ScrollWorkbookTabs Sheets:=1
   ActiveWindow.ScrollWorkbookTabs Sheets:=1
   Sheets("ERROS Dados Fiscais").Select
   Cells.Select
   Selection.ColumnWidth = 8.29
   Cells.EntireColumn.AutoFit
Else
  Sheets("Dados Fiscais").Select
  MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Fiscais"
End If
'
''Application.ScreenUpdating = True
End Sub


Sub ESCONDE_PROTEGE_FORNECEDOR()
    '
    'On Error Resume Next
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    'FISCAIS.Enaled = False
    'SÓ SERA LIBERADO SE CADASTRAIS E PERFIL PREENCHIDOS COM EXITO
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Sheets("Inicio").PERFIL.BackColor = &HC000& And _
       Sheets("Inicio").CADASTRAIS.BackColor = &HC000& Then
       If Sheets("Inicio").FISCAIS.Enabled = False Then
          Sheets("Inicio").FISCAIS.Enabled = True
       End If
    Else
       If Sheets("Inicio").FISCAIS.Enabled = True Then
          Sheets("Inicio").FISCAIS.Enabled = False
       End If
    End If
    '
    Application.ScreenUpdating = False
    '
    Application.EnableEvents = False
    '
    If Sheets("Dados Cadastrais").Visible = False Then
       Sheets("Dados Cadastrais").Visible = True
    End If
    Sheets("ERROS Cadastrais").Visible = True
    If Sheets("Composicao").Visible = False Then
       Sheets("Composicao").Visible = True
    End If
    Sheets("ERROS de Composicao").Visible = True
'    If Sheets("PARAMETROS").Visible = True Then
'       Sheets("PARAMETROS").Visible = False
'     End If
     '
     If Sheets("Dados Regulatorios").Visible = False Then
        Sheets("Dados Regulatorios").Visible = True
     End If
    Sheets("ERROS em Dados Regulatorios").Visible = True
     If Sheets("Embalagens").Visible = False Then
        Sheets("Embalagens").Visible = True
     End If
    '
    Sheets("ERROS em Embalagens").Visible = True
    If Sheets("Dados Qualidade").Visible = False Then
       Sheets("Dados Qualidade").Visible = True
    End If
    '
    Sheets("ERROS em Dados Qualidade").Visible = True
    '
    Sheets("ERROS em Embalagens").Visible = True
    If Sheets("Dados Fiscais").Visible = False Then
       Sheets("Dados Fiscais").Visible = True
    End If
    '
    Sheets("ERROS Dados Fiscais").Visible = True
    '
    If Sheets("Dados Iniciais").Visible = True Then
        Sheets("Dados Iniciais").Visible = False
        Sheets("adobeCemt").Visible = False
        Sheets("Dados Gerais 1").Visible = False
        Sheets("Dados Gerais 2").Visible = False
        Sheets("Dados Gerais 3").Visible = False
        Sheets("Dados Gerais 4").Visible = False
        Sheets("Unidades Medida").Visible = False
        Sheets("Dados Compra").Visible = False
    End If
    '
    Sheets("Dados Cadastrais").Select
    Sheets("Dados Cadastrais").VALIDACAO_CC.Visible = False
    ThisWorkbook.Sheets("Dados Cadastrais").Unprotect Key
    Application.Goto Reference:="DADOS_INICIAIS"
    '
    'ESCONDE CAMPOS DE REPONSABILIDADE DO PORTAL
    Columns("F:G").Select 'DESCR.COMERCIAL
    Selection.EntireColumn.Hidden = True
    Columns("I:P").Select 'CONTROLE LOTE? GRUPO MERCADORIAS   DESCR.GRUPO MERCADORIAS COD.NÓ DE HIERARQUIA    NÓ DE HIERARQUIA    COD.NIVEL SORTIMENTO(GAMA)  NIVEL SORTIMENTO (GAMA) COD.GARANTIA
    Selection.EntireColumn.Hidden = True
    Columns("R:V").Select 'NOVIDADE MERCADO ATÉ    MATERIAL TESTE ATÉ  MATERIAL EXCLUSIVO  INICIO  FIM
    Selection.EntireColumn.Hidden = True
    Columns("X:AO").Select 'UNIDADE MEDIDA BASICA   COD.FAMILIA INTERNACIONAL   DESCR.FAMILIA INTERNACIONAL DESCRICAO MATERIAL SAP (LMB)    COD.NIVEL RENOVACAO DESCR.NIVEL RENOVACAO   COD.ECO SUSTENTAVEL DESCR.ECO SUSTENTAVEL   ALTO RISCO? COD.ESTILO  ESTILO  COD.LM SUBSTITUTO   DESCR.LM SUBSTITUTO PROJETO CLIENTE REF PRODUTO HISTORICO   %PORCENTAGEM REF. PRODUTO   SAZONALIDADE    LM GERADO SAP
    Selection.EntireColumn.Hidden = True
    '
    Application.Goto Reference:="DADOS_INICIAIS"
    '
    ThisWorkbook.Sheets("Dados Cadastrais").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    '
    Sheets("Composicao").Select
    ThisWorkbook.Sheets("Composicao").Unprotect Key
    'CODIGO DE CUIDADOS
    '
    Columns("E:F").Select
    If Selection.EntireColumn.Hidden = False Then
       Selection.EntireColumn.Hidden = True
    End If
    '
    Columns("H:H").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:J").Select
    Selection.EntireColumn.Hidden = True
    Columns("L:L").Select
    Selection.EntireColumn.Hidden = True
    Columns("N:N").Select
    Selection.EntireColumn.Hidden = True
    Columns("P:P").Select
    Selection.EntireColumn.Hidden = True
    Columns("R:R").Select
    Selection.EntireColumn.Hidden = True
    '''''''''''''''''''''''''''''''''''''
    'COMPOSICAO MATERIAL
    '''''''''''''''''''''''''''''''''''''
    ESCONDE_M0STRA_COMPOSICAO (True) 'ESCONDE HIDDEN =
    '
    Range("A1").Select
    '
    ThisWorkbook.Sheets("Composicao").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
    False
    '
    Sheets("Dados Regulatorios").Select
    ThisWorkbook.Sheets("Dados Regulatorios").Unprotect Key
    '
    Columns("E:F").Select
    If Selection.EntireColumn.Hidden = False Then
       Selection.EntireColumn.Hidden = True
    End If
    '
    Columns("H:H").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:J").Select
    Selection.EntireColumn.Hidden = True
    Columns("Q:Q").Select
    Selection.EntireColumn.Hidden = True
    Columns("S:S").Select
    Selection.EntireColumn.Hidden = True
    Columns("AR:AR").Select
    Selection.EntireColumn.Hidden = True
    Columns("BL:BL").Select
    Selection.EntireColumn.Hidden = True
    Columns("BP:BP").Select
    Selection.EntireColumn.Hidden = True
    '
    Range("A1").Select
    '
    ThisWorkbook.Sheets("Dados Regulatorios").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
    False
    '
    Sheets("Embalagens").Select
    ThisWorkbook.Sheets("Embalagens").Unprotect Key
    '
    Columns("E:F").Select
    If Selection.EntireColumn.Hidden = False Then
       Selection.EntireColumn.Hidden = True
    End If
    '
    Columns("Q:Q").Select
    If Selection.EntireColumn.Hidden = False Then
       Selection.EntireColumn.Hidden = True
    End If
    '
    Range("A1").Select
    '
    ThisWorkbook.Sheets("Embalagens").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
    False
    '
    Sheets("Dados Qualidade").Select
    ThisWorkbook.Sheets("Dados Qualidade").Unprotect Key
        '
    Columns("E:F").Select
    If Selection.EntireColumn.Hidden = False Then
       Selection.EntireColumn.Hidden = True
    End If
    '
    ESCONDE_M0STRA_QUALIDADE (True) ' HIDDEN =
    '
    Range("A1").Select
    '
    ThisWorkbook.Sheets("Dados Qualidade").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
    False
    '
    Sheets("Dados Fiscais").Select
    ThisWorkbook.Sheets("Dados Fiscais").Unprotect Key

    'CODIGO SUBSTITUICAO TRIBUTARIA
    'VAI APARECER PARA FORNECEDOR
    Columns("AB:AB").Select
    Selection.EntireColumn.Hidden = True
    '
    Columns("AF:AF").Select
    Selection.EntireColumn.Hidden = True
    '
    Columns("AI:AI").Select
    Selection.EntireColumn.Hidden = False
    Columns("AJ:AJ").Select
    Selection.EntireColumn.Hidden = False
    '
    Columns("AK:AK").Select
    Selection.EntireColumn.Hidden = True
    Columns("AL:AL").Select
    Selection.EntireColumn.Hidden = True
    Columns("AM:AM").Select
    Selection.EntireColumn.Hidden = True
    Range("A1").Select
    '
    ThisWorkbook.Sheets("Dados Fiscais").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
    False
    '
    Sheets("Escolha Perfil Distribuicao").Select
    ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Unprotect Key
    Cells.Select
    Selection.Locked = True
    Range("H5:H125").Select
    Selection.Locked = False
    Selection.FormulaHidden = False
    '
    ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
    False
    '
    Application.EnableEvents = True
End Sub

Sub ESCONDE_PROTEGE_LEROY()
    '
    Dim teste
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    'FISCAIS.Enaled = TRUE
    'PARA FORNECEDORES:
    'SÓ SERA LIBERADO SE CADASTRAIS E PERFIL PREENCHIDOS COM EXITO
    'PARA LEROY:
    'SEMPRE TRUE.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Sheets("Inicio").FISCAIS.Enabled = False Then
       Sheets("Inicio").FISCAIS.Enabled = True
    End If
    '
    Application.ScreenUpdating = False
    '
    Application.EnableEvents = False
    '
    ActiveWorkbook.Unprotect Key
    '
    Sheets("Dados Cadastrais").Visible = True
    Sheets("ERROS Cadastrais").Visible = True
    Sheets("Composicao").Visible = True
    Sheets("ERROS de Composicao").Visible = True
    Sheets("PARAMETROS").Visible = True
    If Sheets("TABELAS").Visible = False Then
       Sheets("TABELAS").Visible = True
    End If
    Sheets("Dados Regulatorios").Visible = True
    Sheets("ERROS em Dados Regulatorios").Visible = True
    Sheets("Embalagens").Visible = True
    Sheets("ERROS em Embalagens").Visible = True
    Sheets("Dados Qualidade").Visible = True
    Sheets("ERROS em Dados Qualidade").Visible = True
    Sheets("Dados FISCAIS").Visible = True
    Sheets("ERROS Dados Fiscais").Visible = True
    Sheets("Dados Iniciais").Visible = True
    Sheets("adobeCemt").Visible = True
    Sheets("Dados Gerais 1").Visible = True
    Sheets("Dados Gerais 2").Visible = True
    Sheets("Dados Gerais 3").Visible = True
    Sheets("Dados Gerais 4").Visible = True
    Sheets("Unidades Medida").Visible = True
    Sheets("Dados Compra").Visible = True
    '
    If Sheets("TRABALHO").Visible = True Then
       ActiveWorkbook.Unprotect Key
       Sheets("TRABALHO").Visible = False
       ActiveWorkbook.Protect Key
    End If
    '
    Sheets("Dados Cadastrais").Select
    Sheets("Dados Cadastrais").VALIDACAO_CC.Visible = True
    '
    ThisWorkbook.Sheets("Dados Cadastrais").Unprotect Key
    Application.Goto Reference:="DADOS_INICIAIS"
    '
    'ESCONDE CAMPOS DE REPONSABILIDADE DO PORTAL
    Columns("F:G").Select 'DESCR.COMERCIAL
    Selection.EntireColumn.Hidden = False
    Columns("I:P").Select 'GRUPO MERCADORIAS   DESCR.GRUPO MERCADORIAS COD.NÓ DE HIERARQUIA    NÓ DE HIERARQUIA    COD.NIVEL SORTIMENTO(GAMA)  NIVEL SORTIMENTO (GAMA) COD.GARANTIA
    Selection.EntireColumn.Hidden = False
    Columns("R:V").Select 'NOVIDADE MERCADO ATÉ    MATERIAL TESTE ATÉ  MATERIAL EXCLUSIVO  INICIO  FIM
    Selection.EntireColumn.Hidden = False
    Columns("X:AO").Select 'UNIDADE MEDIDA BASICA   COD.FAMILIA INTERNACIONAL   DESCR.FAMILIA INTERNACIONAL DESCRICAO MATERIAL SAP (LMB)    COD.NIVEL RENOVACAO DESCR.NIVEL RENOVACAO   COD.ECO SUSTENTAVEL DESCR.ECO SUSTENTAVEL   ALTO RISCO? COD.ESTILO  ESTILO  COD.LM SUBSTITUTO   DESCR.LM SUBSTITUTO PROJETO CLIENTE REF PRODUTO HISTORICO   %PORCENTAGEM REF. PRODUTO   SAZONALIDADE    LM GERADO SAP
    Selection.EntireColumn.Hidden = False
    'BOTAO SALVAR EM DADOS CADASTRAIS
    'ActiveWorkbook.Sheets("Dados Cadastrais").SALVAR_INICIAIS.Enabled = True
    '
    Application.Goto Reference:="DADOS_INICIAIS"
    ThisWorkbook.Sheets("Dados Cadastrais").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
    False
    '
    Sheets("Composicao").Select
    ThisWorkbook.Sheets("Composicao").Unprotect Key

    'CODIGO DE CUIDADOS
    '
    Columns("E:F").Select
    If Selection.EntireColumn.Hidden = True Then
       Selection.EntireColumn.Hidden = False
    End If
    '
    Columns("H:H").Select
    Selection.EntireColumn.Hidden = False
    Columns("J:J").Select
    Selection.EntireColumn.Hidden = False
    Columns("L:L").Select
    Selection.EntireColumn.Hidden = False
    Columns("N:N").Select
    Selection.EntireColumn.Hidden = False
    Columns("P:P").Select
    Selection.EntireColumn.Hidden = False
    Columns("R:R").Select
    Selection.EntireColumn.Hidden = False
    '''''''''''''''''''''''''''''''''''''
    'COMPOSICAO MATERIAL
    '''''''''''''''''''''''''''''''''''''
    ESCONDE_M0STRA_COMPOSICAO (False) 'ESCONDE HIDDEN =
    '
    Range("A1").Select
    ThisWorkbook.Sheets("Composicao").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
    False
    '
    Sheets("Dados Regulatorios").Select
    ThisWorkbook.Sheets("Dados Regulatorios").Unprotect Key
    '
    Columns("E:F").Select
    If Selection.EntireColumn.Hidden = True Then
       Selection.EntireColumn.Hidden = False
    End If
    '
    Columns("H:H").Select
    Selection.EntireColumn.Hidden = False
    Columns("J:J").Select
    Selection.EntireColumn.Hidden = False
    Columns("Q:Q").Select
    Selection.EntireColumn.Hidden = False
    Columns("S:S").Select
    Selection.EntireColumn.Hidden = False
    Columns("BJ:BJ").Select
    Selection.EntireColumn.Hidden = False
    Columns("BN:BN").Select
    Selection.EntireColumn.Hidden = False
    Columns("BR:BR").Select
    Selection.EntireColumn.Hidden = False
    '
    Columns("E:F").Select
    If Selection.EntireColumn.Hidden = True Then
       Selection.EntireColumn.Hidden = False
    End If
    Range("A1").Select
    '
    ThisWorkbook.Sheets("Dados Regulatorios").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
    False
    '
    Sheets("Embalagens").Select
    ThisWorkbook.Sheets("Embalagens").Unprotect Key
    Columns("AK:AK").Select 'COD.CONDICAO ESTOCAGEM
    If Selection.EntireColumn.Hidden = True Then
       Selection.EntireColumn.Hidden = False
    End If
    Columns("AZ:AZ").Select
    If Selection.EntireColumn.Hidden = True Then
        Selection.EntireColumn.Hidden = False
    End If
    Columns("BR:BR").Select
    '
    Range("A1").Select
    ThisWorkbook.Sheets("Embalagens").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
    False
    '
    Sheets("Dados Qualidade").Select
    ThisWorkbook.Sheets("Dados Qualidade").Unprotect Key
    '
    Columns("E:F").Select
    If Selection.EntireColumn.Hidden = True Then
       Selection.EntireColumn.Hidden = False
    End If
    '
    ESCONDE_M0STRA_QUALIDADE (False) ' HIDDEN =
    '
    Range("A1").Select
    '
    ThisWorkbook.Sheets("Dados Qualidade").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
    False
    '
    Sheets("Dados Fiscais").Select
    ThisWorkbook.Sheets("Dados Fiscais").Unprotect Key
    '
    Columns("L:L").Select
    Selection.EntireColumn.Hidden = False
    Columns("O:O").Select
    Selection.EntireColumn.Hidden = False
    '
    Columns("AB:AB").Select
    Selection.EntireColumn.Hidden = False
    '
    Columns("AF:AF").Select
    Selection.EntireColumn.Hidden = False
    '
    Columns("AI:AI").Select
    Selection.EntireColumn.Hidden = False
    Columns("AJ:AJ").Select
    Selection.EntireColumn.Hidden = False
    '
    Columns("AK:AK").Select
    Selection.EntireColumn.Hidden = False
    Columns("AL:AL").Select
    Selection.EntireColumn.Hidden = False
    Range("A1").Select
    ThisWorkbook.Sheets("Dados Fiscais").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
    False
    '
    '
    Sheets("Escolha Perfil Distribuicao").Select
    ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Unprotect Key
    Application.Goto Reference:="INICIO_ESCOLHA_SUBSORTIMENTO"
    If ActiveCell.Value = "" Then
       ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Cells.Select
       Selection.Locked = True
       ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Range("A5:G125").Select
       Selection.Locked = False
       Selection.FormulaHidden = False
       '
    Else
        ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Cells.Select
        Selection.Locked = True
        ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Range("A5:A125").Select
        Selection.Locked = False
        ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Range("Q5:Q101").Select
        Selection.Locked = False
        Selection.FormulaHidden = False
        '
    End If
    ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
    False
    '
    Application.EnableEvents = True
    '
End Sub

Sub VALIDACAO_ESTOQ_COMPR_VEND()
'
'On Error Resume Next
'
Dim CONTA_COMPR_VEND As Integer
'
' VALIDACAO COMPRAS E VENDAS
'
Dim CEL_COMPR As Integer
Dim CEL_VEND As Integer
'
Dim LIN_CELL As Integer
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
Dim ULT_LIN_ERRO As Integer
'
Dim i As Integer
'
Dim ERRO_COPR_VEND As Boolean
'
Dim LIN_TAB As Integer
Dim COL_TAB As Integer
'
Dim LIN_EMBAL As Integer
Dim LIN_UNID  As Integer
'
Dim LIN_Unidades_Medida As Integer
'
Dim SIMNAO As String
Dim resultado
'
Dim CONVERSAO As String
'
'PLANILHA Embalagens
'
Application.ScreenUpdating = False
'
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BUSCA ULTIMA LINHA DA TABELA DE "ERROS em Dados Regulatorios" POR QUE JÁ
'FOI ABASTECIDA PELOS ERROS DO QUADRO PERIGOSO
'
'
Sheets("ERROS em Embalagens").Select
Cells(1, 1).Select ' CELULA A1
ULT_LIN_ERRO = 1
Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
   ULT_LIN_ERRO = ULT_LIN_ERRO + 1
Loop
'
LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'As celulas de TIPO DE MOVIMENACAO: COMPRA E VENDA serão validadas para cada produto(linha) e só poderá
'e deverá acontecer uma vez COMPRA e uma vez VENDA.
'essas celulas estão FIXAS NAS COLUNAS:
'"J"  e "K"  para UNIDADE
'"Z"  e "AA" para CAIXA INNER
'"AN" e "AO" para CAIXA MASTER
'"BB" e "BC" para PALLET
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SE PLANILHA FOR ALTERADA ESTA ROTINA TERÁ QUE SER REFEITA NOS ENDERECOS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
LIN_EMBAL = 5 'PRIMEIRA LINHA EMBALAGENS
LIN_Unidades_Medida = 2  'PRIMEIRA LINHA Unidades Medida A SER GRAVADA
'
'
CEL_COMPR = 0
CEL_VEND = 0
'
'APAGA TABELA NA PLANILHA TRABALHO
'TABELA SERA UTILIZADA PARA EXPORTACAO SAP
'NA PLANILHA Unidades Medida (EMBALAGENS)
'E SERA PREENCHIDA A SEGUIR
'
ActiveWorkbook.Unprotect Key
Sheets("TRABALHO").Visible = True
Sheets("TRABALHO").Unprotect Key
Application.Goto Reference:="EMBAL_EXPORT"
Selection.ClearContents
Application.Goto Reference:="INICIO_EMBAL_COMPR_VEND"
LIN_TAB = ActiveCell.Row
COL_TAB = ActiveCell.Column
'
Sheets("Embalagens").Select
'
For i = 1 To QTDE_LINHAS_LIBERADAS
    Conta_fiscais.NUM_REGISTRO = LIN_EMBAL
    Conta_fiscais.Repaint
    ''''''''''''''''''''''''''''''''''''''''''''''
    'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
    If Cells(LIN_EMBAL, 2).Value = 0 Then
       Exit For
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''
    '
    Cells(LIN_EMBAL, 11).Select 'UNIDADE COMPRA COLUNA "J"
    '
    If ActiveCell.Value = "SIM" Then
       CEL_COMPR = CEL_COMPR + 1
       Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Value = "UNIDADE"
       Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 1).Value = "X"
       '
    End If
    '
    Sheets("Embalagens").Select
    Cells(LIN_EMBAL, 12).Select 'UNIDADE VENDA COLUNA "K"
    '
    If ActiveCell.Value = "SIM" Then
       CEL_VEND = CEL_VEND + 1
       If Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Value = "" Then
          Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Value = "UNIDADE"
          Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 2).Value = "X"
       Else
          Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 2).Value = "X"
       End If
       '
       LIN_TAB = LIN_TAB + 1
       Sheets("TRABALHO").Select
       Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Activate
    End If
    '
    Sheets("Embalagens").Select
    Cells(LIN_EMBAL, 45).Select 'MASTER COMPRA COLUNA "AN"
    '
    If ActiveCell.Value = "SIM" Then
       CEL_COMPR = CEL_COMPR + 1
       If Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Value <> "" Then
          LIN_TAB = LIN_TAB + 1
       End If
       '
       Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Value = "MASTER"
       Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 1).Value = "X"
       '
    End If
    '
    Sheets("Embalagens").Select
    Cells(LIN_EMBAL, 46).Select 'MASTER VENDA COLUNA "AO"
    '
    If ActiveCell.Value = "SIM" Then
       CEL_VEND = CEL_VEND + 1
       If Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Value = "" Then
          Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Value = "MASTER"
          Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 2).Value = "X"
       Else
          LIN_TAB = LIN_TAB + 1
          Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Value = "MASTER"
          Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 2).Value = "X"
       End If
       '
       Sheets("TRABALHO").Select
       Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Activate
    End If
    '
    Sheets("Embalagens").Select
    Cells(LIN_EMBAL, 61).Select 'PALLET COMPRA COLUNA "BC"
    '
    If ActiveCell.Value = "SIM" Then
       CEL_COMPR = CEL_COMPR + 1
       If Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Value <> "" Then
          LIN_TAB = LIN_TAB + 1
       End If
       '
       Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Value = "PALLET"
       Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 1).Value = "X"
       '
    End If
    '
    Sheets("Embalagens").Select
    Cells(LIN_EMBAL, 62).Select 'PALLET VENDA COLUNA "AO"
    '
    If ActiveCell.Value = "SIM" Then
       CEL_VEND = CEL_VEND + 1
       If Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Value = "" Then
          Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Value = "PALLET"
          Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 2).Value = "X"
       Else
          LIN_TAB = LIN_TAB + 1
          Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Value = "PALLET"
          Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 2).Value = "X"
       End If
       '
    End If
    '
    Sheets("Embalagens").INICIO.BackColor = &HC000& 'VERDE
    Sheets("Inicio").EMBALAGENS.BackColor = &HC000& 'VERDE
    ERRO_COPR_VEND = False 'COMECA FALSO SEM ERROS
    '
    If CEL_COMPR = 0 Then 'NENHUMA CELULA POR COMPRA
'       MsgBox "Nenhuma coluna de COMPRA esta assinalada com 'SIM' para O PRODUTO na linha " _
'       & LIN_EMBAL & ". Uma COMPRA e uma VENDA entre UNIDADE, CAIXA MASTER e PALLET deverão ser escolhidas como 'SIM'.", vbOKOnly, "Embalagens - TIPO MOVIMENTACAO COMPRA / VENDA"
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "Nenhuma coluna de COMPRA esta assinalada com 'SIM' para O PRODUTO na linha " _
       & LIN_EMBAL
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Tipo Movimentacao - COMPRA"
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = LIN_EMBAL
       Sheets("Embalagens").INICIO.BackColor = &HFF& 'VERMELHO
       Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
       LINHAS_ERRO = LINHAS_ERRO + 1
       '
       ERRO_COPR_VEND = True
    End If
    '
    If CEL_VEND = 0 Then 'NENHUMA CELULA POR VENDA
'       MsgBox "Nenhuma coluna de VENDA esta assinalada com 'SIM' para O PRODUTO na linha " _
'       & LIN_EMBAL & ". Uma COMPRA e uma VENDA entre UNIDADE, CAIXA MASTER e PALLET deverão ser escolhidas como 'SIM'.", vbOKOnly, "Embalagens - TIPO MOVIMENTACAO COMPRA "
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "Nenhuma coluna de VENDA esta assinalada com 'SIM'. O PRODUTO na linha " & LIN_EMBAL
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Tipo Movimentacao - COMPRA"
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = LIN_EMBAL
       Sheets("Embalagens").INICIO.BackColor = &HFF& 'VERMELHO
       Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
       LINHAS_ERRO = LINHAS_ERRO + 1
       '
       ERRO_COPR_VEND = True
    End If
    '
    If CEL_COMPR > 1 Then 'MAIS DE UMA CELULA POR COMPRA
'       MsgBox "Mais de uma celula de COMPRA esta assinalada com 'SIM' para o PRODUTO na linha " & _
'              LIN_EMBAL & ". Uma COMPRA e uma VENDA entre UNIDADE, CAIXA MASTER e PALLET deverão ser escolhidas como 'SIM'.", vbOKOnly, "Embalagens - TIPO MOVIMENTACAO VENDA"
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "Mais de uma celula de COMPRA esta assinalada com 'SIM' para o PRODUTO na linha " & _
              LIN_EMBAL
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Tipo Movimentacao - COMPRA"
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = LIN_EMBAL
       Sheets("Embalagens").INICIO.BackColor = &HFF& 'VERMELHO
       Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
       LINHAS_ERRO = LINHAS_ERRO + 1
       '
       ERRO_COPR_VEND = True
    End If
    '
    If CEL_VEND > 1 Then ''MAIS DE UMA CELULA POR VENDA
'       MsgBox "Mais de uma celula de VENDA esta assinalada com 'SIM' para o PRODUTO na linha " & _
'              LIN_EMBAL & ". Uma COMPRA e uma VENDA entre UNIDADE, CAIXA MASTER e PALLET deverão ser escolhidas como 'SIM'.", vbOKOnly, "Embalagens - TIPO MOVIMENTACAO VENDA"
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 1).Value = "Mais de uma celula de VENDA esta assinalada com 'SIM' para o PRODUTO na linha " & _
              LIN_EMBAL
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 2).Value = "Embalagens"
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 3).Value = "Tipo Movimentacao - VENDA"
       Sheets("ERROS em Embalagens").Cells(LINHAS_ERRO, 4).Value = LIN_EMBAL
       Sheets("Embalagens").INICIO.BackColor = &HFF& 'VERMELHO
       Sheets("Inicio").EMBALAGENS.BackColor = &HFF& 'VERMELHO
       LINHAS_ERRO = LINHAS_ERRO + 1
       '
       ERRO_COPR_VEND = True
    End If
    '
    'SE NÃO HOUVE ERRO VAMOS EXPORTAR OS DADOS COMPRA E VENDA
    'PARA PLANILHA Unidades Medida (EMBALAGENS)
    'OS DADOS DE ESTOQUE SAIRAM DO QUADRO UNIDADE
    '
    '
    
    If Not ERRO_COPR_VEND Then
        '
       'ESTOQUE - UNIDADE = FORMULAS JA NA PLANILHA
       'COLUNA 12 -   PESO BRUTO
       'COLUNA 13 -   PESO LIQUIDO
       'COLUNA 14 -   UNIDADE PESO
       'COLUNA 15 -   COMPRIMENTO
       'COLUNA 16 -   LARGURA
       'COLUNA 17 -   ALTURA
       'COLUNA 18 -   UNIDADE MEDIDA METROS
       '
       'COMPRAS
       'COLUNA 29 -  PESO BRUTO
       'COLUNA 30 -  PESO LIQUIDO
       'COLUNA 31 -  UNIDADE PESO
       'COLUNA 32 -  COMPRIMENTO
       'COLUNA 33 -  LARGURA
       'COLUNA 34 -  ALTURA
       'COLUNA 35 -  UNIDADE MEDIDA METROS
       '
       'VENDAS
       'COLUNA 47 -  PESO BRUTO
       'COLUNA 48 -  PESO LIQUIDO
       'COLUNA 49 -  UNIDADE PESO
       'COLUNA 50 -  COMPRIMENTO
       'COLUNA 51 -  LARGURA
       'COLUNA 52 -  ALTURA
       'COLUNA 53 -  UNIDADE MEDIDA METROS
       '
       'DESPROTEJE AS EXPRTACOES
       '
       '
    Sheets("Dados Iniciais").Visible = True
    'ThisWorkbook.Sheets("Dados Iniciais").Unprotect Key
    Sheets("Dados Iniciais").Select
    '
    Sheets("adobeCemt").Visible = True
    'ThisWorkbook.Sheets("adobeCemt").Unprotect Key
    Sheets("adobeCemt").Select
    '
    Sheets("Dados Gerais 2").Visible = True
    'ThisWorkbook.Sheets("Dados Gerais 2").Unprotect Key
    Sheets("Dados Gerais 2").Select
    '
    Sheets("Dados Compra").Visible = True
    'ThisWorkbook.Sheets("Dados Compra").Unprotect Key
    Sheets("Dados Compra").Select
    '
    Sheets("Dados Gerais 4").Visible = True
    'ThisWorkbook.Sheets("Dados Gerais 4").Unprotect Key
    Sheets("Dados Gerais 4").Select
    '
    Sheets("Unidades Medida").Visible = True
    'ThisWorkbook.Sheets("Unidades Medida").Unprotect Key
    Sheets("Unidades Medida").Select
    '
    Sheets("Dados Compra").Visible = True
    'ThisWorkbook.Sheets("Dados Compra").Unprotect Key
    Sheets("Dados Compra").Select
    '
    ThisWorkbook.Sheets("Unidades Medida").Unprotect Key
    '
    Sheets("TRABALHO").Select
    Application.Goto Reference:="INICIO_EMBAL_COMPR_VEND"
    LIN_TAB = ActiveCell.Row
    COL_TAB = ActiveCell.Column
    '
       Do While ActiveCell.Value <> ""
            Select Case ActiveCell.Value
            Case "UNIDADE"
                 If Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 1).Value = "X" Then 'COMPRAS
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 23).Value = _
                    Sheets("Embalagens").Cells(LIN_EMBAL, 8).Value 'UMB ALTERNATIVA
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 24).Value = "=" 'SIMBOLO = IGUAL
                    '
                    CONVERSAO = Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 24).Value, ",", ".") 'CONTADOR CONVERSAO
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 25).Value = CONVERSAO

                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 26).Value = _
                    Sheets("Embalagens").Cells(LIN_EMBAL, 8).Value 'UIDADE MEDIDA BASICA
                    '
                    If Sheets("Embalagens").Cells(LIN_EMBAL, 21).Value = "G" Then
                        Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 29).Value = _
                            Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 22).Value / 1000, 3), ",", ".") 'PESO BRUTO
                        Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 30).Value = _
                            Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 23).Value / 1000, 3), ",", ".") 'PESO LIQUIDO
                    Else
                        Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 29).Value = _
                            Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 22).Value, 3), ",", ".") 'PESO BRUTO
                        Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 30).Value = _
                            Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 23).Value, 3), ",", ".") 'PESO LIQUIDO
                    End If
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 31).Value = "KG"
                    'Sheets("Embalagens").Cells(LIN_EMBAL, 21).Value 'UNIDADE PESO
                    '
                    If Sheets("Embalagens").Cells(LIN_EMBAL, 13).Value = "MM" Then
                       'SE NA PLANILHA OS VALORES FORAM INFORMADOS EM "MM"
                       'SÃO CONVERTIDOS EM METROS
                       CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 14).Value / 1000), ",", ".")
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 32).Value = _
                            CONVERSAO 'COMPRIMENTO
                       '
                       CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 15).Value / 1000), ",", ".")
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 33).Value = _
                            CONVERSAO 'LARGURA
                       '
                       CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 16).Value / 1000), ",", ".")
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 34).Value = _
                            CONVERSAO 'ALTURA
                       '
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 35).Value = "M" 'UNIDADE MEDIDA METROS
                    Else
                      If Sheets("Embalagens").Cells(LIN_EMBAL, 13).Value = "CM" Then
                         'SE NA PLANILHA OS VALORES FORAM INFORMADOS EM "CM"
                         'SÃO CONVERTIDOS EM METROS
                         CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 14).Value / 100), ",", ".")
                         Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 32).Value = _
                              CONVERSAO 'COMPRIMENTO
                         '
                         CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 15).Value / 100), ",", ".")
                         Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 33).Value = _
                              CONVERSAO 'LARGURA
                         '
                         CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 16).Value / 100), ",", ".")
                         Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 34).Value = _
                              CONVERSAO 'ALTURA
                         '
                         Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 35).Value = "M" 'UNIDADE MEDIDA METROS
                      Else
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 32).Value = _
                                Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 14).Value, ",", ".") 'COMPRIMENTO
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 33).Value = _
                                Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 15).Value, ",", ".") 'LARGURA
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 34).Value = _
                                Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 16).Value, ",", ".") 'ALTURA
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 35).Value = _
                                Sheets("Embalagens").Cells(LIN_EMBAL, 13).Value 'UNIDADE MEDIDA METROS
                        End If
                     End If
                 End If
                 '
                 If Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 2).Value = "X" Then 'VENDAS
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 40).Value = _
                    Sheets("Embalagens").Cells(LIN_EMBAL, 8).Value 'UMB ALTERNATIVA
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 41).Value = "=" 'SIMBOLO = IGUAL
                    '
                    CONVERSAO = Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 24).Value, ",", ".") 'CONTADOR CONVERSAO
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 42).Value = CONVERSAO
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 43).Value = _
                    Sheets("Embalagens").Cells(LIN_EMBAL, 8).Value 'UIDADE MEDIDA BASICA
                    '
                    If Sheets("Embalagens").Cells(LIN_EMBAL, 21).Value = "G" Then
                         Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 46).Value = _
                            Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 22).Value / 1000, 3), ",", ".") 'PESO BRUTO
                        Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 47).Value = _
                            Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 23).Value / 1000, 3), ",", ".") 'PESO LIQUIDO
                    Else
                        Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 46).Value = _
                             Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 22).Value, 3), ",", ".") 'PESO BRUTO
                        Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 47).Value = _
                            Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 23).Value, 3), ",", ".") 'PESO LIQUIDO
                    End If
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 48).Value = "KG"
                    'Sheets("Embalagens").Cells(LIN_EMBAL, 21).Value 'UNIDADE PESO
                    '
                    If Sheets("Embalagens").Cells(LIN_EMBAL, 13).Value = "MM" Then 'UNIDADE MEDIDA
                       'SE NA PLANILHA OS VALORES FORAM INFORMADOS EM "MM"
                       'SÃO CONVERTIDOS EM METROS
                       CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 14).Value / 1000), ",", ".")
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 49).Value = _
                            CONVERSAO 'COMPRIMENTO
                       CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 15).Value / 1000), ",", ".")
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 50).Value = _
                            CONVERSAO 'LARGURA
                       CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 16).Value / 1000), ",", ".")
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 51).Value = _
                            Replace(CONVERSAO, ",", ".") 'ALTURA
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 52).Value = "M"
                    Else
                        If Sheets("Embalagens").Cells(LIN_EMBAL, 13).Value = "CM" Then 'UNIDADE MEDIDA METROS
                           'SE NA PLANILHA OS VALORES FORAM INFORMADOS EM "CM"
                           'SÃO CONVERTIDOS EM METROS
                           CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 14).Value / 100), ",", ".")
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 49).Value = _
                                CONVERSAO 'COMPRIMENTO
                           CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 15).Value / 100), ",", ".")
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 50).Value = _
                                CONVERSAO 'LARGURA
                           CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 16).Value / 100), ",", ".")
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 51).Value = _
                                Replace(CONVERSAO, ",", ".") 'ALTURA
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 52).Value = "M"
                        Else
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 49).Value = _
                                Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 14).Value, ",", ".") 'COMPRIMENTO
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 50).Value = _
                                Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 15).Value, ",", ".") 'LARGURA
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 51).Value = _
                                Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 16).Value, ",", ".") 'ALTURA
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 52).Value = _
                                Sheets("Embalagens").Cells(LIN_EMBAL, 13).Value 'UNIDADE MEDIDA METROS
                        End If
                    End If
                 Else
                    If Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 1).Value <> "X" Then
                    'ALGO DE ERRADO - COLUNA NA TABELINHA
                    'EMBAL_EXPORT DA PLANILHA TRABALHO NAO DEVERIA TER AS COLUNAS
                    'COMPRA E VENDA EM BRANCO
                       MsgBox ("Ocorreu um erro na UNDADE. Informe o suporte Leroy Merlin.")
                    End If
                 End If
             '
            Case "MASTER"
                 If Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 1).Value = "X" Then 'COMPRAS
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 23).Value = _
                    Sheets("Embalagens").Cells(LIN_EMBAL, 42).Value 'UMB ALTERNATIVA
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 24).Value = "=" 'SIMBOLO = IGUAL
                    '
                    CONVERSAO = Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 55).Value, ",", ".") 'CONTADOR CONVERSAO
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 25).Value = CONVERSAO

                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 26).Value = _
                    Sheets("Embalagens").Cells(LIN_EMBAL, 8).Value 'UIDADE MEDIDA BASICA
                    '
                    If Sheets("Embalagens").Cells(LIN_EMBAL, 52).Value = "G" Then
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 29).Value = _
                       Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 53).Value / 1000, 3), ",", ".") 'PESO BRUTO
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 30).Value = _
                       Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 54).Value / 1000, 3), ",", ".") 'PESO LIQUIDO
                    Else
                        Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 29).Value = _
                            Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 53).Value, 3), ",", ".") 'PESO BRUTO
                        Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 30).Value = _
                            Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 54).Value, 3), ",", ".") 'PESO LIQUIDO
                    End If
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 31).Value = "KG"
                    'Sheets("Embalagens").Cells(LIN_EMBAL, 52).Value 'UNIDADE PESO
                    '
                    If Sheets("Embalagens").Cells(LIN_EMBAL, 48).Value = "MM" Then
                       'SE NA PLANILHA OS VALORES FORAM INFORMADOS EM "MM"
                       'SÃO CONVERTIDOS EM METROS
                       CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 49).Value / 1000), ",", ".")
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 32).Value = _
                            CONVERSAO 'COMPRIMENTO
                       '
                       CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 50).Value / 1000), ",", ".")
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 33).Value = _
                            CONVERSAO 'LARGURA
                       '
                       CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 51).Value / 1000), ",", ".")
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 34).Value = _
                            CONVERSAO 'ALTURA
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 35).Value = "M" 'UNIDADE MEDIDA METROS
                    Else
                        '
                        If Sheets("Embalagens").Cells(LIN_EMBAL, 48).Value = "CM" Then
                           'SE NA PLANILHA OS VALORES FORAM INFORMADOS EM "CM"
                           'SÃO CONVERTIDOS EM METROS
                           CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 49).Value / 100), ",", ".")
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 32).Value = _
                                CONVERSAO 'COMPRIMENTO
                           '
                           CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 50).Value / 100), ",", ".")
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 33).Value = _
                                CONVERSAO 'LARGURA
                           '
                           CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 51).Value / 100), ",", ".")
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 34).Value = _
                                CONVERSAO 'ALTURA
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 35).Value = "M" 'UNIDADE MEDIDA METROS
                        Else
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 32).Value = _
                            Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 49).Value, ",", ".") 'COMPRIMENTO
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 33).Value = _
                            Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 50).Value, ",", ".") 'LARGURA
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 34).Value = _
                            Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 51).Value, ",", ".") 'ALTURA
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 35).Value = _
                          Sheets("Embalagens").Cells(LIN_EMBAL, 48).Value 'UNIDADE MEDIDA METROS
                        End If
                    End If
                 End If
                 '
                 If Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 2).Value = "X" Then 'VENDAS
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 40).Value = _
                    Sheets("Embalagens").Cells(LIN_EMBAL, 42).Value 'UMB ALTERNATIVA
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 41).Value = "=" 'SIMBOLO = IGUAL
                    '
                    CONVERSAO = Sheets("Embalagens").Cells(LIN_EMBAL, 55).Value 'CONTADOR CONVERSAO
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 42).Value = CONVERSAO
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 43).Value = _
                    Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 8).Value, ",", ".") 'UIDADE MEDIDA BASICA
                    '
                    If Sheets("Embalagens").Cells(LIN_EMBAL, 52).Value = "G" Then
                        Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 46).Value = _
                            Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 53).Value / 1000, 3), ",", ".") 'PESO BRUTO
                        Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 47).Value = _
                            Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 54).Value / 1000, 3), ",", ".") 'PESO LIQUIDO
                    Else
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 46).Value = _
                       Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 53).Value, 3), ",", ".") 'PESO BRUTO
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 47).Value = _
                       Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 54).Value, 3), ",", ".") 'PESO LIQUIDO
                    End If
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 48).Value = "KG"
                    'Sheets("Embalagens").Cells(LIN_EMBAL, 52).Value 'UNIDADE PESO
                    '
                    If Sheets("Embalagens").Cells(LIN_EMBAL, 48).Value = "MM" Then
                       'SE NA PLANILHA OS VALORES FORAM INFORMADOS EM "MM"
                       'SÃO CONVERTIDOS EM METROS
                       CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 49).Value / 1000), ",", ".")
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 49).Value = _
                            CONVERSAO 'COMPRIMENTO
                       '
                       CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 50).Value / 1000), ",", ".")
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 50).Value = _
                            CONVERSAO 'LARGURA
                       '
                       CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 51).Value / 1000), ",", ".")
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 51).Value = _
                            CONVERSAO 'ALTURA
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 52).Value = "M" 'UNIDADE MEDIDA METROS
                    Else
                        '
                        If Sheets("Embalagens").Cells(LIN_EMBAL, 48).Value = "CM" Then
                           'SE NA PLANILHA OS VALORES FORAM INFORMADOS EM "CM"
                           'SÃO CONVERTIDOS EM METROS
                           CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 49).Value / 100), ",", ".")
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 49).Value = _
                                CONVERSAO 'COMPRIMENTO
                           '
                           CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 50).Value / 100), ",", ".")
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 50).Value = _
                                CONVERSAO 'LARGURA
                           '
                           CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 51).Value / 100), ",", ".")
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 51).Value = _
                                CONVERSAO 'ALTURA
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 52).Value = "M" 'UNIDADE MEDIDA METROS
                        Else
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 49).Value = _
                            Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 49).Value, ",", ".") 'COMPRIMENTO
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 50).Value = _
                            Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 50).Value, ",", ".") 'LARGURA
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 51).Value = _
                            Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 51).Value, ",", ".") 'ALTURA
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 52).Value = _
                          Sheets("Embalagens").Cells(LIN_EMBAL, 48).Value 'UNIDADE MEDIDA METROS
                        End If
                    End If
                 Else
                    If Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 1).Value <> "X" Then
                    'ALGO DE ERRADO - COLUNA NA TABELINHA
                    'EMBAL_EXPORT DA PLANILHA TRABALHO NAO DEVERIA TER AS COLUNAS
                    'COMPRA E VENDA EM BRANCO
                       MsgBox ("Ocorreu um erro na MASTER. Informe o suporte Leroy Merlin.")
                    End If
                 End If
            Case "PALLET"
                 If Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 1).Value = "X" Then 'COMPRAS
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 23).Value = "PAL"
                    'Sheets("Embalagens").Cells(LIN_EMBAL, 60).Value 'UMB ALTERNATIVA
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 24).Value = "=" 'SIMBOLO = IGUAL
                    '
                    CONVERSAO = Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 74).Value, ",", ".") 'CONTADOR CONVERSAO
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 25).Value = CONVERSAO
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 26).Value = _
                    Sheets("Embalagens").Cells(LIN_EMBAL, 8).Value 'UIDADE MEDIDA BASICA
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 29).Value = _
                        Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 68).Value, 3), ",", ".") 'PESO BRUTO
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 30).Value = _
                        Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 69).Value, 3), ",", ".") 'PESO LIQUIDO
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 31).Value = "KG"
                    'Sheets("Embalagens").Cells(LIN_EMBAL, 67).Value 'UNIDADE PESO
                    '
                    If Sheets("Embalagens").Cells(LIN_EMBAL, 63).Value = "MM" Then
                       'SE NA PLANILHA OS VALORES FORAM INFORMADOS EM "CM"
                       'SÃO CONVERTIDOS EM METROS
                      CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 64).Value / 1000), ",", ".")
                      Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 32).Value = _
                            CONVERSAO 'COMPRIMENTO
                      '
                      CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 65).Value / 1000), ",", ".")
                      Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 33).Value = _
                            CONVERSAO 'LARGURA
                      '
                      CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 66).Value / 1000), ",", ".")
                      Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 34).Value = _
                            CONVERSAO 'ALTURA
                      '
                      Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 35).Value = "M" 'UNIDADE MEDIDA METROS
                    Else
                        If Sheets("Embalagens").Cells(LIN_EMBAL, 63).Value = "CM" Then
                           'SE NA PLANILHA OS VALORES FORAM INFORMADOS EM "CM"
                           'SÃO CONVERTIDOS EM METROS
                          CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 64).Value / 100), ",", ".")
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 32).Value = _
                                CONVERSAO 'COMPRIMENTO
                          '
                          CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 65).Value / 100), ",", ".")
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 33).Value = _
                                CONVERSAO 'LARGURA
                          '
                          CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 66).Value / 100), ",", ".")
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 34).Value = _
                                CONVERSAO 'ALTURA
                          '
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 35).Value = "M" 'UNIDADE MEDIDA METROS
                        Else
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 32).Value = _
                            Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 64).Value, ",", ".") 'COMPRIMENTO
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 33).Value = _
                            Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 65).Value, ",", ".") 'LARGURA
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 34).Value = _
                            Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 66).Value, ",", ".") 'ALTURA
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 35).Value = _
                          Sheets("Embalagens").Cells(LIN_EMBAL, 63).Value 'UNIDADE MEDIDA METROS
                        End If
                    End If
                 End If
                 '
                 If Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 2).Value = "X" Then 'VENDAS
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 40).Value = "PAL"
                    'Sheets("Embalagens").Cells(LIN_EMBAL, 60).Value 'UMB ALTERNATIVA
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 41).Value = "=" 'SIMBOLO = IGUAL
                    '
                    CONVERSAO = Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 74).Value, ",", ".") 'CONTADOR CONVERSAO
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 42).Value = CONVERSAO
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 43).Value = _
                    Sheets("Embalagens").Cells(LIN_EMBAL, 8).Value 'UIDADE MEDIDA BASICA
                    '
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 46).Value = _
                        Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 68).Value, 3), ",", ".") 'PESO BRUTO
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 47).Value = _
                        Replace(Round(Sheets("Embalagens").Cells(LIN_EMBAL, 69).Value, 3), ",", ".") 'PESO LIQUIDO
                    Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 48).Value = "KG"
                    'Sheets("Embalagens").Cells(LIN_EMBAL, 67).Value 'UNIDADE PESO
                    '
                    If Sheets("Embalagens").Cells(LIN_EMBAL, 63).Value = "MM" Then
                       'SE NA PLANILHA OS VALORES FORAM INFORMADOS EM "MM"
                       'SÃO CONVERTIDOS EM METROS
                       CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 64).Value / 1000), ",", ".")
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 49).Value = _
                            CONVERSAO 'COMPRIMENTO
                       '
                       CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 65).Value / 1000), ",", ".")
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 50).Value = _
                            CONVERSAO 'LARGURA
                       '
                       CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 66).Value / 1000), ",", ".")
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 51).Value = _
                            CONVERSAO 'ALTURA
                       Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 52).Value = "M" 'UNIDADE MEDIDA METROS
                    Else
                        '
                        If Sheets("Embalagens").Cells(LIN_EMBAL, 63).Value = "CM" Then
                           'SE NA PLANILHA OS VALORES FORAM INFORMADOS EM "CM"
                           'SÃO CONVERTIDOS EM METROS
                           CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 64).Value / 100), ",", ".")
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 49).Value = _
                                CONVERSAO 'COMPRIMENTO
                           '
                           CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 65).Value / 100), ",", ".")
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 50).Value = _
                                CONVERSAO 'LARGURA
                           '
                           CONVERSAO = Replace(CStr(Sheets("Embalagens").Cells(LIN_EMBAL, 66).Value / 100), ",", ".")
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 51).Value = _
                                CONVERSAO 'ALTURA
                           Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 52).Value = "M" 'UNIDADE MEDIDA METROS
                        Else
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 49).Value = _
                            Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 64).Value, ",", ".") 'COMPRIMENTO
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 50).Value = _
                            Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 65).Value, ",", ".") 'LARGURA
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 51).Value = _
                            Replace(Sheets("Embalagens").Cells(LIN_EMBAL, 66).Value, ",", ".") 'ALTURA
                          Sheets("Unidades Medida").Cells(LIN_Unidades_Medida, 52).Value = _
                            Sheets("Embalagens").Cells(LIN_EMBAL, 63).Value 'UNIDADE MEDIDA METROS
                        End If
                    End If
                 Else
                     If Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB + 1).Value <> "X" Then
                     'ALGO DE ERRADO - COLUNA NA TABELINHA
                     'EMBAL_EXPORT DA PLANILHA TRABALHO NAO DEVERIA TER AS COLUNAS
                     'COMPRA E VENDA EM BRANCO
                        MsgBox ("Ocorreu um erro em PALLET. Informe o suporte Leroy Merlin.")
                     End If
                 End If
            Case Else
                     'ALGO DE ERRADO - NA TABELINHA
                     'EMBAL_EXPORT DA PLANILHA TRABALHO DEVERIA CONTER NA COLUNA EMBALAGEM:
                     'UNIDADE MASTER OU PALLET
                        MsgBox ("Ocorreu um erro inesperado. Envie o problema ao suporte Leroy Merlin")
            End Select
            '
            'PROXIMO REGISTRO TABELA EMBAL_EXPORT PLANILHA TRABALHO
            LIN_TAB = LIN_TAB + 1
            Sheets("TRABALHO").Cells(LIN_TAB, COL_TAB).Select
       Loop

   End If
'
    LIN_EMBAL = LIN_EMBAL + 1
    LIN_Unidades_Medida = LIN_Unidades_Medida + 1
    CEL_COMPR = 0
    CEL_VEND = 0
    Application.Goto Reference:="EMBAL_EXPORT"
    Selection.ClearContents
    Application.Goto Reference:="INICIO_EMBAL_COMPR_VEND"
    LIN_TAB = ActiveCell.Row
    '
    ThisWorkbook.Sheets("TRABALHO").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    '
    ThisWorkbook.Sheets("Unidades Medida").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    Sheets("Embalagens").Select
    '
Next i

On Error GoTo 0

End Sub

Sub CARREGA_FLUXO()
    Dim lin As Integer
    Dim col As Integer
    Dim LIN_PERFIL As Integer
    Dim COL_PERFIL As Integer
    '
    Dim CENTRO As String
    Dim CODIGO As String
    Dim REGIAO As String
    '
    Application.ScreenUpdating = False
    '
    Sheets("Escolha Perfil Distribuicao").Select
    Application.Goto Reference:="INICIO_CENTRO"
    LIN_PERFIL = ActiveCell.Row
    COL_PERFIL = ActiveCell.Column
    '
    Sheets("TABELAS").Visible = True
    Sheets("TABELAS").Select
    '
    Application.Goto Reference:="PRACAS_TBL"
    lin = ActiveCell.Row
    col = ActiveCell.Column
    '
    CENTRO = Cells(lin, col).Value
    REGIAO = Cells(lin, col + 1).Value
    CODIGO = Cells(lin, col + 2).Value
    ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Unprotect Key
    '
    Do While ActiveCell.Value <> ""
       '
       Sheets("Escolha Perfil Distribuicao").Select
       '
       Cells(LIN_PERFIL, COL_PERFIL).Value = CODIGO
       Cells(LIN_PERFIL, COL_PERFIL + 1).Value = CENTRO
       Cells(LIN_PERFIL, COL_PERFIL + 2).Value = REGIAO
       lin = lin + 1
       LIN_PERFIL = LIN_PERFIL + 1
       Sheets("TABELAS").Select
       Cells(lin, col).Activate
       CENTRO = Cells(lin, col).Value
       REGIAO = Cells(lin, col + 1).Value
       CODIGO = Cells(lin, col + 2).Value
       '
    Loop
    '
    ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
End Sub

'
'NA ABA Escolha Perfil Distribuicao A PRIMEIRA CAIXA DE SELECAO
'É TODAS (ALL).
'ESTA ROTINA MARCARÁ OU DESMARCARÁ TODAS
'
Sub ESCOLHIDA_TODAS()
    Range("E5").Select
    Range(Selection, "E125").Select
    Selection.FillDown
End Sub


Sub APAGA_FICAIS()
    '
    Application.ScreenUpdating = False
    '
    ActiveWorkbook.Sheets("Dados Fiscais").Select
    ThisWorkbook.Sheets("Dados Fiscais").Unprotect Key
    ActiveSheet.Range("B3:J10000").Select
    Selection.ClearContents
    '
    ActiveSheet.Range("M3").Select
    Selection.ClearContents
    ActiveSheet.Range("N3").Select
    Selection.ClearContents
    ActiveSheet.Range("P3").Select
    Selection.ClearContents
    ActiveSheet.Range("Q3").Select
    Selection.ClearContents
    ActiveSheet.Range("S3").Select
    Selection.ClearContents
    ActiveSheet.Range("T3").Select
    Selection.ClearContents
    ActiveSheet.Range("U3").Select
    Selection.ClearContents
    ActiveSheet.Range("V3").Select
    Selection.ClearContents
    ActiveSheet.Range("W3").Select
    Selection.ClearContents
    ActiveSheet.Range("X3").Select
    Selection.ClearContents
    ActiveSheet.Range("Y3").Select
    Selection.ClearContents
    ActiveSheet.Range("Z3").Select
    Selection.ClearContents
    ActiveSheet.Range("AA3").Select
    Selection.ClearContents
    '
    ActiveSheet.Range("AC3").Select
    Selection.ClearContents
    ActiveSheet.Range("AD3").Select
    Selection.ClearContents
    ActiveSheet.Range("AE3").Select
    Selection.ClearContents
    '
    ActiveSheet.Range("AG3").Select
    Selection.ClearContents
    ActiveSheet.Range("AJ3").Select
    Selection.ClearContents
    '
    ActiveSheet.Range("J3:AJ5000").Select
    Selection.FillDown
    '
    '
    ThisWorkbook.Sheets("Dados Fiscais").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    '
    ThisWorkbook.Sheets("Dados Cadastrais").Activate
    '
End Sub
Function ADDITEM_COMBOBOX_FORNEC(ByRef lin As Integer, _
                                 ByRef LIN_TBL As Integer, _
                                 ByRef COL_TBL As Integer, _
                                 ByRef DESCR_CENTRO As String, _
                                 ByRef CODFORNECSAP As String, _
                                 ByRef CNPJ_FORNEC As String, _
                                 ByRef RAZAO_SOCIAL As String, _
                                 ByRef CENTRO As String, _
                                 ByRef PERFIL As String)
         Dim COMBOBOX As Object
         Dim CENTRO_TBL As String
         Dim FORNEC_VIRTUA As String
         '
         Dim ORG
         '
         Dim LIN_VIRTUA As Integer
         Dim COL_VIRTUA As Integer
         '
         Dim CODFORNECBUSCA As String
         '
         Application.ScreenUpdating = False
         '
         If Sheets("TABELAS").Visible = False Then
            Sheets("TABELAS").Visible = True
         End If
        '
        Application.Goto Reference:="COD_FORNEC_SAP"
        On Error Resume Next
        ORG = Selection.Find(What:=CODFORNECSAP, _
                            After:=ActiveCell, LookIn:=xlValues, _
                            LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                            MatchCase:=False, SearchFormat:=False).Activate
        On Error GoTo 0
        If Not ORG Then
           MsgBox ("Não encontrado Fornecedor com esse CNPJ nas tabelas internas da pasta. Verifique por favor.")
           ADDITEM_COMBOBOX_FORNEC = False
           Exit Function
        End If
        '
        Application.EnableEvents = False
        ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Cells(LIN_TBL, COL_TBL).Value = DESCR_CENTRO
        ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Cells(LIN_TBL, COL_TBL + 1).Value = CODFORNECSAP
        'COD.FORNEC SAP  CNPJ    RAZAO SOCIAL
        ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Cells(LIN_TBL, COL_TBL + 2).Value = "#" & _
                                                                                               Trim(Right("0000" & CENTRO, 4)) & _
                                                                                               Trim(CODFORNECSAP)
        ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Cells(LIN_TBL, COL_TBL + 3).Value = CNPJ_FORNEC
        ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Cells(LIN_TBL, COL_TBL + 4).Value = RAZAO_SOCIAL
        Application.EnableEvents = True
        ''''''''''''''''''''''''''''''''''
        'LINHAS DA PLANILHA
        'NÃO ENCONTREI COMO AUTOMATIZAR
        'A DEFINIÇÃO DO NOME DO COMBOBOX NA LINHA
        ''''''''''''''''''''''''''''''''''
        '
        ''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case LIN_TBL
        Case 5
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").fornec5
        Case 6
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC6
        Case 7
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC7
        Case 8
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC8
        Case 9
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC9
        Case 10
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC10
        Case 11
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC11
        Case 12
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC12
        Case 13
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC13
        Case 14
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC14
        Case 15
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC15
        Case 16
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC16
        Case 17
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC17
        Case 18
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC18
        Case 19
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC19
        Case 20
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC20
        Case 21
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC21
        Case 22
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC22
        Case 23
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC23
        Case 24
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC24
        Case 25
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC25
        Case 26
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC26
        Case 27
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC27
        Case 28
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC28
        Case 29
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC29
        Case 30
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC30
        Case 31
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC31
        Case 32
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC32
        Case 33
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC33
        Case 34
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC34
        Case 35
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC35
        Case 36
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC36
        Case 37
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC37
        Case 38
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC38
        Case 39
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC39
        Case 40
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC40
        Case 41
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC41
        Case 42
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC42
        Case 43
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC43
        Case 44
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC44
        Case 45
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC45
        Case 46
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC46
        Case 47
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC47
        Case 48
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC48
        Case 49
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC49
        Case 50
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC50
        Case 51
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC51
        Case 52
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC52
        Case 53
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC53
        Case 54
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC54
        Case 55
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC55
        Case 56
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC56
        Case 57
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC57
        Case 58
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC58
        Case 59
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC59
        Case 60
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC60
        Case 61
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC61
        Case 62
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC62
        Case 63
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC63
        Case 64
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC64
        Case 65
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC65
        Case 66
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC66
        Case 67
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC67
        Case 68
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC68
        Case 69
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC69
        Case 70
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC70
        Case 71
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC71
        Case 72
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC72
        Case 73
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC73
        Case 74
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC74
        Case 75
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC75
        Case 76
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC76
        Case 77
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC77
        Case 78
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC78
        Case 79
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC79
        Case 80
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC80
        Case 81
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC81
        Case 82
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC82
        Case 83
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC83
        Case 84
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC84
        Case 85
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC85
        Case 86
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC86
        Case 87
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC87
        Case 88
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC88
        Case 89
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC89
        Case 90
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC90
        Case 91
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC91
        Case 92
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC92
        Case 93
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC93
        Case 94
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC94
        Case 95
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC95
        Case 96
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC96
        Case 97
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC97
        Case 98
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC98
        Case 99
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC99
        Case 100
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC100
        Case 101
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC101
        Case 102
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC102
        Case 103
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC103
        Case 104
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC104
        Case 105
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC105
        Case 106
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC106
        Case 107
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC107
        Case 108
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC108
        Case 109
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC109
        Case 110
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC110
        Case 111
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC111
        Case 112
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC112
        Case 113
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC113
        Case 114
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC114
        Case 115
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC115
        Case 116
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC116
        Case 117
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC117
        Case 118
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC118
        Case 119
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC119
        Case 120
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC120
        Case 121
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC121
        Case 122
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC122
        Case 123
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC123
        Case 124
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC124
        Case 125
          Set COMBOBOX = ThisWorkbook.Sheets("Escolha Perfil Distribuicao").FORNEC125
        End Select
        '
        COMBOBOX.Clear
        '
        On Error Resume Next
        '
        LIN_VIRTUA = ActiveCell.Row
        COL_VIRTUA = ActiveCell.Column
        '
        With COMBOBOX
           Do While (ActiveCell.Value <> "" And ActiveCell.Value <> 0) And _
                     ActiveCell.Value = CODFORNECSAP
              'CENTRO
               If CENTRO = Sheets("TABELAS").Cells(LIN_VIRTUA, COL_VIRTUA + 3).Value Or _
                  (Sheets("TABELAS").Cells(LIN_VIRTUA, COL_VIRTUA + 3).Value = "" Or _
                   Sheets("TABELAS").Cells(LIN_VIRTUA, COL_VIRTUA + 3).Value = 0) Then
                   'COD.SAP/CENTRO/PERFIL/SUBSORTIMENTO
                   If (Sheets("TABELAS").Cells(LIN_VIRTUA, COL_VIRTUA + 3).Value = "" Or _
                       Sheets("TABELAS").Cells(LIN_VIRTUA, COL_VIRTUA + 3).Value = 0) Then
                       CENTRO_TBL = "____"
                   Else
                       CENTRO_TBL = Right("0000" & CENTRO, 4)
                   End If
                   '
                   If PERFIL = "Z01" Then
                      If Sheets("TABELAS").Cells(LIN_VIRTUA, COL_VIRTUA + 4).Value = "Y2" Then
                         FORNEC_VIRTUA = Mid(Val(Cells(LIN_VIRTUA, COL_VIRTUA + 5)) + 10000000000#, 2, 10) & _
                                        "/" & CENTRO_TBL & "/" & Sheets("TABELAS").Cells(LIN_VIRTUA, COL_VIRTUA + 4).Value & _
                                        "/" & Sheets("TABELAS").Cells(LIN_VIRTUA, COL_VIRTUA + 2).Value
                      End If
                   Else
                     If Sheets("TABELAS").Cells(LIN_VIRTUA, COL_VIRTUA + 4).Value <> "Y2" Then
                        FORNEC_VIRTUA = Mid(Val(Cells(LIN_VIRTUA, COL_VIRTUA + 5)) + 10000000000#, 2, 10) & _
                                        "/" & CENTRO_TBL & "/" & Sheets("TABELAS").Cells(LIN_VIRTUA, COL_VIRTUA + 4).Value & _
                                        "/" & Sheets("TABELAS").Cells(LIN_VIRTUA, COL_VIRTUA + 2).Value
                     End If
                   End If
                   .AddItem FORNEC_VIRTUA
               End If
               '
               LIN_VIRTUA = LIN_VIRTUA + 1
               Sheets("TABELAS").Cells(LIN_VIRTUA, COL_VIRTUA).Activate
           Loop
        End With
        ADDITEM_COMBOBOX_FORNEC = True
        '
        Application.EnableEvents = False
        Sheets("Escolha Perfil Distribuicao").Select
        Application.EnableEvents = True
        '
        '''Application.ScreenUpdating = True
End Function

Public Sub ADDITEM_FORNEC_PD()
    Dim CNPJ_FORNEC As String
    Dim CNPJ_FORNEC_INTEIRO As String
    Dim COD_FORNEC_SAP As String
    Dim RAZAO As String
    Dim ORG
    Dim col As Integer
    Dim lin As Integer
    Dim COL_TBL As Integer
    Dim LIN_TBL As Integer
    '
    Application.ScreenUpdating = False
    '
    ThisWorkbook.Sheets("PARAMETROS").Select
    Application.Goto Reference:="CNPJ_RAIZ"
    CNPJ_FORNEC = ActiveCell.Value
    '
    If Sheets("TABELAS").Visible = False Then
       Sheets("TABELAS").Visible = True
    End If
    '
    ThisWorkbook.Sheets("TABELAS").Unprotect Key
    '
    Application.Goto Reference:="CNPJ_BUSCA"
    ORG = Selection.Find(What:=CNPJ_FORNEC, _
                        After:=ActiveCell, LookIn:=xlValues, _
                        LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                        MatchCase:=False, SearchFormat:=False).Activate

    If Not ORG Then
       MsgBox ("Não encontrado Fornecedor com esse CNPJ nas tabelas internas da pasta. Verifique por favor.")
    Else
        ActiveCell.Select
        lin = ActiveCell.Row
        col = ActiveCell.Column
        CNPJ_FORNEC_INTEIRO = ActiveCell.Value
        RAZAO = ThisWorkbook.Sheets("TABELAS").Cells(lin, col - 4).Value
        COD_FORNEC_SAP = ThisWorkbook.Sheets("TABELAS").Cells(lin, col - 3).Value
        '
        
        '
        Application.Goto Reference:="CNPJ_ESCOLHIDO_PROCV"
        '
        LIN_TBL = ActiveCell.Row
        COL_TBL = ActiveCell.Column
        '
        Selection.Clear
        '
        Do While True
            '
            ThisWorkbook.Sheets("TABELAS").Cells(LIN_TBL, COL_TBL).Value = _
            RAZAO '& " - " & Trim(CNPJ_FORNEC_INTEIRO)
            ThisWorkbook.Sheets("TABELAS").Cells(LIN_TBL, COL_TBL + 1).Value = _
            Trim(CNPJ_FORNEC_INTEIRO)
            ThisWorkbook.Sheets("TABELAS").Cells(LIN_TBL, COL_TBL + 2).Value = _
            Trim(COD_FORNEC_SAP)
            '
            lin = lin + 1
            
            ThisWorkbook.Sheets("TABELAS").Cells(lin, col).Activate
            '
            If InStr(1, ActiveCell.Value, CNPJ_FORNEC) = 0 Then
               Exit Do
            End If
            LIN_TBL = LIN_TBL + 1
            '
            Cells(lin, col).Activate
            '
            CNPJ_FORNEC_INTEIRO = ActiveCell.Value
            RAZAO = ThisWorkbook.Sheets("TABELAS").Cells(lin, col - 4).Value
            COD_FORNEC_SAP = ThisWorkbook.Sheets("TABELAS").Cells(lin, col - 3).Value
        Loop
    End If
    ThisWorkbook.Sheets("PARAMETROS").Select
    Application.Goto Reference:="CNPJ_RAIZ"
    '
    ThisWorkbook.Sheets("TABELAS").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    '
End Sub

Function ADDITEM_FORNEC_CENTRO(ByRef COD_FORNEC_SAP As String, _
                               ByRef CNPJ_FORNEC As String, _
                               ByRef RAZAO_SOCIAL As String, _
                               ByRef PERFIL As String)
    Dim CNPJ_FORNEC As String
    Dim COD_FORNEC_SAP As String
    Dim RAZAO As String
    Dim ORG
    Dim col As Integer
    Dim lin As Integer
    Dim COL_TBL As Integer
    Dim LIN_TBL As Integer
    '
    Application.ScreenUpdating = False
    '
    ThisWorkbook.Sheets("PARAMETROS").Select
    Application.Goto Reference:="CNPJ_RAIZ"
    CNPJ_FORNEC = ActiveCell.Value
    '
    ThisWorkbook.Sheets("TABELAS").Select
    Application.Goto Reference:="CNPJ_BUSCA"
    ORG = Selection.Find(What:=CNPJ_FORNEC, _
                        After:=ActiveCell, LookIn:=xlValues, _
                        LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                        MatchCase:=False, SearchFormat:=False).Activate
    If Not ORG Then
       MsgBox ("Não encontrado Fornecedor com esse CNPJ nas tabelas internas da pasta. Verifique por favor.")
    Else
        ActiveCell.Select
        lin = ActiveCell.Row
        col = ActiveCell.Column
        COD_FORNEC_SAP = ThisWorkbook.Sheets("TABELAS").Cells(lin, col - 3).Value
        RAZAO = ThisWorkbook.Sheets("TABELAS").Cells(lin, col - 4).Value
        '
        ThisWorkbook.Sheets("TABELAS").Select
        Application.Goto Reference:="COD_FORNEC_SAP"
        ORG = Selection.Find(What:=COD_FORNEC_SAP, _
                            After:=ActiveCell, LookIn:=xlValues, _
                            LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                            MatchCase:=False, SearchFormat:=False).Activate
        If ORG Then
            lin = ActiveCell.Row
            col = ActiveCell.Column
            '
            ThisWorkbook.Sheets("PARAMETROS").Select
            Application.Goto Reference:="FORNECEDOR_VIRTUAL"
            LIN_TBL = ActiveCell.Row
            COL_TBL = ActiveCell.Column
            '
            ThisWorkbook.Sheets("TABELAS").Select
            '
            Do While ActiveCell.Value = COD_FORNEC_SAP
                '
                ORG_COMPRAS = ThisWorkbook.Sheets("TABELAS").Cells(lin, col + 1).Value
                SUBSORTIMENTO = ThisWorkbook.Sheets("TABELAS").Cells(lin, col + 2).Value
                CENTRO = ThisWorkbook.Sheets("TABELAS").Cells(lin, col + 3).Value
                FLUXO = ThisWorkbook.Sheets("TABELAS").Cells(lin, col + 4).Value
                COD_FORNEC_VIRTUAL = ThisWorkbook.Sheets("TABELAS").Cells(lin, col + 5).Value
                '
                ThisWorkbook.Sheets("PARAMETROS").Select
                Cells(LIN_TBL, COL_TBL) = COD_FORNEC_SAP
                Cells(LIN_TBL, COL_TBL + 1) = ORG_COMPRAS
                Cells(LIN_TBL, COL_TBL + 2) = SUBSORTIMENTO
                Cells(LIN_TBL, COL_TBL + 3) = CENTRO
                Cells(LIN_TBL, COL_TBL + 4) = FLUXO
                Cells(LIN_TBL, COL_TBL + 5) = COD_FORNEC_VIRTUAL
                '
                ThisWorkbook.Sheets("TABELAS").Select
                '
                LIN_TBL = LIN_TBL + 1
                lin = lin + 1
                '
                ThisWorkbook.Sheets("TABELAS").Cells(lin, col).Activate
                '
            Loop
        Else
           MsgBox ("Não encontrado Fornecedor Virtual para esse CNPJ. Verifique por favor.")
        End If
    End If
    '
    '''Application.ScreenUpdating = True
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VALIDACAO_PERIGOSOS()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VALIDACAO_PERIGOSOS()
'
' VALIDACAO CONTEUDO PLANILHAS
'
Dim ERROS As Boolean
'
Dim COL_CELL As Integer
Dim LIN_CELL As Integer
Dim CORRE_COL As Integer
Dim CELL_ADRESS As String
Dim QTDE_LINHAS_LIBERADAS As Integer
'
Dim LINHAS_ERRO As Integer
'
Dim i As Integer
'
Dim SIMNAO As String
Dim resultado
'
'PLANILHA DADOS GERAIS 2
'
Application.ScreenUpdating = False
'
'SE NAO ESCOLHEU "VALIDAR" SAI FORA
'
Sheets("PARAMETROS").Select
Application.Goto Reference:="QTDE_LINHAS_LIBERADAS"
QTDE_LINHAS_LIBERADAS = ActiveCell.Value
'
Sheets("ERROS em Dados Regulatorios").Select
Cells.Select
Selection.Delete Shift:=xlUp
'
Range("A1").Select
ActiveCell.FormulaR1C1 = "MENSAGEM"
Range("B1").Select
ActiveCell.FormulaR1C1 = "PLANILHA"
Range("C1").Select
ActiveCell.FormulaR1C1 = "QUADRO"
Range("D1").Select
ActiveCell.FormulaR1C1 = "INTERVALO"
Range("A2").Select
LINHAS_ERRO = 2 'SEGUNDA LINHA
'
Application.Goto Reference:="VALIDA_SO_ESTA_REGULA"
If ActiveCell.FormulaR1C1 = "TRUE" Then
    Sheets("Dados Regulatorios").Select
    Application.Goto Reference:="INICIO_PERIGOSOS"
    '
    COL_CELL = ActiveCell.Column
    LIN_CELL = ActiveCell.Row
    CORRE_COL = ActiveCell.Column
    '
    Sheets("Dados Regulatorios").PERIGOSOS.BackColor = &HC000& 'VERDE
    Sheets("Inicio").REGULATORIOS.BackColor = &HC000& 'VERDE
    Sheets("Dados Regulatorios").PERIGOSOS.ForeColor = &H8000000E  'BRANCO
    Sheets("Inicio").REGULATORIOS.ForeColor = &H8000000E  'BRANCO
    '
    ERROS = False
    '
    For i = 1 To QTDE_LINHAS_LIBERADAS
        ''''''''''''''''''''''''''''''''''''''''''''''
        'SE NÃO TIVER CODIGO DO FORNRCEDOR NA LINHA
        If Cells(LIN_CELL, 2).Value = 0 Then
           Exit For
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''
        '
    
            '&H00FF8080& BOTAO AZUL
            '&H0000FFFF& AMARELO
            '&H0000C000& VERDE
            '&H000000FF& VERMELHO
            
         SIMNAO = ActiveCell.Value
         If SIMNAO = "SIM" Or SIMNAO = "NÃO" Then
            CORRE_COL = CORRE_COL + 1
            Cells(LIN_CELL, CORRE_COL).Activate
            '
            If ActiveCell.Value = "" And SIMNAO = "SIM" Then 'CLASSE
               CELL_ADRESS = ActiveCell.Address
               'MsgBox "O campo COD.CLASSE deve ser preenchido, escolha em CLASSE, por que o campo MATERIAL PERIGOSO? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - PERIGOSOS"
               ERROS = True
               Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo COD.CLASSE deve ser preenchido, escolha em CLASSE, por que o campo MATERIAL PERIGOSO? = 'SIM'."
               Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
               Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "Material Perigoso"
               Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Dados Regulatorios").PERIGOSOS.BackColor = &HFF&     'VERMELHO
               Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
               LINHAS_ERRO = LINHAS_ERRO + 1
            Else
               If ActiveCell.Value <> "" And _
                  ActiveCell.Value <> 0 And SIMNAO = "NÃO" Then 'CLASSE
'                  RESULTADO = MsgBox("A Pergunta Inicial 'PRODUTO CONTROLADO' desta linha, foi colocada como 'NÃO'. Portanto " & _
'                                     "O campo 'CLASSE' não poderia estar preenchido. Posso LIMPA-LO?", vbYesNo, "VALIDACAO PERIGOSOS.")
                  If resultado = vbYes Then
                     Cells(ActiveCell.Row, ActiveCell.Column + 1).ClearContents
                  Else
                      ERROS = True
                      CELL_ADRESS = ActiveCell.Address
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo CLASSE  NÃO deve ser preenchido por que o campo PRODUTO CONTROLADO?(PERIGOSOS) = 'NÃO'."
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "PERIGOSOS"
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                      Sheets("Dados Regulatorios").ANVISA.BackColor = &HFF& 'VERMELHO
                      Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                      LINHAS_ERRO = LINHAS_ERRO + 1
                  End If
               End If
            End If
            '
            CORRE_COL = CORRE_COL + 2
            Cells(LIN_CELL, CORRE_COL).Activate
            '
            If ActiveCell.Value = "" And SIMNAO = "SIM" Then 'SUBCLASSE
               CELL_ADRESS = ActiveCell.Address
               'MsgBox "O campo COD.SUB CLASSE deve ser preenchido, escolha em SUB CLASSE, por que o campo MATERIAL PERIGOSO? = 'SIM'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - PERIGOSOS"
               ERROS = True
               Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo COD.SUB CLASSE deve ser preenchido, escolha em SUB CLASSE, por que o campo MATERIAL PERIGOSO? = 'SIM'."
               Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
               Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "Material Perigoso"
               Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
               Sheets("Dados Regulatorios").PERIGOSOS.BackColor = &HFF&   'VERMELHO
               Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
               LINHAS_ERRO = LINHAS_ERRO + 1
            Else
               If ActiveCell.Value <> "" And _
                  ActiveCell.Value <> 0 And SIMNAO = "NÃO" Then 'CLASSE
'                  RESULTADO = MsgBox("A Pergunta Inicial 'PRODUTO CONTROLADO' desta linha, foi colocada como 'NÃO'. Portanto " & _
'                                     "O campo 'SUBCLASSE' não poderia estar preenchido. Posso LIMPA-LO?", vbYesNo, "VALIDACAO PERIGOSOS.")
                  If resultado = vbYes Then
                     Cells(ActiveCell.Row, ActiveCell.Column + 1).ClearContents
                  Else
                      ERROS = True
                      CELL_ADRESS = ActiveCell.Address
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo SUBCLASSE  NÃO deve ser preenchido por que o campo PRODUTO CONTROLADO?(PERIGOSOS) = 'NÃO'."
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "PERIGOSOS"
                      Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
                      Sheets("Dados Regulatorios").ANVISA.BackColor = &HFF& 'VERMELHO
                      Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
                      LINHAS_ERRO = LINHAS_ERRO + 1
                  End If
               End If
           End If
           '
        Else
             CELL_ADRESS = ActiveCell.Address
             'MsgBox "O campo MATERIAL PERIGOSO? deve ser preenchido, 'SIM' ou 'NÃO'. Veja em " & CELL_ADRESS, vbOKOnly, "Dados Regulatorios - PERIGOSOS"
             ERROS = True
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 1).Value = "O campo MATERIAL PERIGOSO? deve ser preenchido, 'SIM' ou 'NÃO'."
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 2).Value = "Dados Regulatorios"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 3).Value = "Material Perigoso"
             Sheets("ERROS em Dados Regulatorios").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Inicio").REGULATORIOS.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Regulatorios").PERIGOSOS.BackColor = &HFF&
             LINHAS_ERRO = LINHAS_ERRO + 1
       End If
       '
       LIN_CELL = LIN_CELL + 1
       CORRE_COL = COL_CELL
       Cells(LIN_CELL, CORRE_COL).Activate
    Next i
    '
    Sheets("ERROS em Dados Regulatorios").Select
    Cells.Select
    Selection.ColumnWidth = 8.29
    Cells.EntireColumn.AutoFit
'    If ERROS Then
'       MsgBox "ERROS DE VALIDACAO, verifique na pagina 'ERROS em Dados Regulatorios'.", vbOKOnly, "Dados Regulatorios - PERIGOSOS"
'       Sheets("ERROS em Dados Regulatorios").Select
'       Cells.Select
'       Selection.ColumnWidth = 8.29
'       Cells.EntireColumn.AutoFit
'    Else
'      Sheets("Dados Regulatorios").Select
'      MsgBox "SEM ERROS DE VALIDACAO.", vbOKOnly, "Dados Regulatorios - PERIGOSOS"
'    End If
End If
Application.Goto Reference:="VALIDA_SO_ESTA_REGULA"
Sheets("Dados Regulatorios").Select
'
'''Application.ScreenUpdating = True
End Sub


'COD.CEST - O - 15
'CEST - P - 16
'PRECO UNITARIO ESTOQUE - Q - 17
'TOTAL COMPRAS = R - 18
'FABRICADO OU REVENDIDO - S - 19
'SIMPLES NACIONAL? - T - 20
'DIFERIMENTO ICMS? - U - 21
'PAUTA DE PRECO? - V - 22
'ISENTO OU IMUNE ICMS? - W - 23
'ISENTO PIS E COFINS? - X - 24
'%REDUCAO ICMS - Y - 25
'ALIQUOTA ICMS - Z - 26
'%ALIQUOTA IPI - AA - 27
'COD.SUBST.TRIBUT. - AB - 28
'SUBSTITUICAO TRIBUTARIA - AC - 29
'CONTRIBUINTE SUBSTITUTO - AD - 30
'% MVA OU IVA - AE- 31
'COD.ORIG.MATER. - AF - 32
'ORIGEM MATERIAL - AG - 33
'ESTADO DE ORIGEM - AH - 34
'COD.IVA DESCR.IVA - AI - 35
'DESCR.IVA - AJ - 36

Sub SUGERE_IVA(lin As Integer)
    Dim CODIVA As String
    Dim DESCR_CODIVA As String
    '
Sheets("Dados Fiscais").Select
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Regra 1 - Simples Nacional - NÃO e Fabricado ou Revendido
'  MOVE  eg_centro_aux-aliquotaipi TO vl_p
'  IF  eg_centro_aux-prodicms4         EQ '0'
'      AND eg_centro_aux-prodisentopis EQ '1'
'      AND vl_p                        GT 0.
'    eg_centro_aux-codiva = 'ZQ'.
''''''''''''''''''''''''''''''''''''''''
'ISENTO OU IMUNE ICMS? - W - 23
'ISENTO PIS E COFINS? - X - 24
'%ALIQUOTA IPI - AA - 27
''''''''''''''''''''''''''''''''''''''''
If Cells(lin, 23) = "SIM" And _
   Cells(lin, 24) = "NÃO" And _
   Cells(lin, 27) > 0 Then
   '
   'Cells(LIN, 35) = "ZQ" 'ZQ: Isento ICMS + PIS / COFINS + Al.IPI > 0
   CODIVA = "ZQ"
Else
  ''''''''''''''''''''''''''''''''''''''
  'ELSEIF eg_centro_aux-prodisentopis EQ '1'
  '       AND eg_centro_aux-produtost EQ '3'.
  '    eg_centro_aux-codiva = 'ZS'.
  ''''''''''''''''''''''''''''''''''''''
  'ISENTO PIS E COFINS? - X - 24
  'COD.SUBST.TRIBUT. - AB - 28
  ''''''''''''''''''''''''''''''''''''''
  If Cells(lin, 24) = "NÃO" And _
     Cells(lin, 28) = 3 Then
     '
     'Cells(LIN, 35) = "ZS" 'ZS: ST Recolhida Anteriormente + ICMS + IPI + PIS/COFINS
     CODIVA = "ZS"
  Else
    '''''''''''''''''''''''''''''''''''''''
    '  ELSEIF eg_centro_aux-prodicms4         EQ '0'
    '         AND eg_centro_aux-prodisentopis EQ '0'
    '         AND eg_centro_aux-aliquota      EQ '0.00'
    '         AND eg_centro_aux-aliquotaicms  EQ '0.00'
    '         AND eg_centro_aux-aliquotaipi   EQ '0.00'
    '         AND eg_centro_aux-produtost     EQ '0'
    '         AND eg_centro_aux-mva           EQ '0'.
    '    eg_centro_aux-codiva = 'ZX'.
    '''''''''''''''''''''''''''''''''''''''
    'ISENTO OU IMUNE ICMS? - W - 23
    'ISENTO PIS E COFINS? - X - 24
    '??? - eg_centro_aux-aliquota      EQ '0.00'
    'ALIQUOTA ICMS - Z - 26
    '%ALIQUOTA IPI - AA - 27
    'COD.SUBST.TRIBUT. - AB - 28
    '??? - eg_centro_aux-mva           EQ '0'.
    ''''''''''''''''''''''''''''''''''''''''
    If Cells(lin, 23) = "SIM" And _
       Cells(lin, 24) = "SIM" And _
       Cells(lin, 26) = 0 And _
       Cells(lin, 27) = 0 And _
       Cells(lin, 28) = 0 Then
       '
       'Cells(LIN, 35) = "ZX" 'ZX: ST Recolhida Anteriormente + PIS/COFINS
       CODIVA = "ZX"
       ''''''''''''''''''''''''''''''''''''
       'ELSEIF ( eg_centro_aux-prodfabricado      EQ '1' OR
       '* 016 - Fim Alteração - Inclusão do campo CEST
       '
       'eg_centro_aux-prodfabricado      EQ '0' )
       'AND eg_centro_aux-fonecedoropt  EQ '1'
       'AND eg_centro_aux-prodicms4     EQ '0'
       'AND eg_centro_aux-prodisentopis EQ '1'
       'AND eg_centro_aux-produtost     EQ '0'
       'AND eg_centro_aux-aliquotaicms  EQ '0.00'
       'AND eg_centro_aux-aliquotaipi   EQ '0.00'.
       'eg_centro_aux-codiva = 'ZA'.
       '
       'FABRICADO OU REVENDIDO - S - 19
       'SIMPLES NACIONAL? - T - 20
       'ISENTO OU IMUNE ICMS? - W - 23
       'ISENTO PIS E COFINS? - X - 24
       'COD.SUBST.TRIBUT. - AB - 28
       'ALIQUOTA ICMS - Z - 26
       '%ALIQUOTA IPI - AA - 27
    Else
      If Cells(lin, 19) = "Fabricado" And _
         Cells(lin, 20) = "NÃO" And _
         Cells(lin, 23) = "SIM" And _
         Cells(lin, 24) = "NAO" And _
         Cells(lin, 28) = 0 And _
         Cells(lin, 26) = 0 And _
         Cells(lin, 27) = 0 Then
         '
         'Cells(LIN, 35) = "ZA" 'ZA: COMPRA PIS + COFINS
         CODIVA = "ZA"
         ''''''''''''''''''''''''''''''''''''
         'ELSEIF eg_centro_aux-prodfabricado    EQ '0'
         'AND eg_centro_aux-fonecedoropt        EQ '1'
         'AND eg_centro_aux-prodicms4           EQ '1'
         'AND eg_centro_aux-prodisentopis       EQ '1'
         'AND eg_centro_aux-produtost           EQ '0'
         'AND eg_centro_aux-aliquotaicms        NE '0.00'
         'AND eg_centro_aux-aliquotaipi         EQ '0.00'.
         'eg_centro_aux-codiva = 'ZB'.
         'FABRICADO OU REVENDIDO - S - 19
         'SIMPLES NACIONAL? - T - 20
         'ISENTO OU IMUNE ICMS? - W - 23
         'ISENTO PIS E COFINS? - X - 24
         'COD.SUBST.TRIBUT. - AB - 28
         'ALIQUOTA ICMS - Z - 26
         '%ALIQUOTA IPI - AA - 27
      Else
        If Cells(lin, 19) = "Fabricado" And _
           Cells(lin, 20) = "NÃO" And _
           Cells(lin, 23) = "NÃO" And _
           Cells(lin, 24) = "NÃO" And _
           Cells(lin, 28) = 0 And _
           Cells(lin, 26) <> 0 And _
           Cells(lin, 27) = 0 Then
           '
           'Cells(LIN, 35) = "ZB" 'ZB: COMPRA ICMS +PIS + COFINS
           CODIVA = "ZB"
           '
           'ELSEIF eg_centro_aux-prodfabricado  EQ '1'
           'AND eg_centro_aux-fonecedoropt      EQ '1'
           'AND eg_centro_aux-prodicms4         EQ '1'
           'AND eg_centro_aux-prodisentopis     EQ '1'
           'AND eg_centro_aux-produtost         EQ '0'
           'AND eg_centro_aux-aliquotaicms      NE '0.00'
           'AND eg_centro_aux-aliquotaipi       NE '0.00'.
           'eg_centro_aux-codiva = 'ZC'.
           '
           'FABRICADO OU REVENDIDO - S - 19
           'SIMPLES NACIONAL? - T - 20
           'ISENTO OU IMUNE ICMS? - W - 23
           'ISENTO PIS E COFINS? - X - 24
           'COD.SUBST.TRIBUT. - AB - 28
           'ALIQUOTA ICMS - Z - 26
           '%ALIQUOTA IPI - AA - 27
        Else
          If Cells(lin, 19) = "Revendido" And _
             Cells(lin, 20) = "NÃO" And _
             Cells(lin, 23) = "NÃO" And _
             Cells(lin, 24) = "NÃO" And _
             Cells(lin, 28) = 0 And _
             Cells(lin, 26) <> 0 And _
             Cells(lin, 27) <> 0 Then
             '
             'Cells(LIN, 35) = "ZC" 'ZC: COMPRA ICMS + IPI + PIS + COFINS
             CODIVA = "ZC"
             '
             'ELSEIF eg_centro_aux-prodfabricado  EQ '1'
             'AND eg_centro_aux-fonecedoropt      EQ '1'
             'AND eg_centro_aux-prodicms4         EQ '1'
             'AND eg_centro_aux-prodisentopis     EQ '1'
             'AND eg_centro_aux-produtost         EQ '1'
             'AND eg_centro_aux-aliquotaicms      NE '0.00'
             'AND eg_centro_aux-aliquotaipi       NE '0.00'.
             'eg_centro_aux-codiva = 'ZD'.
             '
             'FABRICADO OU REVENDIDO - S - 19
             'SIMPLES NACIONAL? - T - 20
             'ISENTO OU IMUNE ICMS? - W - 23
             'ISENTO PIS E COFINS? - X - 24
             'COD.SUBST.TRIBUT. - AB - 28
             'ALIQUOTA ICMS - Z - 26
             '%ALIQUOTA IPI - AA - 27
          Else
            If Cells(lin, 19) = "Revendido" And _
               Cells(lin, 20) = "NÃO" And _
               Cells(lin, 23) = "NÃO" And _
               Cells(lin, 24) = "NÃO" And _
               Cells(lin, 28) = 1 And _
               Cells(lin, 26) <> 0 And _
               Cells(lin, 27) <> 0 Then
               '
               'Cells(LIN, 35) = "ZD" 'ZD: COMPRA ICMS + ST(CDNF) + IPI + PIS + COFINS
               CODIVA = "ZD"
               '
               'ELSEIF eg_centro_aux-prodfabricado  EQ '1'
               'AND eg_centro_aux-fonecedoropt      EQ '1'
               'AND eg_centro_aux-prodicms4         EQ '1'
               'AND eg_centro_aux-prodisentopis     EQ '1'
               'AND eg_centro_aux-produtost         EQ '2'
               'AND eg_centro_aux-aliquotaicms      NE '0.00'
               'AND eg_centro_aux-aliquotaipi       NE '0.00'.
               'eg_centro_aux-codiva = 'ZE'.
               '
               'FABRICADO OU REVENDIDO - S - 19
               'SIMPLES NACIONAL? - T - 20
               'ISENTO OU IMUNE ICMS? - W - 23
               'ISENTO PIS E COFINS? - X - 24
               'COD.SUBST.TRIBUT. - AB - 28
               'ALIQUOTA ICMS - Z - 26
               '%ALIQUOTA IPI - AA - 27
            Else
              If Cells(lin, 19) = "Revendido" And _
                 Cells(lin, 20) = "NÃO" And _
                 Cells(lin, 23) = "NÃO" And _
                 Cells(lin, 24) = "NÃO" And _
                 Cells(lin, 28) = 2 And _
                 Cells(lin, 26) <> 0 And _
                 Cells(lin, 27) <> 0 Then
                 '
                'Cells(LIN, 35) = "ZE" 'ZE: COMPRA ICMS + ST(SDNF) + IPI + PIS + COFINS
                CODIVA = "ZE"
                '
                 'ELSEIF eg_centro_aux-prodfabricado  EQ '0'
                 'AND eg_centro_aux-fonecedoropt      EQ '1'
                 'AND eg_centro_aux-prodicms4         EQ '1'
                 'AND eg_centro_aux-prodisentopis     EQ '1'
                 'AND eg_centro_aux-produtost         EQ '1'
                 'AND eg_centro_aux-aliquotaicms      NE '0.00'
                 'AND eg_centro_aux-aliquotaipi       EQ '0.00'.
                 'eg_centro_aux-codiva = 'ZF'.
                 '
                 'FABRICADO OU REVENDIDO - S - 19
                 'SIMPLES NACIONAL? - T - 20
                 'ISENTO OU IMUNE ICMS? - W - 23
                 'ISENTO PIS E COFINS? - X - 24
                 'COD.SUBST.TRIBUT. - AB - 28
                 'ALIQUOTA ICMS - Z - 26
                 '%ALIQUOTA IPI - AA - 27
              Else
                If Cells(lin, 19) = "Fabricado" And _
                   Cells(lin, 20) = "NÃO" And _
                   Cells(lin, 23) = "NÃO" And _
                   Cells(lin, 24) = "NÃO" And _
                   Cells(lin, 28) = 1 And _
                   Cells(lin, 26) <> 0 And _
                   Cells(lin, 27) = 0 Then
                   '
                   'Cells(LIN, 35) = "ZF" 'ZF: COMPRA ICMS + ST(CDNF) + PIS + COFINS
                   CODIVA = "ZF"
                   '
                   'ELSEIF eg_centro_aux-prodfabricado    EQ '0'
                   'AND eg_centro_aux-fonecedoropt        EQ '1'
                   'AND eg_centro_aux-prodicms4           EQ '1'
                   'AND eg_centro_aux-prodisentopis       EQ '1'
                   'AND eg_centro_aux-produtost           EQ '2'
                   'AND eg_centro_aux-aliquotaicms        NE '0.00'
                   'AND eg_centro_aux-aliquotaipi         EQ '0.00'.
                   'eg_centro_aux-codiva = 'ZG'.
                   '
                   'FABRICADO OU REVENDIDO - S - 19
                   'SIMPLES NACIONAL? - T - 20
                   'ISENTO OU IMUNE ICMS? - W - 23
                   'ISENTO PIS E COFINS? - X - 24
                   'COD.SUBST.TRIBUT. - AB - 28
                   'ALIQUOTA ICMS - Z - 26
                   '%ALIQUOTA IPI - AA - 27
                Else
                  If (Cells(lin, 19) = "Fabricado" Or _
                      Cells(lin, 19) = "Revendido") And _
                     Cells(lin, 20) = "NÃO" And _
                     Cells(lin, 23) = "NÃO" And _
                     Cells(lin, 24) = "NÃO" And _
                     Cells(lin, 28) = 2 And _
                     Cells(lin, 26) <> 0 And _
                     Cells(lin, 27) = 0 Then
                     '
                     'Cells(LIN, 35) = "ZG" 'ZG: COMPRA ICMS + ST(SDNF) + PIS + COFINS
                     CODIVA = "ZG"
                     '
                     'ELSEIF ( eg_centro_aux-prodfabricado  EQ '1' OR
                     '         eg_centro_aux-prodfabricado  EQ '0' )
                     'AND eg_centro_aux-fonecedoropt        EQ '1'
                     'AND eg_centro_aux-prodicms4           EQ '0'
                     'AND eg_centro_aux-prodisentopis       EQ '0'
                     'AND eg_centro_aux-produtost           EQ '0'
                     'AND eg_centro_aux-aliquotaicms        EQ '0.00'
                     'AND eg_centro_aux-aliquotaipi         EQ '0.00'.
                     'eg_centro_aux-codiva = 'ZN'.
                     '
                     'FABRICADO OU REVENDIDO - S - 19
                     'SIMPLES NACIONAL? - T - 20
                     'ISENTO OU IMUNE ICMS? - W - 23
                     'ISENTO PIS E COFINS? - X - 24
                     'COD.SUBST.TRIBUT. - AB - 28
                     'ALIQUOTA ICMS - Z - 26
                     '%ALIQUOTA IPI - AA - 27
                  Else
                    If (Cells(lin, 19) = "Fabricado" Or _
                        Cells(lin, 19) = "Revebdico") And _
                        Cells(lin, 20) = "NÃO" And _
                        Cells(lin, 23) = "SIM" And _
                        Cells(lin, 24) = "SIM" And _
                        Cells(lin, 28) = 0 And _
                        Cells(lin, 26) = 0 And _
                        Cells(lin, 27) = 0 Then
                        '
                        'Cells(LIN, 35) = "ZN" 'ZN : COMPRA ISENTOS DE TODOS OS IMPOSTOS
                        CODIVA = "ZN"
                        '
                        '*----------------------------------------------------------------
                        '*----------------------------------------------------------------
                        '
                        '*** Regra 2 - Simples Nacional - NÃO
                        '
                        'ELSEIF eg_centro_aux-fonecedoropt     EQ '1'
                        'AND eg_centro_aux-prodicms4     EQ '0'
                        'AND eg_centro_aux-prodisentopis EQ '1'
                        'AND eg_centro_aux-produtost     EQ '0'
                        'AND eg_centro_aux-aliquotaicms  EQ '0.00'
                        'AND eg_centro_aux-aliquotaipi   EQ '0.00'.
                        'eg_centro_aux-codiva = 'ZA'.
                        '
                        'SIMPLES NACIONAL? - T - 20
                        'ISENTO OU IMUNE ICMS? - W - 23
                        'ISENTO PIS E COFINS? - X - 24
                        'COD.SUBST.TRIBUT. - AB - 28
                        'ALIQUOTA ICMS - Z - 26
                        '%ALIQUOTA IPI - AA - 27
                    Else
                      If Cells(lin, 20) = "NÃO" And _
                         Cells(lin, 23) = "SIM" And _
                         Cells(lin, 24) = "NÃO" And _
                         Cells(lin, 28) = 0 And _
                         Cells(lin, 26) = 0 And _
                         Cells(lin, 27) = 0 Then
                         '
                         'Cells(LIN, 35) = "ZA" 'ZG: COMPRA ICMS + ST(SDNF) + PIS + COFINS
                         CODIVA = "ZA"
                         '
                         'ELSEIF  eg_centro_aux-fonecedoropt  EQ '1'
                         'AND eg_centro_aux-prodicms4         EQ '1'
                         'AND eg_centro_aux-prodisentopis     EQ '1'
                         'AND eg_centro_aux-produtost         EQ '0'
                         'AND eg_centro_aux-aliquotaicms      NE '0.00'
                         'AND eg_centro_aux-aliquotaipi       EQ '0.00'.
                         'eg_centro_aux-codiva = 'ZB'.
                         '
                         'SIMPLES NACIONAL? - T - 20
                         'ISENTO OU IMUNE ICMS? - W - 23
                         'ISENTO PIS E COFINS? - X - 24
                         'COD.SUBST.TRIBUT. - AB - 28
                         'ALIQUOTA ICMS - Z - 26
                         '%ALIQUOTA IPI - AA - 27
                      Else
                        If Cells(lin, 20) = "NÃO" And _
                           Cells(lin, 23) = "NÃO" And _
                           Cells(lin, 24) = "NÃO" And _
                           Cells(lin, 28) = 0 And _
                           Cells(lin, 26) <> 0 And _
                           Cells(lin, 27) = 0 Then
                           '
                           'Cells(LIN, 35) = "ZB" 'ZB: COMPRA ICMS +PIS + COFINS
                           CODIVA = "ZB"
                           '
                           'ELSEIF eg_centro_aux-fonecedoropt  EQ '1'
                           'AND eg_centro_aux-prodicms4        EQ '1'
                           'AND eg_centro_aux-prodisentopis    EQ '1'
                           'AND eg_centro_aux-produtost        EQ '0'
                           'AND eg_centro_aux-aliquotaicms     NE '0.00'
                           'AND eg_centro_aux-aliquotaipi      NE '0.00'.
                           'eg_centro_aux-codiva = 'ZC'.
                        Else
                          If Cells(lin, 20) = "NÃO" And _
                             Cells(lin, 23) = "NÃO" And _
                             Cells(lin, 24) = "NÃO" And _
                             Cells(lin, 28) = 0 And _
                             Cells(lin, 26) <> 0 And _
                             Cells(lin, 27) <> 0 Then
                             '
                             'Cells(LIN, 35) = "ZC" 'ZC: COMPRA ICMS + IPI + PIS + COFINS
                             CODIVA = "ZC"
                             '
                             'ELSEIF eg_centro_aux-fonecedoropt  EQ '1'
                             'AND eg_centro_aux-prodicms4        EQ '1'
                             'AND eg_centro_aux-prodisentopis    EQ '1'
                             'AND eg_centro_aux-produtost        EQ '1'
                             'AND eg_centro_aux-aliquotaicms     NE '0.00'
                             'AND eg_centro_aux-aliquotaipi      NE '0.00'.
                             'eg_centro_aux-codiva = 'ZD'.
                          Else
                            If Cells(lin, 20) = "NÃO" And _
                               Cells(lin, 23) = "NÃO" And _
                               Cells(lin, 24) = "NÃO" And _
                               Cells(lin, 28) = 1 And _
                               Cells(lin, 26) <> 0 And _
                               Cells(lin, 27) <> 0 Then
                               '
                               'Cells(LIN, 35) = "ZD" 'ZD: COMPRA ICMS + ST(CDNF) + IPI + PIS + COFINS
                               CODIVA = "ZD"
                               '
                               'ELSEIF eg_centro_aux-fonecedoropt  EQ '1'
                               'AND eg_centro_aux-prodicms4        EQ '1'
                               'AND eg_centro_aux-prodisentopis    EQ '1'
                               'AND eg_centro_aux-produtost        EQ '2'
                               'AND eg_centro_aux-aliquotaicms     NE '0.00'
                               'AND eg_centro_aux-aliquotaipi      NE '0.00'.
                               'eg_centro_aux-codiva = 'ZE'.
                            Else
                              If Cells(lin, 20) = "NÃO" And _
                                 Cells(lin, 23) = "NÃO" And _
                                 Cells(lin, 24) = "NÃO" And _
                                 Cells(lin, 28) = 2 And _
                                 Cells(lin, 26) <> 0 And _
                                 Cells(lin, 27) <> 0 Then
                                 '
                                 'Cells(LIN, 35) = "ZE" 'ZE: COMPRA ICMS + ST(SDNF) + IPI + PIS + COFINS
                                 CODIVA = "ZE"
                                 '
                                 'ELSEIF eg_centro_aux-fonecedoropt EQ '1'
                                 'AND eg_centro_aux-prodicms4       EQ '1'
                                 'AND eg_centro_aux-prodisentopis   EQ '1'
                                 'AND eg_centro_aux-produtost       EQ '1'
                                 'AND eg_centro_aux-aliquotaicms    NE '0.00'
                                 'AND eg_centro_aux-aliquotaipi     EQ '0.00'.
                                 'eg_centro_aux-codiva = 'ZF'.
                              Else
                                If Cells(lin, 20) = "NÃO" And _
                                   Cells(lin, 23) = "NÃO" And _
                                   Cells(lin, 24) = "NÃO" And _
                                   Cells(lin, 28) = 1 And _
                                   Cells(lin, 26) <> 0 And _
                                   Cells(lin, 27) = 0 Then
                                   '
                                   'Cells(LIN, 35) = "ZF" 'ZF: COMPRA ICMS + ST(CDNF) + PIS + COFINS
                                   CODIVA = "ZF"
                                Else 'YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
                                     'YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
                                  '*IVA - "YD - COMPRA ST + IPI + PIS + COFINS"
                                  'AND eg_centro_aux-prodisentopis     EQ '1'
                                  'AND eg_centro_aux-produtost         EQ '1'
                                  'AND eg_centro_aux-aliquotaipi       NE '0.00'.
                                  'COD.ANP NÃO PREENCHIDO
                                  'eg_centro_aux-codiva = 'YD'
                                  '
                                  'ISENTO PIS E COFINS? - X - 24
                                  'COD.SUBST.TRIBUT. - AB - 28
                                  'ALIQUOTA ICMS - Z - 26
                                  '%ALIQUOTA IPI - AA - 27
                                  'COD_ANP - AN - 40
                                  If Cells(lin, 24) = "NÃO" And _
                                     Cells(lin, 28) = 1 And _
                                     Cells(lin, 27) <> 0 And _
                                     (Cells(lin, 40) = "" Or Cells(lin, 40) = 0) Then
                                     'Cells(LIN, 35) = "YD" 'YD - COMPRA ST + IPI + PIS + COFINS
                                     CODIVA = "YD"
                                  Else
                                    '* IVA - "YE - COMPRA ST + PIS + COFINS"
                                    'AND eg_centro_aux-prodisentopis     EQ '1'
                                    'AND eg_centro_aux-produtost         EQ '2'
                                    'AND eg_centro_aux-aliquotaipi       NE '0.00'.
                                    'COD.ANP NÃO PREENCHIDO
                                    'eg_centro_aux-codiva = 'YE'
                                    '
                                    'ISENTO PIS E COFINS? - X - 24
                                    'COD.SUBST.TRIBUT. - AB - 28
                                    'ALIQUOTA ICMS - Z - 26
                                    '%ALIQUOTA IPI - AA - 27
                                    If Cells(lin, 24) = "NÃO" And _
                                       Cells(lin, 28) = 2 And _
                                       Cells(lin, 27) <> 0 And _
                                      (Cells(lin, 40) = "" Or Cells(lin, 40) = 0) Then
                                       '
                                       'Cells(LIN, 35) = "YE" 'YE - COMPRA ST + PIS + COFINS
                                       CODIVA = "YE"
                                    Else
                                      '
                                      '* IVA - "YF - COMPRA ST + PIS + COFINS"
                                      'AND eg_centro_aux-prodisentopis     EQ '1'
                                      'AND eg_centro_aux-produtost         EQ '1'
                                      'COD.ANP NÃO PREENCHIDO
                                      'eg_centro_aux-codiva = 'YF'
                                      '
                                      'ISENTO PIS E COFINS? - X - 24
                                      'COD.SUBST.TRIBUT. - AB - 28
                                      If Cells(lin, 24) = "NÃO" And _
                                         Cells(lin, 28) = 1 And _
                                        (Cells(lin, 40) = "" Or Cells(lin, 40) = 0) Then
                                         '
                                         'Cells(LIN, 35) = "YF" 'YE - COMPRA ST + PIS + COFINS
                                         CODIVA = "YF"
                                      Else
                                        '* IVA - "YG - COMPRA ST (SDNF) + PIS + COFINS"
                                        'AND eg_centro_aux-prodisentopis     EQ '1'
                                        'AND eg_centro_aux-produtost         EQ '2'
                                        'COD.ANP NÃO PREENCHIDO
                                        'eg_centro_aux-codiva = 'YG'
                                        '
                                        'ISENTO PIS E COFINS? - X - 24
                                        'COD.SUBST.TRIBUT. - AB - 28
                                        If Cells(lin, 24) = "NÃO" And _
                                           Cells(lin, 28) = 2 And _
                                          (Cells(lin, 40) = "" Or Cells(lin, 40) = 0) Then
                                           '
                                           'Cells(LIN, 35) = "YG" 'YE - COMPRA ST + PIS + COFINS
                                           CODIVA = "YG"
                                        Else
                                          '* IVA - "YH - COMPRA ST REVENDA + IPI + PIS + COFINS"
                                          'AND eg_centro_aux-prodisentopis     EQ '1'
                                          'AND eg_centro_aux-produtost         EQ '3'
                                          'AND eg_centro_aux-aliquotaipi       NE '0.00'.
                                          'eg_centro_aux-codiva = 'YH'
                                          '
                                          'ISENTO PIS E COFINS? - X - 24
                                          'COD.SUBST.TRIBUT. - AB - 28
                                          'ALIQUOTA ICMS - Z - 26
                                          '%ALIQUOTA IPI - AA - 27
                                          If Cells(lin, 24) = "NÃO" And _
                                             Cells(lin, 28) = 3 And _
                                             Cells(lin, 27) <> 0 Then
                                             '
                                             'Cells(LIN, 35) = "YH" 'YH - COMPRA ST REVENDA + IPI + PIS + COFINS
                                             CODIVA = "YH"
                                          Else
                                            '* IVA - "YI - COMPRA ICMS ISENTO  + PIS + COFINS"
                                            'AND eg_centro_aux-prodicms4         EQ '0'
                                            'AND eg_centro_aux-prodisentopis     EQ '1'
                                            'AND eg_centro_aux-aliquotaicms      EQ '0.00'
                                            'AND REDUCAO ICMA = 0
                                            'eg_centro_aux-codiva = 'YI'
                                            '
                                            'ISENTO OU IMUNE ICMS? - W - 23
                                            'ISENTO PIS E COFINS? - X - 24
                                            'ALIQUOTA ICMS - Z - 26
                                            '%REDUCAO ICNS - Y - 25
                                            '%ALIQUOTA IPI - AA - 27
                                            If Cells(lin, 23) = "SIM" And _
                                               Cells(lin, 24) = "NÃO" And _
                                               Cells(lin, 28) = 2 And _
                                               Cells(lin, 26) <> 0 And _
                                               Cells(lin, 27) <> 0 Then
                                               '
                                               'Cells(LIN, 35) = "YI" 'YI - COMPRA ICMS ISENTO  + PIS + COFINS
                                               CODIVA = "YI"
                                            Else
                                              '* IVA - "YM - ICMS ISENTO E COMPRA ST (CDNF) + IPI + PIS + COFINS"
                                              '  AND eg_centro_aux-prodicms4         EQ '0'
                                              '  AND eg_centro_aux-prodisentopis     EQ '1'
                                              '  AND eg_centro_aux-produtost         EQ '1'
                                              '  AND eg_centro_aux-aliquotaipi       NE '0.00'.
                                              'COD.ANP PREENCHIDO
                                              '    eg_centro_aux-codiva = 'YM'
                                              '
                                              'ISENTO PIS E COFINS? - X - 24
                                              'SIMPLES NACIONAL? - T - 20
                                              'ISENTO OU IMUNE ICMS? - W - 23
                                              'ISENTO PIS E COFINS? - X - 24
                                              'COD.SUBST.TRIBUT. - AB - 28
                                              'ALIQUOTA ICMS - Z - 26
                                              '%ALIQUOTA IPI - AA - 27
                                              If Cells(lin, 23) = "SIM" And _
                                                 Cells(lin, 24) = "NÃO" And _
                                                 Cells(lin, 28) = 1 And _
                                                 Cells(lin, 27) <> 0 And _
                                                 (Cells(lin, 40) <> "" And Cells(lin, 40) <> 0) Then
                                                 '
                                                 'Cells(LIN, 35) = "YM" 'YM - ICMS ISENTO E COMPRA ST (CDNF) + IPI + PIS + COFINS
                                                 CODIVA = "YM"
                                              Else
                                                '* IVA - "YN - ICMS ISENTO E COMPRA ST (SDNF) + IPI + PIS + COFINS"
                                                'AND eg_centro_aux-prodicms4         EQ 'O'
                                                'AND eg_centro_aux-prodisentopis     EQ '1'
                                                'AND eg_centro_aux-produtost         EQ '2'
                                                'AND eg_centro_aux-aliquotaipi       NE '0.00'.
                                              'COD.ANP PREENCHIDO
                                                'eg_centro_aux-codiva = 'YN'
                                                'ISENTO PIS E COFINS? - X - 24
                                                'SIMPLES NACIONAL? - T - 20
                                                'ISENTO OU IMUNE ICMS? - W - 23
                                                'ISENTO PIS E COFINS? - X - 24
                                                'COD.SUBST.TRIBUT. - AB - 28
                                                'ALIQUOTA ICMS - Z - 26
                                                '%ALIQUOTA IPI - AA - 27
                                                If Cells(lin, 23) = "SIM" And _
                                                   Cells(lin, 24) = "NÃO" And _
                                                   Cells(lin, 28) = 2 And _
                                                   Cells(lin, 27) <> 0 And _
                                                 (Cells(lin, 40) <> "" And Cells(lin, 40) <> 0) Then
                                                   '
                                                   'Cells(LIN, 35) = "YN" 'YN - ICMS ISENTO E COMPRA ST (SDNF) + IPI + PIS + COFINS
                                                   CODIVA = "YN"
                                                Else
                                                  '* IVA - "YO - ICMS ISENTO E COMPRA ST (CDNF) + PIS + COFINS"
                                                  'AND eg_centro_aux-prodicms4         EQ 'O'
                                                  'AND eg_centro_aux-prodisentopis     EQ '1'
                                                  'AND eg_centro_aux-produtost         EQ '1'
                                              'COD.ANP PREENCHIDO
                                                  'eg_centro_aux-codiva = 'YO'
                                                  'ISENTO PIS E COFINS? - X - 24
                                                  'SIMPLES NACIONAL? - T - 20
                                                  'ISENTO OU IMUNE ICMS? - W - 23
                                                  'ISENTO PIS E COFINS? - X - 24
                                                  'COD.SUBST.TRIBUT. - AB - 28
                                                  'ALIQUOTA ICMS - Z - 26
                                                  '%ALIQUOTA IPI - AA - 27
                                                  If Cells(lin, 23) = "SIM" And _
                                                     Cells(lin, 24) = "NÃO" And _
                                                     Cells(lin, 28) = 1 And _
                                                 (Cells(lin, 40) <> "" And Cells(lin, 40) <> 0) Then
                                                     '
                                                     'Cells(LIN, 35) = "YO" 'YO - ICMS ISENTO E COMPRA ST (CDNF) + PIS + COFINS
                                                     CODIVA = "YO"
                                                  Else
                                                    '* IVA - "YP - ICMS ISENTO E COMPRA ST (SDNF) + IPI + PIS + COFINS"
                                                    'AND eg_centro_aux-prodicms4         EQ 'O'
                                                    'AND eg_centro_aux-prodisentopis     EQ '1'
                                                    'AND eg_centro_aux-produtost         EQ '2'
                                              'COD.ANP PREENCHIDO
                                                    'eg_centro_aux-codiva = 'YP'
                                                    'ISENTO PIS E COFINS? - X - 24
                                                    'SIMPLES NACIONAL? - T - 20
                                                    'ISENTO OU IMUNE ICMS? - W - 23
                                                    'ISENTO PIS E COFINS? - X - 24
                                                    'COD.SUBST.TRIBUT. - AB - 28
                                                    'ALIQUOTA ICMS - Z - 26
                                                    '%ALIQUOTA IPI - AA - 27
                                                    If Cells(lin, 23) = "SIM" And _
                                                       Cells(lin, 24) = "NÃO" And _
                                                       Cells(lin, 28) = 2 And _
                                                 (Cells(lin, 40) <> "" And Cells(lin, 40) <> 0) Then
                                                       '
                                                       'Cells(LIN, 35) = "YP" 'YP - ICMS ISENTO E COMPRA ST (SDNF) + IPI + PIS + COFINS
                                                       CODIVA = "YP"
                                                    Else
                                                      CODIVA = "ZN"
                                                    End If
                                                  End If
                                                End If
                                              End If
                                            End If
                                          End If
                                        End If
                                      End If
                                    End If
                                  End If
                                End If
                              End If
                            End If
                          End If
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    End If
  End If
End If
'
ThisWorkbook.Sheets("Dados Fiscais").Unprotect Key
Cells(lin, 35).Value = CODIVA
'
Application.Goto Reference:="IVA_II_PROCV"
'
Selection.Find(What:=CODIVA, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
'
DESCR_CODIVA = Cells(ActiveCell.Row, ActiveCell.Column + 1)
'
Sheets("Dados Fiscais").Select
Cells(lin, 36).Value = DESCR_CODIVA
'
End Sub '
'CAUCULO IVA
'YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY

'* IVA - "YD - COMPRA ST + IPI + PIS + COFINS"
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-produtost         EQ '1'
'  AND eg_centro_aux-aliquotaipi       NE '0.00'.
'    eg_centro_aux-codiva = 'YD'
'
'* IVA - "YE - COMPRA ST + PIS + COFINS"
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-produtost         EQ '2'
'  AND eg_centro_aux-aliquotaipi       NE '0.00'.
'    eg_centro_aux-codiva = 'YE'
'
'* IVA - "YF - COMPRA ST + PIS + COFINS"
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-produtost         EQ '1'
'    eg_centro_aux-codiva = 'YF'
'
'* IVA - "YG - COMPRA ST (SDNF) + PIS + COFINS"
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-produtost         EQ '2'
'    eg_centro_aux-codiva = 'YE'
'
'* IVA - "YH - COMPRA ST REVENDA + IPI + PIS + COFINS"
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-produtost         EQ '3'
'  AND eg_centro_aux-aliquotaipi       NE '0.00'.
'    eg_centro_aux-codiva = 'YH'
'
'* IVA - "YH - COMPRA ST REVENDA + IPI + PIS + COFINS"
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-produtost         EQ '3'
'  AND eg_centro_aux-aliquotaipi       NE '0.00'.
'    eg_centro_aux-codiva = 'YH'
'
'
'* IVA - "YI - COMPRA ICMS ISENTO  + PIS + COFINS"
'  AND eg_centro_aux-prodicms4         EQ '0'
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-aliquotaicms      EQ '0.00'
'  AND REDUCAO ICMA = 0
'    eg_centro_aux-codiva = 'YI'
'
'* IVA - "YM - ICMS ISENTO E COMPRA ST (CDNF) + IPI + PIS + COFINS"
'  AND eg_centro_aux-prodicms4         EQ '0'
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-produtost         EQ '2'
'  AND eg_centro_aux-aliquotaipi       NE '0.00'.
'    eg_centro_aux-codiva = 'YM'
'
'* IVA - "YN - ICMS ISENTO E COMPRA ST (SDNF) + IPI + PIS + COFINS"
'  AND eg_centro_aux-prodicms4         EQ 'O'
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-produtost         EQ '2'
'  AND eg_centro_aux-aliquotaipi       NE '0.00'.
'    eg_centro_aux-codiva = 'YN'
'
'* IVA - "YO - ICMS ISENTO E COMPRA ST (CDNF) + PIS + COFINS"
'  AND eg_centro_aux-prodicms4         EQ 'O'
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-produtost         EQ '1'
'    eg_centro_aux-codiva = 'YO'
'
'* IVA - "YP - ICMS ISENTO E COMPRA ST (SDNF) + IPI + PIS + COFINS"
'  AND eg_centro_aux-prodicms4         EQ 'O'
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-produtost         EQ '2'
'    eg_centro_aux-codiva = 'YP'
'
'** Regra 1 - Simples Nacional - NÃO e Fabricado ou Revendido
'* IVA - "ZA - LMB: COMPRA PIS + COFINS"
'* 016 - Início Alteração - Inclusão do campo CEST
'*  IF ( eg_centro_aux-prodfabricado      EQ '1' OR
'  DATA: vl_p   TYPE p DECIMALS 2,
'        vl_anp TYPE mara-zzcod_anp.
'
'* IVA - "ZQ - LMB: Isento ICMS + PIS/COFINS + Al. IPI > 0
'  MOVE  eg_centro_aux-aliquotaipi TO vl_p.
'  IF  eg_centro_aux-prodicms4         EQ '0'
'      AND eg_centro_aux-prodisentopis EQ '1'
'      AND vl_p                        GT 0.
'    eg_centro_aux-codiva = 'ZQ'.
'
'* IVA - "ZS - LMB: ST Recolhida Anteriormente + PIS/COFINS
'  ELSEIF eg_centro_aux-prodisentopis EQ '1'
'         AND eg_centro_aux-produtost EQ '3'.
'    eg_centro_aux-codiva = 'ZS'.
'
'* IVA - "ZX - LMB: ST Recolhida Anteriormente + PIS/COFINS
'  ELSEIF eg_centro_aux-prodicms4         EQ '0'
'         AND eg_centro_aux-prodisentopis EQ '0'
'         AND eg_centro_aux-aliquota      EQ '0.00'
'         AND eg_centro_aux-aliquotaicms  EQ '0.00'
'         AND eg_centro_aux-aliquotaipi   EQ '0.00'
'         AND eg_centro_aux-produtost     EQ '0'
'         AND eg_centro_aux-mva           EQ '0'.
'    eg_centro_aux-codiva = 'ZX'.
'
'* IVA - "ZA - LMB: COMPRA PIS + COFINS"
'  ELSEIF ( eg_centro_aux-prodfabricado      EQ '1' OR
'* 016 - Fim Alteração - Inclusão do campo CEST
'
'       eg_centro_aux-prodfabricado      EQ '0' )
'        AND eg_centro_aux-fonecedoropt  EQ '1'
'        AND eg_centro_aux-prodicms4     EQ '0'
'        AND eg_centro_aux-prodisentopis EQ '1'
'        AND eg_centro_aux-produtost     EQ '0'
'        AND eg_centro_aux-aliquotaicms  EQ '0.00'
'        AND eg_centro_aux-aliquotaipi   EQ '0.00'.
'    eg_centro_aux-codiva = 'ZA'.
'
'* IVA - "ZB - LMB: COMPRA ICMS +PIS + COFINS"
'  ELSEIF eg_centro_aux-prodfabricado    EQ '0'
'  AND eg_centro_aux-fonecedoropt        EQ '1'
'  AND eg_centro_aux-prodicms4           EQ '1'
'  AND eg_centro_aux-prodisentopis       EQ '1'
'  AND eg_centro_aux-produtost           EQ '0'
'  AND eg_centro_aux-aliquotaicms        NE '0.00'
'  AND eg_centro_aux-aliquotaipi         EQ '0.00'.
'    eg_centro_aux-codiva = 'ZB'.
'
'* IVA - "ZC - LMB: COMPRA ICMS + IPI + PIS + COFINS"
'  ELSEIF eg_centro_aux-prodfabricado  EQ '1'
'  AND eg_centro_aux-fonecedoropt      EQ '1'
'  AND eg_centro_aux-prodicms4         EQ '1'
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-produtost         EQ '0'
'  AND eg_centro_aux-aliquotaicms      NE '0.00'
'  AND eg_centro_aux-aliquotaipi       NE '0.00'.
'    eg_centro_aux-codiva = 'ZC'.
'
'* IVA - "ZD - LMB: COMPRA ICMS + ST(CDNF) + IPI + PIS + COFINS"
'  ELSEIF eg_centro_aux-prodfabricado  EQ '1'
'  AND eg_centro_aux-fonecedoropt      EQ '1'
'  AND eg_centro_aux-prodicms4         EQ '1'
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-produtost         EQ '1'
'  AND eg_centro_aux-aliquotaicms      NE '0.00'
'  AND eg_centro_aux-aliquotaipi       NE '0.00'.
'    eg_centro_aux-codiva = 'ZD'.
'
'* IVA - "ZE - LMB: COMPRA ICMS + ST(SDNF) + IPI + PIS + COFINS"
'  ELSEIF eg_centro_aux-prodfabricado  EQ '1'
'  AND eg_centro_aux-fonecedoropt      EQ '1'
'  AND eg_centro_aux-prodicms4         EQ '1'
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-produtost         EQ '2'
'  AND eg_centro_aux-aliquotaicms      NE '0.00'
'  AND eg_centro_aux-aliquotaipi       NE '0.00'.
'    eg_centro_aux-codiva = 'ZE'.
'
'* IVA - "ZF - LMB: COMPRA ICMS + ST(CDNF) + PIS + COFINS"
'  ELSEIF eg_centro_aux-prodfabricado  EQ '0'
'  AND eg_centro_aux-fonecedoropt      EQ '1'
'  AND eg_centro_aux-prodicms4         EQ '1'
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-produtost         EQ '1'
'  AND eg_centro_aux-aliquotaicms      NE '0.00'
'  AND eg_centro_aux-aliquotaipi       EQ '0.00'.
'    eg_centro_aux-codiva = 'ZF'.
'
'* IVA - "ZG - LMB: COMPRA ICMS + ST(SDNF) + PIS + COFINS"
'  ELSEIF eg_centro_aux-prodfabricado    EQ '0'
'  AND eg_centro_aux-fonecedoropt        EQ '1'
'  AND eg_centro_aux-prodicms4           EQ '1'
'  AND eg_centro_aux-prodisentopis       EQ '1'
'  AND eg_centro_aux-produtost           EQ '2'
'  AND eg_centro_aux-aliquotaicms        NE '0.00'
'  AND eg_centro_aux-aliquotaipi         EQ '0.00'.
'    eg_centro_aux-codiva = 'ZG'.
'
'* IVA - "ZN - LMB: COMPRA ISENTOS DE TODOS OS IMPOSTOS"
'  ELSEIF ( eg_centro_aux-prodfabricado  EQ '1' OR
'           eg_centro_aux-prodfabricado  EQ '0' )
'  AND eg_centro_aux-fonecedoropt        EQ '1'
'  AND eg_centro_aux-prodicms4           EQ '0'
'  AND eg_centro_aux-prodisentopis       EQ '0'
'  AND eg_centro_aux-produtost           EQ '0'
'  AND eg_centro_aux-aliquotaicms        EQ '0.00'
'  AND eg_centro_aux-aliquotaipi         EQ '0.00'.
'    eg_centro_aux-codiva = 'ZN'.
'
'*----------------------------------------------------------------
'*----------------------------------------------------------------
'
'*** Regra 2 - Simples Nacional - NÃO
'* IVA - "ZA - LMB: COMPRA PIS + COFINS"
'  ELSEIF eg_centro_aux-fonecedoropt     EQ '1'
'        AND eg_centro_aux-prodicms4     EQ '0'
'        AND eg_centro_aux-prodisentopis EQ '1'
'        AND eg_centro_aux-produtost     EQ '0'
'        AND eg_centro_aux-aliquotaicms  EQ '0.00'
'        AND eg_centro_aux-aliquotaipi   EQ '0.00'.
'    eg_centro_aux-codiva = 'ZA'.
'
'* IVA - "ZB - LMB: COMPRA ICMS +PIS + COFINS"
'  ELSEIF  eg_centro_aux-fonecedoropt  EQ '1'
'  AND eg_centro_aux-prodicms4         EQ '1'
'  AND eg_centro_aux-prodisentopis     EQ '1'
'  AND eg_centro_aux-produtost         EQ '0'
'  AND eg_centro_aux-aliquotaicms      NE '0.00'
'  AND eg_centro_aux-aliquotaipi       EQ '0.00'.
'    eg_centro_aux-codiva = 'ZB'.
'
'* IVA - "ZC - LMB: COMPRA ICMS + IPI + PIS + COFINS"
'  ELSEIF eg_centro_aux-fonecedoropt  EQ '1'
'  AND eg_centro_aux-prodicms4        EQ '1'
'  AND eg_centro_aux-prodisentopis    EQ '1'
'  AND eg_centro_aux-produtost        EQ '0'
'  AND eg_centro_aux-aliquotaicms     NE '0.00'
'  AND eg_centro_aux-aliquotaipi      NE '0.00'.
'    eg_centro_aux-codiva = 'ZC'.
'
'* IVA - "ZD - LMB: COMPRA ICMS + ST(CDNF) + IPI + PIS + COFINS"
'  ELSEIF eg_centro_aux-fonecedoropt  EQ '1'
'  AND eg_centro_aux-prodicms4        EQ '1'
'  AND eg_centro_aux-prodisentopis    EQ '1'
'  AND eg_centro_aux-produtost        EQ '1'
'  AND eg_centro_aux-aliquotaicms     NE '0.00'
'  AND eg_centro_aux-aliquotaipi      NE '0.00'.
'    eg_centro_aux-codiva = 'ZD'.
'
'* IVA - "ZE - LMB: COMPRA ICMS + ST(SDNF) + IPI + PIS + COFINS"
'  ELSEIF eg_centro_aux-fonecedoropt  EQ '1'
'  AND eg_centro_aux-prodicms4        EQ '1'
'  AND eg_centro_aux-prodisentopis    EQ '1'
'  AND eg_centro_aux-produtost        EQ '2'
'  AND eg_centro_aux-aliquotaicms     NE '0.00'
'  AND eg_centro_aux-aliquotaipi      NE '0.00'.
'    eg_centro_aux-codiva = 'ZE'.
'
'* IVA - "ZF - LMB: COMPRA ICMS + ST(CDNF) + PIS + COFINS"
'  ELSEIF eg_centro_aux-fonecedoropt EQ '1'
'  AND eg_centro_aux-prodicms4       EQ '1'
'  AND eg_centro_aux-prodisentopis   EQ '1'
'  AND eg_centro_aux-produtost       EQ '1'
'  AND eg_centro_aux-aliquotaicms    NE '0.00'
'  AND eg_centro_aux-aliquotaipi     EQ '0.00'.
'    eg_centro_aux-codiva = 'ZF'.
'
'* IVA - "ZG - LMB: COMPRA ICMS + ST(SDNF) + PIS + COFINS"
'  ELSEIF eg_centro_aux

Public Sub CARREGA_FORNEC_VIRTUAL()
        Dim lin  As Integer
        Dim col As Integer
        Dim LIN_PAR  As Integer
        Dim COL_PAR As Integer
        Dim LIN_TBL  As Integer
        Dim COL_TBL As Integer
        '
        'CENTRO  DESCR. CENTRO   REGIÃO DE ENTRADA   COD.PERFILDISTRIBUICAO  DESCR.PERFILDISTRIBUICAO
        Dim CENTRO As String
        Dim DESCR_CENTRO As String
        Dim REGIO As String
        Dim PERFIL As String
        Dim DESCR_PERFIL As String
        '
        Dim COD_FORNEC_SAP As String
        Dim RAZAO_SOCIAL As String
        Dim CNPJ_FORNEC As String
        '
        Dim SUCESSO As Boolean
        '
        Application.ScreenUpdating = False
        '
        ActiveWorkbook.Sheets("Escolha Perfil Distribuicao").Select
        ActiveWorkbook.Sheets("Escolha Perfil Distribuicao").Unprotect Key
        'ActiveWorkbook.Unprotect Key
        '
        If Sheets("TABELAS").Visible = False Then
           Sheets("TABELAS").Visible = True
        End If
        ActiveWorkbook.Sheets("Escolha Perfil Distribuicao").Select
        Application.Goto Reference:="INICIO_ESCOLHA_SUBSORTIMENTO"
        LIN_TBL = ActiveCell.Row
        COL_TBL = ActiveCell.Column
        '
        ActiveWorkbook.Sheets("PARAMETROS").Select
        ThisWorkbook.Sheets("PARAMETROS").Unprotect Key
        Application.Goto Reference:="INICIO_PRACAS"
        LIN_PAR = ActiveCell.Row
        COL_PAR = ActiveCell.Column
        '
        'LIMPA TABELA_LOCAIS
        Application.Goto Reference:="TABELA_LOCAIS"
        Selection.ClearContents
        '
        Application.EnableEvents = False
        Application.Goto Reference:="FORNEC_VIRTUAL_CLEAR"
        Selection.ClearContents
        Application.EnableEvents = True
        '
        ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Select
        Application.Goto Reference:="INICIO_PRACAS_TBL"
        lin = ActiveCell.Row
        col = ActiveCell.Column
        '
        COD_FORNEC_SAP = 0 'provocar primeira quebra
        PERFIL = "0"
        '
        Do While ActiveCell.Value <> ""
           '
           'CENTRO  DESCR. CENTRO   REGIÃO DE ENTRADA   COD.PERFILDISTRIBUICAO  DESCR.PERFILDISTRIBUICAO
           'primeira linha TODAS AS PRACAS
           If Cells(lin, col + 2).Value = "Verdadeiro" Then
              If Cells(lin, col - 1).Value <> "ALL" Then
                 CENTRO = Cells(lin, col - 1).Value
                 DESCR_CENTRO = Cells(lin, col).Value
                 REGIO = Cells(lin, col + 1).Value
                 'PERFIL = Cells(LIN, COL + 3).Value
                 DESCR_PERFIL = Cells(lin, col + 4).Value
                 '
                 RAZAO_SOCIAL = Cells(lin, col + 5).Value
                 CNPJ_FORNEC = "'" & Cells(lin, col + 6).Value
                 '
                 'pergunta da quebra que vai gravar "COD.FORNEC SAP  CNPJ    RAZAO SOCIAL    FORNECEDOR VIRTUAL  SUBSORTIMENTO"
                 'na tabela ao lado ESCOLHA SUBSORTIMENTO
                 'na aba "Escolha Perfil Distribuicao"
                 If COD_FORNEC_SAP <> Cells(lin, col + 7).Value Or _
                    PERFIL <> Cells(lin, col + 7).Value Then
                    '
                    COD_FORNEC_SAP = Right("0" & Cells(lin, col + 7).Value, 10)
                    PERFIL = Cells(lin, col + 3).Value
                    '
                    SUCESSO = ADDITEM_COMBOBOX_FORNEC(lin, LIN_TBL, COL_TBL, DESCR_CENTRO, COD_FORNEC_SAP, _
                              CNPJ_FORNEC, RAZAO_SOCIAL, CENTRO, PERFIL)
                    If Not SUCESSO Then
                       MsgBox ("Não encontrado Fornecedor Virtual para o CNPJ/CENTRO/PERFIL: " & _
                                CNPJ_FORNEC & "/" & CENTRO & "/" & PERFIL & ".")
                       Sheets("Inicio").PERFIL.BackColor = &HFF& 'VERMELHO
                       Sheets("Escolha Perfil Distribuicao").SUBSORTIMENTO.BackColor = &HFF& 'VERMELHO
                    Else
                       LIN_TBL = LIN_TBL + 1
                    End If
                    '
                 End If           '
                    
                 ActiveWorkbook.Sheets("PARAMETROS").Activate
                 '
                 ActiveWorkbook.Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR).Select
                 ActiveWorkbook.Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR).Value = CENTRO
                 ActiveWorkbook.Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR + 1).Value = DESCR_CENTRO
                 ActiveWorkbook.Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR + 2).Value = REGIO
                 ActiveWorkbook.Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR + 3).Value = PERFIL
                 ActiveWorkbook.Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR + 4).Value = DESCR_PERFIL
                 ActiveWorkbook.Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR + 5).Value = Right("0" & COD_FORNEC_SAP, 10)
                 ActiveWorkbook.Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR + 6).Value = RAZAO_SOCIAL
                 ActiveWorkbook.Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR + 7).Value = CNPJ_FORNEC
                 LIN_PAR = LIN_PAR + 1
                 ActiveWorkbook.Sheets("PARAMETROS").Cells(LIN_PAR, COL_PAR).Select
                 '
              End If
           End If
           ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Activate
           lin = lin + 1
           ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Cells(lin, col).Select
        Loop
        '
        Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
        Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
        Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
        Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
        Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
        '
        'MARCA CALCULO DADOS FISCAIS
        If Sheets("TRABALHO").Visible = False Then
           ActiveWorkbook.Unprotect Key
           Sheets("TRABALHO").Visible = True
           Application.Goto Reference:="DADOS_FISCAIS_FIRSTONE"
           If ActiveCell.Value <> "X" Then
              ActiveCell.Value = "X"
           End If
        Else
          Sheets("TRABALHO").Select
          Application.Goto Reference:="DADOS_FISCAIS_FIRSTONE"
          If ActiveCell.Value <> "X" Then
             ActiveCell.Value = "X"
          End If
        End If
        Sheets("TRABALHO").Visible = False
        '
        ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Select
        Application.Goto Reference:="CELLS_COLUNA_OH"
        lin = ActiveCell.Row
        col = ActiveCell.Column
        '
        ActiveWorkbook.Save
        MsgBox "Pasta de Trabalho Salva.", vbOKOnly, "INTERFACE FORNECEDOR"
        Sheets("Escolha Perfil Distribuicao").SALVA.Enabled = False
        '
        Application.Goto Reference:="TITULO_SUBSORTIMENTO"
        '
        MsgBox ("A coluna FORNECEDOR VIRTUAL deve ser preenchida para cada linha informada." & _
               "Há linhas sem essa informação. Verifique")
        '
        ThisWorkbook.Sheets("Escolha Perfil Distribuicao").Protect Key, DrawingObjects:=False, Contents:=True, Scenarios:= _
            False
End Sub


Sub Verifica_Arquivo()
    Dim strPath As Variant
    Dim strFile As Variant
    Dim strCheck As Boolean
    Dim WORK_NAME As Variant
    Dim SUCESSO As Boolean
    'apaga mensagens de impotacao
    '
    If Sheets("Mensagens Importacao Tabelas").Visible = False Then
       Sheets("Mensagens Importacao Tabelas").Visible = True
    End If
    '
    Sheets("Mensagens Importacao Tabelas").Select
    ActiveSheet.Unprotect Key
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    '
    Transforme_Form.Show vbModeless
    Transforme_Form.Caption = "IMPORTA E TRANSFORMA TABELAS SAP"
    '
    If ThisWorkbook.Sheets("TRABALHO").Visible = False Then
       ThisWorkbook.Sheets("TRABALHO").Visible = True
    End If
    '
    ThisWorkbook.Sheets("TRABALHO").Select
    Application.Goto Reference:="DIRETORIO_ARQUIVOS_IMPORTACAO"
    strPath = ActiveCell.Value
    '
    Application.Goto Reference:="PRIMEIRA_TABELA_IMPORTAR"
    '
    Do While ActiveCell.Value <> ""
       strFile = ActiveCell.Value
       If Dir(strPath & strFile) = vbNullString Then
          strCheck = False
       Else
          strCheck = True
       End If
       '
       Sheets("Mensagens Importacao Tabelas").Select
       If strCheck Then
          Sheets("Mensagens Importacao Tabelas").Cells(ActiveCell.Row, ActiveCell.Column).Value = "O arquivo: " & strPath & strFile & " foi encontrado!"
          Transforme_Form.NOM_TABELA = strFile
          Transforme_Form.Repaint
          Call Transforme(strPath, strFile)
          '
       Else
          Sheets("Mensagens Importacao Tabelas").Cells(ActiveCell.Row, ActiveCell.Column).Value = "O arquivo: " & strPath & strFile & " NÃO foi encontrado!"
          'Bloco de ação da Rotina caso o arquivo não exista.
       End If
       Sheets("Mensagens Importacao Tabelas").Select
       Sheets("Mensagens Importacao Tabelas").Cells(ActiveCell.Row + 1, ActiveCell.Column).Select
       Sheets("TRABALHO").Select
       Cells(ActiveCell.Row + 1, ActiveCell.Column).Select
    Loop
    '
    WORK_NAME = ThisWorkbook.Name
    Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
    '
    Workbooks(WORK_NAME).Sheets("TABELAS").Select
    '
    Cells.Replace What:="0", Replacement:="'0", LookAt:=xlWhole, SearchOrder _
    :=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
    '
    Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
    '
    Unload Transforme_Form
    '
End Sub

Sub Transforme(str_PATH As Variant, str_FILE As Variant)
             'NÓ HIERARQUIA II
             'CONTAR_NO_HIERARQUIA
    '
    Dim NUM_REG As Long
    Dim SORT_CELS As Variant
    Dim SORT_KEY As Variant
    Dim TABELA_ADDRESS As Variant
    Dim WORK_NAME As Variant
    Dim address_um As Variant
    Dim address_DOIS As Variant
    Application.ScreenUpdating = False
    '
    WORK_NAME = ThisWorkbook.Name
    '
    Workbooks(WORK_NAME).Sheets("TABELAS").Select
    '
    Select Case str_FILE
           '
           'J_1BTCESTT.txt
           'CEST
           Case "J_1BTCESTT.txt"
                '
                Application.Goto Reference:="CEST_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:= _
                    65001, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=True, Comma:=False, _
                    Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), Array(2, 2)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Columns("B:B").Select
                Selection.Cut
                Columns("A:A").Select
                Selection.Insert Shift:=xlToRight
                '                '
                SORT_CELS = "A1:B" & NUM_REG
                Range(SORT_CELS).Select
                SORT_KEY = "A2:A" & NUM_REG
                '
                ActiveWorkbook.Worksheets("J_1BTCESTT").Sort.SortFields.Clear
                ActiveWorkbook.Worksheets("J_1BTCESTT").Sort.SortFields.Add Key:=Range(SORT_KEY) _
                    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets("J_1BTCESTT").Sort
                    .SetRange Range(SORT_CELS)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                '
                Application.CutCopyMode = False
                '
                SORT_CELS = "A2:B" & NUM_REG
                Range(SORT_CELS).Select
                '
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                '
                CONTAR_CEST_II
                '
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key
                '
           'FORNECEDOR
           Case "LFA1.txt"
                '
                Application.Goto Reference:="CNPJ_FORNEC_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
'
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:= _
                    65001, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, Comma:=False _
                    , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), Array(2, 1), _
                    Array(3, 1), Array(4, 1), Array(5, 2)), TrailingMinusNumbers:=True
                '
                Windows("LFA1.txt").Activate
                Columns("C:C").Select
                Selection.Cut
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Selection.Insert Shift:=xlToRight
                Columns("C:C").Select
                Selection.Cut
                Columns("E:E").Select
                Selection.Insert Shift:=xlToRight
                ActiveWindow.SmallScroll Down:=51
                Columns("B:B").EntireColumn.AutoFit
                '
                SORT_CELS = "A1:E" & NUM_REG
                Range(SORT_CELS).Select
                SORT_KEY = "E1:E" & NUM_REG
                '
                ActiveWorkbook.Worksheets("LFA1").Sort.SortFields.Clear
                ActiveWorkbook.Worksheets("LFA1").Sort.SortFields.Add Key:=Range(SORT_KEY) _
                    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets("LFA1").Sort
                    .SetRange Range(SORT_CELS)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                Application.CutCopyMode = False
                '
                SORT_CELS = "A2:E" & NUM_REG - 1
                Range(SORT_CELS).Select
                '
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                '
                CONTAR_CNPJ_FORNEC
                '
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA

                '
           'CENTROS - REFSITES - LOJAS
           Case "T001W.txt"
                '
                Application.Goto Reference:="PRACAS_TBL_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:= _
                    65001, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                    ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, Comma:=False _
                    , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                SORT_CELS = "A1:D" & NUM_REG
                Range(SORT_CELS).Select
                SORT_KEY = "A1:A" & NUM_REG
                '
                ActiveWorkbook.Worksheets("T001W").Sort.SortFields.Clear
                ActiveWorkbook.Worksheets("T001W").Sort.SortFields.Add Key:=Range(SORT_KEY) _
                    , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets("T001W").Sort
                    .SetRange Range(SORT_CELS)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                '
                Columns("A:A").Select
                '
                Selection.Cut
                Range("D1").Select
                ActiveSheet.Paste
                '
                Application.CutCopyMode = False
                '
                SORT_CELS = "B2:D" & NUM_REG
                Range(SORT_CELS).Select
                '
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
           'GRUPO MERCADORIAS
           Case "T023.txt"
                Application.Goto Reference:="GRP_MERCADORIAS_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                    ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, Comma:=False _
                    , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Columns("B:B").Select
                Selection.Cut
                Columns("A:A").Select
                Selection.Insert Shift:=xlToRight
                '
                Range("A1:B1").Select
                Range(Selection, Selection.End(xlDown)).Select
                Columns("B:B").EntireColumn.AutoFit
                '
                SORT_CELS = "A1:B" & NUM_REG
                Range(SORT_CELS).Select
                SORT_KEY = "A1:A" & NUM_REG
                '
                ActiveWorkbook.Worksheets("T023").Sort.SortFields.Clear
                ActiveWorkbook.Worksheets("T023").Sort.SortFields.Add Key:=Range(SORT_KEY) _
                    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets("T023").Sort
                    .SetRange Range(SORT_CELS)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                '
                SORT_CELS = "A2:C" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                '
                CONTAR_GRP_MERCADORIAS
                '
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
           'COMPRADORES
           Case "T024.txt"
                Application.Goto Reference:="GRP_COMPRADORES_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                    ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, Comma:=False _
                    , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Selection.Cut
                Range("C1").Select
                ActiveSheet.Paste
                Range("B1:C1").Select
                Range(Selection, Selection.End(xlDown)).Select
                Columns("C:C").EntireColumn.AutoFit
                '
                SORT_CELS = "B2:C" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'DESPROTEGE PLANILHA
          'ESTOCAGEM_PROCV
           Case "T142T.txt"
                Application.Goto Reference:="ESTOCAGEM_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, Comma:=False, _
                    Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Selection.Cut
                Range("C1").Select
                ActiveSheet.Paste
                Range("B1:C1").Select
                Range(Selection, Selection.End(xlDown)).Select
                Columns("C:C").EntireColumn.AutoFit
                '
                SORT_CELS = "B2:C" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
        'T179T.txt
        'FAM_INTER_PROCV
           Case "T179T.txt"
                Application.Goto Reference:="FAM_INTER_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                    ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, Comma:=False _
                    , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Range("A2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.NumberFormat = "0"
                Columns("A:A").EntireColumn.AutoFit
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Selection.Cut
                Range("C1").Select
                ActiveSheet.Paste
                Range("B1:C1").Select
                Range(Selection, Selection.End(xlDown)).Select
                Columns("C:C").EntireColumn.AutoFit
                '
                SORT_CELS = "B1:C" & NUM_REG
                Range(SORT_CELS).Select
                SORT_KEY = "B1:B" & NUM_REG
                '
                ActiveWorkbook.Worksheets("T179T").Sort.SortFields.Clear
                ActiveWorkbook.Worksheets("T179T").Sort.SortFields.Add Key:=Range(SORT_KEY) _
                    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets("T179T").Sort
                    .SetRange Range(SORT_CELS)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                '
                SORT_CELS = "B2:C" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
        'T604N.txt
        'NCM
           Case "T604N.txt"
                Application.Goto Reference:="NCM_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=True, Comma:=False, _
                    Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 2), Array( _
                    3, 1)), TrailingMinusNumbers:=True
                '
                Columns("C:C").Select
                Selection.Cut
                Range("A1").Select
                ActiveSheet.Paste
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                SORT_CELS = "A1:B" & NUM_REG
                Range(SORT_CELS).Select
                SORT_KEY = "A1:A" & NUM_REG
                '
                ActiveWorkbook.Worksheets("T604N").Sort.SortFields.Clear
                ActiveWorkbook.Worksheets("T604N").Sort.SortFields.Add Key:=Range(SORT_KEY), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets("T604N").Sort
                    .SetRange Range(SORT_CELS)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                '
                SORT_CELS = "A2:B" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                '
                CONTAR_NCM_II
                '
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
     'TMFPF.txt
     'Perfil distribuicao
           Case "TMFPF.txt"
                Application.Goto Reference:="FLUXO_TBL_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=True, Comma:=False, _
                    Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Columns("B:B").Select
                Selection.Cut
                Columns("A:A").Select
                Selection.Insert Shift:=xlToRight
                '
                SORT_CELS = "A2:B" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
      'TWSST.txt
      'GAMA_PROCV
           Case "TWSST.txt"
                Application.Goto Reference:="GAMA_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=True, Comma:=False, _
                    Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Columns("B:B").Select
                Selection.Cut
                Columns("A:A").Select
                Selection.Insert Shift:=xlToRight
                '
                SORT_CELS = "A2:B" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
        'TWSVT.txt
        'GARANTIA_PROCV
           Case "TWSVT.txt"
                Application.Goto Reference:="GARANTIA_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=True, Comma:=False, _
                    Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Columns("B:B").Select
                Selection.Cut
                Columns("A:A").Select
                Selection.Insert Shift:=xlToRight
                '
                SORT_CELS = "A2:B" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
           'Composicao Material
           Case "WRF_FIBER_CODES.txt"
                Application.Goto Reference:="COMPOSICAO_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=True, Comma:=False, _
                    Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Columns("B:B").Select
                Selection.Cut
                Columns("A:A").Select
                Selection.Insert Shift:=xlToRight
                '
                SORT_CELS = "A2:B" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
    'FORNECEDOR VIRTUAL
    'FORNEC_VIRTUA_SAP
           Case "WYT3.txt"
                Application.Goto Reference:="FORNEC_VIRTUA_SAP"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, Comma:=False _
                    , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), Array(2, 1), _
                    Array(3, 1), Array(4, 1), Array(5, 1)), TrailingMinusNumbers:=True
                '
                Columns("A:A").EntireColumn.AutoFit
                '
                Columns("B:B").Select
                Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                SORT_CELS = "A1:F" & NUM_REG
                Range(SORT_CELS).Select
                SORT_KEY = "A1:A" & NUM_REG
                '
                ActiveWorkbook.Worksheets("WYT3").Sort.SortFields.Clear
                ActiveWorkbook.Worksheets("WYT3").Sort.SortFields.Add Key:=Range(SORT_KEY), _
                    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                '
                SORT_KEY = "F1:F" & NUM_REG
                '
                ActiveWorkbook.Worksheets("WYT3").Sort.SortFields.Add Key:=Range(SORT_KEY), _
                    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets("WYT3").Sort
                    .SetRange Range(SORT_CELS)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                '
                SORT_CELS = "A2:F" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
           'CODANP
           Case "ZTMMC_COD_ANP.txt"
                Application.Goto Reference:="CODANP_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=True, Comma:=False, _
                    Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").EntireColumn.AutoFit
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Columns("B:B").Select
                Selection.Cut
                Columns("A:A").Select
                Selection.Insert Shift:=xlToRight
                '
                SORT_CELS = "A2:B" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
    'ZTMMD_CLASSE.txt
    'CLASSE_PROCV
           Case "ZTMMD_CLASSE.txt"
                Application.Goto Reference:="CLASSE_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=True, Comma:=False, _
                    Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Columns("B:B").Select
                Selection.Cut
                Columns("A:A").Select
                Selection.Insert Shift:=xlToRight
                '
                SORT_CELS = "A2:B" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
    'ZTMMD_ECOSUS.txt
    'ECO_SUSTENTA_PROCV
           Case "ZTMMD_ECOSUS.txt"
                Application.Goto Reference:="ECO_SUSTENTA_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=True, Comma:=False, _
                    Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Columns("B:B").Select
                Selection.Cut
                Columns("A:A").Select
                Selection.Insert Shift:=xlToRight
                '
                SORT_CELS = "A2:B" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
    'ZTMMD_IBAMA_PASC.txt
    'PRINCIPIO_ATIVO_PROCV
            Case "ZTMMD_IBAMA_PASC.txt"
                Application.Goto Reference:="PRINCIPIO_ATIVO_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=True, Comma:=False, _
                    Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Columns("B:B").Select
                Selection.Cut
                Columns("A:A").Select
                Selection.Insert Shift:=xlToRight
                '
                SORT_CELS = "A1:B" & NUM_REG
                Range(SORT_CELS).Select
                SORT_KEY = "A1:A" & NUM_REG
                '
                ActiveWorkbook.Worksheets("ZTMMD_IBAMA_PASC").Sort.SortFields.Clear
                ActiveWorkbook.Worksheets("ZTMMD_IBAMA_PASC").Sort.SortFields.Add Key:=Range(SORT_KEY), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets("ZTMMD_IBAMA_PASC").Sort
                    .SetRange Range(SORT_CELS)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                '
                SORT_CELS = "A2:B" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
    'ZTMMD_POL_CIVIL.txt
    'PRINCIPIO_ATV_CIVIL_PROCV
            Case "ZTMMD_POL_CIVIL.txt"
                Application.Goto Reference:="PRINCIPIO_ATV_CIVIL_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=True, Comma:=False, _
                    Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Columns("B:B").Select
                Selection.Cut
                Columns("A:A").Select
                Selection.Insert Shift:=xlToRight
                '
                SORT_CELS = "A1:B" & NUM_REG
                Range(SORT_CELS).Select
                SORT_KEY = "A1:A" & NUM_REG
                '
                ActiveWorkbook.Worksheets("ZTMMD_POL_CIVIL").Sort.SortFields.Clear
                ActiveWorkbook.Worksheets("ZTMMD_POL_CIVIL").Sort.SortFields.Add Key:=Range(SORT_KEY), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets("ZTMMD_POL_CIVIL").Sort
                    .SetRange Range(SORT_CELS)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                '
                SORT_CELS = "A2:B" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
'ZTMMD_POL_FEDER.txt
'PRINCIPIO_ATV_FEDERAL_PROCV
            Case "ZTMMD_POL_FEDER.txt"
                Application.Goto Reference:="PRINCIPIO_ATV_FEDERAL_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=True, Comma:=False, _
                    Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Columns("B:B").Select
                Selection.Cut
                Columns("A:A").Select
                Selection.Insert Shift:=xlToRight
                '
                SORT_CELS = "A1:B" & NUM_REG
                Range(SORT_CELS).Select
                SORT_KEY = "A1:A" & NUM_REG
                '
                ActiveWorkbook.Worksheets("ZTMMD_POL_FEDER").Sort.SortFields.Clear
                ActiveWorkbook.Worksheets("ZTMMD_POL_FEDER").Sort.SortFields.Add Key:=Range(SORT_KEY), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets("ZTMMD_POL_FEDER").Sort
                    .SetRange Range(SORT_CELS)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                '
                SORT_CELS = "A2:B" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
    'ZTMMD_SUBCLASSE.txt
    'SUBCLASSE_PROCV
            Case "ZTMMD_SUBCLASSE.txt"
                Application.Goto Reference:="SUBCLASSE_PROCV"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin:=65001 _
                    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlNone, _
                    ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=True, Comma:=False, _
                    Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                Columns("B:B").Select
                Selection.Cut
                Columns("A:A").Select
                Selection.Insert Shift:=xlToRight
                '
                SORT_CELS = "A1:B" & NUM_REG
                Range(SORT_CELS).Select
                SORT_KEY = "A1:A" & NUM_REG
                '
                ActiveWorkbook.Worksheets("ZTMMD_SUBCLASSE").Sort.SortFields.Clear
                ActiveWorkbook.Worksheets("ZTMMD_SUBCLASSE").Sort.SortFields.Add Key:=Range(SORT_KEY), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets("ZTMMD_SUBCLASSE").Sort
                    .SetRange Range(SORT_CELS)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                '
                SORT_CELS = "A2:B" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
    'NO DE HIERARQUIA
    'WRF_MATGRP_STRUC.txt
           Case "WRF_MATGRP_STRUC.txt"
                Application.Goto Reference:="FORNEC_VIRTUA_SAP"
                Workbooks(WORK_NAME).Sheets("TABELAS").Unprotect Key 'DESPROTEGE PLANILHA
                Selection.Clear
                '
                ActiveCell.Select
                TABELA_ADDRESS = ActiveCell.Address
                '
                Workbooks.OpenText Filename:= _
                    str_PATH & str_FILE, Origin _
                    :=65001, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
                    xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, _
                    Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
                    Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1)), TrailingMinusNumbers:=True
                            Columns("A:A").EntireColumn.AutoFit
                '
                Columns("E:E").Select
                Selection.Cut
                Columns("A:A").Select
                Selection.Insert Shift:=xlToRight
                Range("A1:E1").Select
                Range(Selection, Selection.End(xlDown)).Select
                '
                Columns("A:A").Select
                '
                NUM_REG = Application.CountA(Range("A:A"))
                '
                SORT_CELS = "A1:E" & NUM_REG
                Range(SORT_CELS).Select
                SORT_KEY = "B1:B" & NUM_REG
                '
                ActiveWorkbook.Worksheets("WRF_MATGRP_STRUC").Sort.SortFields.Clear
                ActiveWorkbook.Worksheets("WRF_MATGRP_STRUC").Sort.SortFields.Add Key:=Range _
                    (SORT_KEY), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                    xlSortNormal
                With ActiveWorkbook.Worksheets("WRF_MATGRP_STRUC").Sort
                    .SetRange Range(SORT_CELS)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                '
                SORT_CELS = "A2:B" & NUM_REG
                Range(SORT_CELS).Select
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                Workbooks(WORK_NAME).Sheets("TABELAS").Protect Key 'PROTEGE PLANILHA
                '
                
    End Select
    ThisWorkbook.Sheets("TABELAS").Protect Key
End Sub

Sub ESPECIE_MADEIRA(NOME_ARQ As Variant)
                Dim NUM_REG As Long
                Dim SORT_CELS As Variant
                Dim RANGE_REGS As Variant
                Dim NOME_PLANILHA As Variant
                Dim TABELA_ADDRESS As Variant
                Dim WORK_NAME As String
                Dim CSV_WORK_NAME As String
                '
                On Error GoTo ERR1
                '
                WORK_NAME = ThisWorkbook.Name
                '
                Workbooks.Open Filename:=NOME_ARQ
                '
                CSV_WORK_NAME = ThisWorkbook.Name
                '
                'Workbooks.OpenText Filename:= _
                '    NOME_ARQ, Origin:=65001 _
                '    , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                '    ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, Comma:=False _
                '    , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                '    TrailingMinusNumbers:=True
                '
                Columns("A:A").Select
                NUM_REG = Application.CountA(Range("A:A"))
                Selection.Cut Destination:=Columns("C:C")
                Range("B1:C1").Select
                Range(Selection, Selection.End(xlDown)).Select
                '
                ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
                ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range( _
                    "B1:B200"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                    xlSortNormal
                SORT_CELS = "B1:C" & NUM_REG
                With ActiveWorkbook.ActiveSheet.Sort
                    .SetRange Range(SORT_CELS)
                    .Header = xlGuess
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                ActiveWindow.SmallScroll Down:=-24
                '
                Selection.Copy Workbooks(WORK_NAME).Worksheets("TABELAS").Range(TABELA_ADDRESS)
                ActiveWorkbook.Close SaveChanges:=False
                '
                Sheets("Mensagens Importacao Tabelas").Select
                Sheets("Mensagens Importacao Tabelas").Cells(ActiveCell.Row + 1, ActiveCell.Column).Select
                Sheets("Mensagens Importacao Tabelas").Cells(ActiveCell.Row, ActiveCell.Column).Value = "Numero Registros ESPECIE_MADEIRA: " & NUM_REG
                '
                Sheets("Mensagens Importacao Tabelas").Cells(ActiveCell.Row + 1, ActiveCell.Column).Select
                Sheets("Mensagens Importacao Tabelas").Cells(ActiveCell.Row, ActiveCell.Column).Value = "SUCESSO ESPECIE_MADEIRA"
                '
                Exit Sub
ERR1:
    Workbook(WOROK_NAME).Select
    ThisWorkbook.Sheets("Mensagens Importacao Tabelas").Select
    ThisWorkbook.Sheets("Mensagens Importacao Tabelas").Cells(ActiveCell.Row + 1, ActiveCell.Column).Select
    ThisWorkbook.Sheets("Mensagens Importacao Tabelas").Cells(ActiveCell.Row, ActiveCell.Column).Value = Err.Number & " - " & _
    Err.Description & " - " & Err.Source
'
'Number Fornece Numero do erro gerado
'Description Fornece a descrição do erro.
'Source Identifica o nome do objeto que gerou o erro
'
End Sub

'CNPJ FORNECEDOR
Public Sub CONTAR_CNPJ_FORNEC()
    Dim DESCR_1 As String
    Dim COD_1 As String
    Dim LIN_BUSCA As Long
    Dim COL_BUSCA As Long
    Dim LIN_BUSCA_II As Long
    Dim COL_BUSCA_II As Long
    '
    Sheets("TABELAS").Select
    '
    Application.Goto Reference:="CNPJ_BUSCA_II"
    '
    COL_BUSCA_II = ActiveCell.Column
    '
    Application.Goto Reference:="CNPJ_FORNEC_PROCV"
    '
    LIN_BUSCA = ActiveCell.Row
    COL_BUSCA = ActiveCell.Column
    '
    Do While ActiveCell.Value <> ""
       Transforme_Form.NUM_REGISTRO = LIN_BUSCA
       Transforme_Form.Repaint
       'ALTERA TABELA PRINCIPAL
       DESCR_1 = Cells(LIN_BUSCA, COL_BUSCA).Value
       Cells(LIN_BUSCA, COL_BUSCA).Value = Cells(LIN_BUSCA, COL_BUSCA).Value & _
             " - " & CDbl(Cells(LIN_BUSCA, COL_BUSCA + 4).Value)
       '
       'CONSTROE TABELA II
       '
       'Cells(LIN_BUSCA, COL_BUSCA_II).Value = CDbl(Cells(LIN_BUSCA, COL_BUSCA + 4).Value)
       'Cells(LIN_BUSCA, COL_BUSCA_II).Value = Cells(LIN_BUSCA, COL_BUSCA_II).Value & " - " & DESCR_1
       'Cells(LIN_BUSCA, COL_BUSCA_II + 1).Value = (Cells(LIN_BUSCA, COL_BUSCA + 4).Value)
        
       LIN_BUSCA = LIN_BUSCA + 1
       Cells(LIN_BUSCA, COL_BUSCA).Select
    Loop
End Sub

'GRUPO MERCADORIAS II
Public Sub CONTAR_GRP_MERCADORIAS()
    Dim DESCR_1 As String
    Dim COD_1 As String
    Dim LIN_BUSCA As Long
    Dim COL_BUSCA As Long
    Dim LIN_BUSCA_II As Long
    Dim COL_BUSCA_II As Long
    '
    Sheets("TABELAS").Select
    '
    Application.Goto Reference:="GRP_MERCADORIAS_II"
    '
    COL_BUSCA_II = ActiveCell.Column
    '
    Application.Goto Reference:="GRP_MERCADORIAS_PROCV"
    '
    LIN_BUSCA = ActiveCell.Row
    COL_BUSCA = ActiveCell.Column
    '
    Do While ActiveCell.Value <> ""
       Transforme_Form.NUM_REGISTRO = LIN_BUSCA
       Transforme_Form.Repaint
       'ALTERA TABELA PRINCIPAL
       If IsNumeric(Cells(LIN_BUSCA, COL_BUSCA + 1).Value) Then
          DESCR_1 = Cells(LIN_BUSCA, COL_BUSCA).Value
          Cells(LIN_BUSCA, COL_BUSCA).Value = Cells(LIN_BUSCA, COL_BUSCA).Value & _
             " - " & CDbl(Cells(LIN_BUSCA, COL_BUSCA + 1).Value)
          '
          'CONSTROE TABELA II
          '
          'Cells(LIN_BUSCA, COL_BUSCA_II).Value = CDbl(Cells(LIN_BUSCA, COL_BUSCA + 1).Value)
          'Cells(LIN_BUSCA, COL_BUSCA_II).Value = Cells(LIN_BUSCA, COL_BUSCA_II).Value & " - " & DESCR_1
          'Cells(LIN_BUSCA, COL_BUSCA_II + 1).Value = (Cells(LIN_BUSCA, COL_BUSCA + 1).Value)
       End If
       LIN_BUSCA = LIN_BUSCA + 1
       Cells(LIN_BUSCA, COL_BUSCA).Select
    Loop
End Sub

'NÓ HIERARQUIA II
Public Sub CONTAR_NO_HIERARQUIA()
    Dim DESCR_1 As String
    Dim COD_1 As String
    Dim LIN_BUSCA As Long
    Dim COL_BUSCA As Long
    Dim LIN_BUSCA_II As Long
    Dim COL_BUSCA_II As Long
    '
    Sheets("TABELAS").Select
    ThisWorkbook.Sheets("TABELAS").Unprotect Key
    '
    Application.Goto Reference:="NO_HIERARQ_PROCV_II"
    '
    COL_BUSCA_II = ActiveCell.Column
    '
    Application.Goto Reference:="NO_HIERARQ_PROCV"
    '
    LIN_BUSCA = ActiveCell.Row
    COL_BUSCA = ActiveCell.Column
    '
    Do While ActiveCell.Value <> ""
       'ALTERA TABELA PRINCIPAL
       If IsNumeric(Cells(LIN_BUSCA, COL_BUSCA + 1).Value) Then
          DESCR_1 = Cells(LIN_BUSCA, COL_BUSCA).Value
          Cells(LIN_BUSCA, COL_BUSCA).Value = Cells(LIN_BUSCA, COL_BUSCA).Value & _
             " - " & CDbl(Cells(LIN_BUSCA, COL_BUSCA + 1).Value)
          '
          'CONSTROE TABELA II
          '
          'Cells(LIN_BUSCA, COL_BUSCA_II).Value = CDbl(Cells(LIN_BUSCA, COL_BUSCA + 1).Value)
          'Cells(LIN_BUSCA, COL_BUSCA_II).Value = Cells(LIN_BUSCA, COL_BUSCA_II).Value & " - " & DESCR_1
          'Cells(LIN_BUSCA, COL_BUSCA_II + 1).Value = (Cells(LIN_BUSCA, COL_BUSCA + 1).Value)
       End If
       LIN_BUSCA = LIN_BUSCA + 1
       Cells(LIN_BUSCA, COL_BUSCA).Select
    Loop
    
End Sub

'NCM II
Public Sub CONTAR_NCM_II()
    Dim DESCR_1 As String
    Dim COD_1 As String
    Dim LIN_BUSCA As Long
    Dim COL_BUSCA As Long
    Dim LIN_BUSCA_II As Long
    Dim COL_BUSCA_II As Long
    '
    Sheets("TABELAS").Select
    '
    Application.Goto Reference:="NCM_PROCV_II"
    '
    COL_BUSCA_II = ActiveCell.Column
    '
    Application.Goto Reference:="NCM_PROCV"
    '
    LIN_BUSCA = ActiveCell.Row
    COL_BUSCA = ActiveCell.Column
    '
    Do While ActiveCell.Value <> ""
       Transforme_Form.NUM_REGISTRO = LIN_BUSCA
       Transforme_Form.Repaint
       'ALTERA TABELA PRINCIPAL
        DESCR_1 = Cells(LIN_BUSCA, COL_BUSCA).Value
        Cells(LIN_BUSCA, COL_BUSCA).Value = Cells(LIN_BUSCA, COL_BUSCA).Value & _
        " - " & Cells(LIN_BUSCA, COL_BUSCA + 1).Value
        '
        'CONSTROE TABELA II
        '
        'Cells(LIN_BUSCA, COL_BUSCA_II).Value = Cells(LIN_BUSCA, COL_BUSCA + 1).Value
        'Cells(LIN_BUSCA, COL_BUSCA_II).Value = Cells(LIN_BUSCA, COL_BUSCA_II).Value & " - " & DESCR_1
        'Cells(LIN_BUSCA, COL_BUSCA_II + 1).Value = (Cells(LIN_BUSCA, COL_BUSCA + 1).Value)
        '
       LIN_BUSCA = LIN_BUSCA + 1
       Cells(LIN_BUSCA, COL_BUSCA).Select
    Loop
End Sub

'CEST II
Public Sub CONTAR_CEST_II()
    Dim DESCR_1 As String
    Dim COD_1 As String
    Dim LIN_BUSCA As Long
    Dim COL_BUSCA As Long
    Dim LIN_BUSCA_II As Long
    Dim COL_BUSCA_II As Long
    '
    Sheets("TABELAS").Select
    '
    Application.Goto Reference:="CEST_PROCV_II"
    '
    COL_BUSCA_II = ActiveCell.Column
    '
    Application.Goto Reference:="CEST_PROCV"
    '
    LIN_BUSCA = ActiveCell.Row
    COL_BUSCA = ActiveCell.Column
    '
    Do While ActiveCell.Value <> ""
       Transforme_Form.NUM_REGISTRO = LIN_BUSCA
       Transforme_Form.Repaint
       'ALTERA TABELA PRINCIPAL
        DESCR_1 = Cells(LIN_BUSCA, COL_BUSCA).Value
        Cells(LIN_BUSCA, COL_BUSCA).Value = Cells(LIN_BUSCA, COL_BUSCA).Value & _
        " - " & Cells(LIN_BUSCA, COL_BUSCA + 1).Value
        '
        'CONSTROE TABELA II
        '
        'Cells(LIN_BUSCA, COL_BUSCA_II).Value = Cells(LIN_BUSCA, COL_BUSCA + 1).Value
        'Cells(LIN_BUSCA, COL_BUSCA_II).Value = Cells(LIN_BUSCA, COL_BUSCA_II).Value & " - " & DESCR_1
        'Cells(LIN_BUSCA, COL_BUSCA_II + 1).Value = (Cells(LIN_BUSCA, COL_BUSCA + 1).Value)
        '
       LIN_BUSCA = LIN_BUSCA + 1
       Cells(LIN_BUSCA, COL_BUSCA).Select
    Loop
End Sub



Function IsLoaded(ByVal strFormName As String) As Integer
' Retorna True se o formulário especificado estiver aberto
' no modo Formulário ou no modo Folha de Dados.
Const conObjStateClosed = 0
Const conDesignView = 0

If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
   If Forms(strFormName).CurrentView <> conDesignView Then
      IsLoaded = True
   End If
End If
'
End Function


Function VALIDACAO_COD_EAN_UNICO() As Boolean
    '
    On Error Resume Next
    Dim SORT_ADDRESS As Range
    Dim lin As Integer
    Dim col As Integer
    Dim COD_EAN_ANT As Variant
    Dim CELL_ADRESS As String
    '
    Dim ULT_LIN_ERRO As Integer
    '
    Dim LINHAS_ERRO As Integer
    '
    Application.ScreenUpdating = False
    '
    Sheets("ERROS Cadastrais").Select
    Cells(1, 1).Select ' CELULA A1
    ULT_LIN_ERRO = 1
    Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
       ULT_LIN_ERRO = ULT_LIN_ERRO + 1
    Loop
    '
    LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ThisWorkbook.Unprotect Key
    '
    If Sheets("TRABALHO").Visible = False Then
       Sheets("TRABALHO").Visible = True
    End If
    '
    Sheets("TRABALHO").Select
    ActiveSheet.Unprotect Key
    Range("E51:E251").ClearContents
    '
    Sheets("Dados Cadastrais").Select
    '
    Application.Goto Reference:="COD_EAN_UNICO"
    '
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    '
    Sheets("TRABALHO").Select
    Application.Goto Reference:="COD_EAN_UNICO_TRAB"
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    '    :=True, Transpose:=False
    ActiveSheet.Paste
    '
    ActiveWorkbook.Worksheets("TRABALHO").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TRABALHO").Sort.SortFields.Add Key:=Range( _
        "E49:E251"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("TRABALHO").Sort
        .SetRange Range("E50:E251")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    '
    Application.Goto Reference:="COD_EAN_UNICO_TRAB"
    lin = ActiveCell.Row + 1
    col = ActiveCell.Column
    '
    COD_EAN_ANT = Cells(lin, col).Value
    '
    Do While Cells(lin, col).Value <> ""
       lin = lin + 1
       If Cells(lin, col).Value = COD_EAN_ANT Then
          'CELL_ADRESS = ActiveCell.Address
          MsgBox ("CODIGOS EANS DEVEM SER UNICOS. O Codigo Ean " & COD_EAN_ANT & " Esta duplicado.")
          Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 1).Value = "CODIGOS EANS DEVEM SER UNICOS. O Codigo Ean " & COD_EAN_ANT & " esta duplicado."
          Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 2).Value = "Iniciais"
          Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 3).Value = "Dados Iniciais"
          'Sheets("ERROS Cadastrais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
          '
          'CODIGO EAN = "INTERNO" DUPLICADO PODE. AVISA MAS NAO DEIXA BOTAO VERMELHO
          '
          If UCase(COD_EAN_ANT) <> "INTERNO" Then
             Sheets("Inicio").CADASTRAIS.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").INICIO.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Cadastrais").INICIO.ForeColor = &H8000000E  'BRANCO
          End If
          VALIDACAO_COD_EAN_UNICO = False
          Exit Function
       Else
          COD_EAN_ANT = Cells(lin, col).Value
          VALIDACAO_COD_EAN_UNICO = True
       End If
   Loop
   '
   On Error GoTo 0
End Function

Function VALIDACAO_SUBSORT_UNICO() As Boolean
    '
    Dim SORT_ADDRESS As Range
    Dim lin As Integer
    Dim col As Integer
    Dim CNPJ_ANT As Variant
    Dim SUBSORT_ANT As Variant
    Dim CELL_ADRESS As String
    '
    Dim ULT_LIN_ERRO As Integer
    '
    Dim LINHAS_ERRO As Integer
    '
    Application.ScreenUpdating = False
    '
    Sheets("ERROS Dados Fiscais").Select
    Cells(1, 1).Select ' CELULA A1
    ULT_LIN_ERRO = 1
    Do While Not IsEmpty(Cells(ULT_LIN_ERRO, 1))
       ULT_LIN_ERRO = ULT_LIN_ERRO + 1
    Loop
    '
    LINHAS_ERRO = ULT_LIN_ERRO 'PROXIMA LINHA NAO UTILIZADA
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ThisWorkbook.Unprotect Key
    '
    If Sheets("TRABALHO").Visible = False Then
       Sheets("TRABALHO").Visible = True
    End If
    '
    Sheets("TRABALHO").Select
    ActiveSheet.Unprotect Key
    Range("E260:E357").ClearContents
    '
    Sheets("Escolha Perfil Distribuicao").Select
    '
    Application.Goto Reference:="SUBSORT_PERFIL"
    '
    Selection.Copy
    '
    Sheets("TRABALHO").Select
    Application.Goto Reference:="SUBSORT_TRAB"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("TRABALHO").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TRABALHO").Sort.SortFields.Add Key:=Range( _
        "E261:E356"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    ActiveWorkbook.Worksheets("TRABALHO").Sort.SortFields.Add Key:=Range( _
        "H261:H356"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("TRABALHO").Sort
        .SetRange Range("E260:H356")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    '
    Application.Goto Reference:="SUBSORT_TRAB"
    lin = ActiveCell.Row + 1
    col = ActiveCell.Column
    '
    CNPJ_ANT = Cells(lin, col).Value 'CNPJ
    SUBSORT_ANT = Cells(lin, col + 3).Value 'SUBSORTIMENTO
    '
    Do While Cells(lin, col).Value <> ""
       '
       lin = lin + 1
       '
       If Cells(lin, col).Value = CNPJ_ANT Then
          If Cells(lin, 8).Value <> SUBSORT_ANT And Trim(Cells(lin, 8).Value) <> "" Then
             CELL_ADRESS = ActiveCell.Address
             MsgBox ("SUBSORTIMENTO DEVE SER UNICO POR CNPJ. O CNPJ " & CNPJ_ANT & " TEM MAIS DE UM SUBSORTIMENTO.")
             Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 1).Value = "SUBSORTIMENTO DEVE SER UNICO POR CNPJ. O CNPJ " & CNPJ_ANT & " TEM MAIS DE UM SUBSORTIMENTO."
             Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 2).Value = "Fiscais"
             Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 3).Value = "Dados Fiscais"
             Sheets("ERROS Dados Fiscais").Cells(LINHAS_ERRO, 4).Value = CELL_ADRESS
             Sheets("Inicio").FISCAIS.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Fiscais").NAO_VALIDADO.Visible = True 'ESCONDE BOTAO
             Sheets("Dados Fiscais").NAO_VALIDADO.Caption = "COM  ERRO NA VALIDACAO"
             Sheets("Dados Fiscais").NAO_VALIDADO.BackColor = &HFF& 'VERMELHO
             Sheets("Dados Fiscais").NAO_VALIDADO.ForeColor = &H8000000E  'BRANCO
             VALIDACAO_SUBSORT_UNICO = False
             Exit Function
          Else
            CNPJ_ANT = Cells(lin, col).Value 'CNPJ
            SUBSORT_ANT = Cells(lin, col + 3).Value 'SUBSORTIMENTO
          End If
       Else
         CNPJ_ANT = Cells(lin, col).Value 'CNPJ
         VALIDACAO_SUBSORT_UNICO = True
       End If
       '
   Loop
   '
End Function
