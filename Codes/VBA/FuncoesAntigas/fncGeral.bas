Attribute VB_Name = "fncGeral"
'================================================================================================= ===================
'M�dulo fun��o geral do painel de controle
'Obs.: Todo desenvolvimento de fun��es, mensagens e afins para o painel de controle de forma geral
'verifica prazo 
'
'================================================================================================= ===================

Option Explicit

Public Const MSG_INFORMACAO1 As String = "Procedimento realizado. Sem pend�ncias"
Public Const MSG_INFORMACAO2 As String = "Procedimento realizado. Verificar coluna LOG"

'Campos do painel de controle
Public Const DADOS_CAMPOB As String = "Situa��o"
Public Const DADOS_CAMPOC As String = "Empresa"
Public Const DADOS_CAMPOD As String = "Descri��o da atividade"
Public Const DADOS_CAMPOE As String = "Item"
Public Const DADOS_CAMPOF As String = "Aplica��o"
Public Const DADOS_CAMPOG As String = "C�digo da solicita��o"
Public Const DADOS_CAMPOH As String = "Chave solicitante"
Public Const DADOS_CAMPOI As String = "D. solicitante"
Public Const DADOS_CAMPOJ As String = "Ger�ncia solicitante"
Public Const DADOS_CAMPOK As String = "Qtd."
Public Const DADOS_CAMPOL As String = "C�digo OS"
Public Const DADOS_CAMPOM As String = "D. abertura"
Public Const DADOS_CAMPON As String = "D. fechamento"
Public Const DADOS_CAMPOO As String = "Prazo combinado"
Public Const DADOS_CAMPOP As String = "Qtd."
Public Const DADOS_CAMPOQ As String = "Agrupamento"
Public Const DADOS_CAMPOR As String = "Classifica��o"
Public Const DADOS_CAMPOS As String = "FDM"
Public Const DADOS_CAMPOT As String = "Obs. sigla"
Public Const DADOS_CAMPOU As String = "FA"
Public Const DADOS_CAMPOV As String = "Localidade"
Public Const DADOS_CAMPOW As String = "Munic�pio"
Public Const DADOS_CAMPOX As String = "UF"
Public Const DADOS_CAMPOY As String = "Localidade"
Public Const DADOS_CAMPOZ As String = "Munic�pio"
Public Const DADOS_CAMPOAA As String = "UF"


Function gerar_periodo(iPerMes As Integer, iPerAno As Integer) As String
'Fun��o para gerar per�odo entre datas

Dim iMesAnt As Integer
Dim iAnoAnt As Integer
Dim sPerAnt As String

iMesAnt = iPerMes - 1
If iMesAnt = 0 Then
  iMesAnt = 12
  iAnoAnt = iPerAno - 1
  sPerAnt = 26 & "/" & iMesAnt & "/" & iAnoAnt
ElseIf iMesAnt > 0 Then
  sPerAnt = 26 & "/" & iMesAnt & "/" & iPerAno
End If

gerar_periodo = sPerAnt & "  �  " & 25 & "/" & iPerMes & "/" & iPerAno

End Function


Function converter_numParaLetra(iColNum As Integer) As String
'Fun��o para converter n�mero para letra em colunas

Dim vArr As Variant

vArr = Split(Cells(1, iColNum).Address(True, False), "$")

converter_numParaLetra = vArr(0)
    
End Function

Function informar_coluna(nValor As String) As String
'Fun��o para informar coluna no resultado de estilo L1C1

Dim nResultado As String
Dim i As Integer

'Para cada caractere na vari�vel, copia d�gitos para retornar string
For i = 1 To Len(nValor)
    If Mid(nValor, i, 1) >= "0" And Mid(nValor, i, 1) <= "9" Then
        nResultado = nResultado + Mid(nValor, i, 1)
    End If
Next i

informar_coluna = nResultado

End Function

Function ExibirMsg(Msg As String) As String
'Fun��o para descrever a mensagem no campo LOG referente ao tipo de pend�ncia

Select Case Msg
  'Quando n�o possui valor de resultado
  Case "ValorNaoEncontrado"
     ExibirMsg = "Descri��o da pesquisa n�o encontrada"
     
  'Quando alguma c�lula estiver vazia
  'Ser� utilizado para as c�lulas vazias no painel de controle
  Case "CampoVazio"
     ExibirMsg = "Campo vazio"
     
  'Quando o grupo de munic�pio n�o permite servi�o na atividade
  Case "FDMNaoPertimido"
     ExibirMsg = "N�o foi permitido calcular FDM"
     
  'Quando a descri��o inserida na c�lula n�o se econtra na planilha tabelaA
  Case "descricaoTabelaA"
     ExibirMsg = "Descri��o n�o encontrada na planilha tabelaA"
     
  'Quando o item n�o est� vinculado a atividade
  Case "itemXAtividade"
     ExibirMsg = "Campo item n�o pertence ao campo atividade"
     
  'Quando chave de identifica��o possui quantidade de caracteres diferente do padr�o de usu�rio
  Case "chaveUsuario"
     ExibirMsg = "Padr�o de caracteres divergente para chave de usu�rio"
     
  'Quando data inserida n�o foi considerada uma data v�lida
  Case "dataInvalida"
     ExibirMsg = "Data inserida inv�lida"
     
  'Quando texto inserido possui algum caractere inv�lido para aquele texto
  Case "textoInvalido"
     ExibirMsg = "Caractere inv�lido para texto inserido: "

  'Quando usu�rio inseriu texto no lugar de valores
  Case "valorInvalido"
     ExibirMsg = "Caractere texto em c�lula num�rica"

  'Quando usu�rio n�o inseriu prazo obrigat�rio para medir FDM
  Case "atividadeSemPrazo"
     ExibirMsg = "Prazo obrigat�rio para atividade informada"

  'Quando dado inserido no campo n�o � obrigat�rio para a atividade
  Case "dadoNaoObrigatorio"
     ExibirMsg = "Dado n�o obrigat�rio"

  'Quando dado inserido no campo � obrigat�rio para a atividade
  Case "dadoObrigatorio"
     ExibirMsg = "Dado obrigat�rio"

  'Quando a pasta est� vazia
  Case "pastaVazia"
     ExibirMsg = "Pasta vazia para diret�rio informado"

  'Quando a descri��o de centro n�o foi inserida para Km adicional em orige/destino
  Case "origemDestino"
     ExibirMsg = "Centro n�o localizado para origem X destino"
     
  'Quando o munic�pio inserido n�o foi localizado na planilha calend�rio
  Case "CalendarioMunicipio"
     ExibirMsg = "Munic�pio n�o encontrado na planilha calend�rio"
     
  'Quando o contrato est� divergente com o estado
  Case "contratoUF"
     ExibirMsg = "Empresa/Contrato divergente com localidade/UF"
     
  'Quando atividade expresso n�o pode ser atendido em grupo espec�fico ...
  '... conforme prazos espec�ficos no contrato de guarda externa.
  'Obs.: Nestes casos, somente atividade NORMAL poder� atender
  Case "atividadeLocalidade"
     ExibirMsg = "Atividade EXPRESSO n�o atendido para localidade informada"

  'Quando munic�pio n�o foi localizado no grupo
  'Neste caso, acrescent�-lo ao grupo (A, B, C ou D) conforme contrato de guarda externa
  Case "municipioSemGrupo"
     ExibirMsg = "Munic�pio n�o encontrado no grupo de munic�pios"
     
  'Quando o munic�pio n�o est� vinculado a UF
  Case "municipioXUF"
     ExibirMsg = "Munic�pio n�o pertence a UF"
     
  End Select

End Function




Function calcular_dataSLA(dDataInicial As Date, dHoraInicial As Date, vTempoResposta As Variant, sMunicipio As String, sUF As String) As Date
'Calcula prazo referente data de abertura da OS (n�o executa c�lculo com data final de atendimento)


'Declara��o de vari�veis
Dim dTempoResposta As Date
Dim dREPOSTA As Date
Dim bAlmo�o As Boolean
Dim dHora_Entrada As Date
Dim dHora_Sa�da As Date
Dim dHora_Ini_Almo�o
Dim dHora_Fim_Almo�o
Dim bFeriado As Boolean
        
Dim Entrada_Seg As Date, Ini_Almo�o_Seg As Date, fim_Almo�o_Seg As Date, Sa�da_Seg As Date
Dim Entrada_Ter As Date, Ini_Almo�o_Ter As Date, fim_Almo�o_Ter As Date, Sa�da_Ter As Date
Dim Entrada_Qua As Date, Ini_Almo�o_Qua As Date, fim_Almo�o_Qua As Date, Sa�da_Qua As Date
Dim Entrada_Qui As Date, Ini_Almo�o_Qui As Date, fim_Almo�o_Qui As Date, Sa�da_Qui As Date
Dim Entrada_Sex As Date, Ini_Almo�o_Sex As Date, fim_Almo�o_Sex As Date, Sa�da_Sex As Date
Dim Entrada_Sab As Date, Ini_Almo�o_Sab As Date, fim_Almo�o_Sab As Date, Sa�da_Sab As Date
Dim Entrada_Dom As Date, Ini_Almo�o_Dom As Date, fim_Almo�o_Dom As Date, Sa�da_Dom As Date
    
Dim Expediente_Seg As Boolean, Almo�o_Na_Seg As Boolean
Dim Expediente_Ter As Boolean, Almo�o_Na_Ter As Boolean
Dim Expediente_Qua As Boolean, Almo�o_Na_Qua As Boolean
Dim Expediente_Qui As Boolean, Almo�o_Na_Qui As Boolean
Dim Expediente_Sex As Boolean, Almo�o_Na_Sex As Boolean
Dim Expediente_Sab As Boolean, Almo�o_No_Sab As Boolean
Dim Expediente_Dom As Boolean, Almo�o_No_Dom As Boolean
    
'Converter o Tempo de resposta para n�mero formato num�rico
dTempoResposta = CDbl(CDate(vTempoResposta))
    
'Definidir dias que haver�o expedientes: True -> Tem expediente, False -> N�o tem Expediente
Expediente_Seg = True:  Almo�o_Na_Seg = True
Expediente_Ter = True:  Almo�o_Na_Ter = True
Expediente_Qua = True:  Almo�o_Na_Qua = True
Expediente_Qui = True:  Almo�o_Na_Qui = True
Expediente_Sex = True:  Almo�o_Na_Sex = True
Expediente_Sab = False: Almo�o_No_Sab = True
Expediente_Dom = False: Almo�o_No_Dom = True
    
'Definir Hor�rio de Expediente
Entrada_Seg = "8:00": Sa�da_Seg = "17:00"
Entrada_Ter = "8:00": Sa�da_Ter = "17:00"
Entrada_Qua = "8:00": Sa�da_Qua = "17:00"
Entrada_Qui = "8:00": Sa�da_Qui = "17:00"
Entrada_Sex = "8:00": Sa�da_Sex = "17:00"
Entrada_Sab = "8:00": Sa�da_Sab = "17:00"
Entrada_Dom = "8:00": Sa�da_Dom = "17:00"
     
'Definir in�cio e fim do almo�o
Ini_Almo�o_Seg = "12:00": fim_Almo�o_Seg = "13:00"
Ini_Almo�o_Ter = "12:00": fim_Almo�o_Ter = "13:00"
Ini_Almo�o_Qua = "12:00": fim_Almo�o_Qua = "13:00"
Ini_Almo�o_Qui = "12:00": fim_Almo�o_Qui = "13:00"
Ini_Almo�o_Sex = "12:00": fim_Almo�o_Sex = "13:00"
Ini_Almo�o_Sab = "12:00": fim_Almo�o_Sab = "13:00"
Ini_Almo�o_Dom = "12:00": fim_Almo�o_Dom = "13:00"
              
Do Until dTempoResposta = "00:00"

  bFeriado = pesquisar_EFeriado(dDataInicial, sMunicipio, sUF)
  
  'Validar Feriado. Se for, pula o dia
  If bFeriado Then
      dHoraInicial = 0
      GoTo Pr�ximoDia
  End If
        
  'VALIDAR DIA DA SEMANA. DEFINE HOR�RIOS DE ENTRADA, SA�DA E ALMO�O NAS VARI�VEIS DE CONTROLE
  Select Case WorksheetFunction.Weekday(dDataInicial, 2)
    Case 1 'Segunda-feira
      If Expediente_Seg Then
        dHora_Entrada = Entrada_Seg: dHora_Sa�da = Sa�da_Seg
        dHora_Ini_Almo�o = Ini_Almo�o_Seg: dHora_Fim_Almo�o = fim_Almo�o_Seg
        bAlmo�o = Almo�o_Na_Seg
      Else
        GoTo Pr�ximoDia
      End If
    Case 2 'Ter�a-Feira
      If Expediente_Ter Then
        dHora_Entrada = Entrada_Ter: dHora_Sa�da = Sa�da_Ter
        dHora_Ini_Almo�o = Ini_Almo�o_Ter: dHora_Fim_Almo�o = fim_Almo�o_Ter
        bAlmo�o = Almo�o_Na_Ter
      Else
        GoTo Pr�ximoDia
      End If
    Case 3 'Quarta-feira
      If Expediente_Qua Then
        dHora_Entrada = Entrada_Qua: dHora_Sa�da = Sa�da_Qua
        dHora_Ini_Almo�o = Ini_Almo�o_Qua: dHora_Fim_Almo�o = fim_Almo�o_Qua
        bAlmo�o = Almo�o_Na_Qua
      Else
        GoTo Pr�ximoDia
      End If
    Case 4 'Quinta-feira
      If Expediente_Qui Then
        dHora_Entrada = Entrada_Qui: dHora_Sa�da = Sa�da_Qui
        dHora_Ini_Almo�o = Ini_Almo�o_Qui: dHora_Fim_Almo�o = fim_Almo�o_Qui
        bAlmo�o = Almo�o_Na_Qui
      Else
        GoTo Pr�ximoDia
      End If
    Case 5 'Sexta-feira
      If Expediente_Sex Then
        dHora_Entrada = Entrada_Sex: dHora_Sa�da = Sa�da_Sex
        dHora_Ini_Almo�o = Ini_Almo�o_Sex: dHora_Fim_Almo�o = fim_Almo�o_Sex
        bAlmo�o = Almo�o_Na_Sex
      Else
        GoTo Pr�ximoDia
      End If
    Case 6 'S�bado
      If Expediente_Sab Then
        dHora_Entrada = Entrada_Sab: dHora_Sa�da = Sa�da_Sab
        dHora_Ini_Almo�o = Ini_Almo�o_Sab: dHora_Fim_Almo�o = fim_Almo�o_Sab
        bAlmo�o = Almo�o_No_Sab
      Else
        GoTo Pr�ximoDia
      End If
    Case 7 'Domingo
      If Expediente_Dom Then
        dHora_Entrada = Entrada_Dom: dHora_Sa�da = Sa�da_Dom
        dHora_Ini_Almo�o = Ini_Almo�o_Dom: dHora_Fim_Almo�o = fim_Almo�o_Dom
        bAlmo�o = Almo�o_No_Dom
      Else
        GoTo Pr�ximoDia
      End If
  End Select
                    
  'VALIDAR dHoraInicial PARA DENTRO DO EXPEDIENTE DO DIA
  If dHoraInicial > dHora_Sa�da Then 'Caso hora inicial esteja ap�s termino do expediente
    dHoraInicial = dHora_Sa�da
    
  ElseIf dHoraInicial < dHora_Entrada Then 'Caso hora inicial esteja antes do in�cio do expediente
    dHoraInicial = dHora_Entrada
    
  ElseIf dHoraInicial > dHora_Ini_Almo�o And _
  dHoraInicial < dHora_Fim_Almo�o Then 'Caso, no horario de almo�o
    dHoraInicial = dHora_Fim_Almo�o
    
  End If
                        
  'IN�CIO DO C�LCULO DO SLA DESCONTANDO O TEMPO AT� ZERAR dTempoResposta
  If bAlmo�o Then 'SE HORARIO ALMO�O = TRUE
    If dHoraInicial <= dHora_Ini_Almo�o Then 'Inicia contagem antes do almo�o
      If dDataInicial + dHoraInicial + dTempoResposta > _
      dDataInicial + dHora_Ini_Almo�o Then 'SE n�o zerou o dTempoResposta
        dTempoResposta = dTempoResposta - (dHora_Ini_Almo�o - dHoraInicial) 'Desconta tempo da manh�
        dHoraInicial = dHora_Fim_Almo�o 'Definir novamente hor�rio inicial
        dDataInicial = dDataInicial - 1 'Reduzir um dia para rodar novamente o la�o e cair no mesmo dia
      Else 'Terminou o chamado no dia
        dREPOSTA = dDataInicial + dHoraInicial + dTempoResposta 'Define o tempo SLA
        dTempoResposta = CDate("00:00") 'Zera tempo resposta para sair do la�o
      End If
    Else 'Inicia contagem ap�s o almo�o
      If dDataInicial + dHoraInicial + dTempoResposta > _
      dDataInicial + dHora_Sa�da Then 'Checa se ultrapassa o dia atual
        dTempoResposta = dTempoResposta - (dHora_Sa�da - dHoraInicial) 'Desconta o tempo do dia
        dHoraInicial = 0 'Zera hora inicial para definir novamente in�cio do expediente ao voltar o la�o
      Else 'Terminou o chamado no dia
        dREPOSTA = dDataInicial + dHoraInicial + dTempoResposta 'Define o tempo SLA
        dTempoResposta = CDate("00:00") 'Zera tempo resposta para sair do la�o
      End If
    End If
  Else 'SE HORARIO ALMO�O = FALSE
    If dHoraInicial <= dHora_Sa�da Then 'Checa se hora de in�cio est� dentro do expediente
      If dDataInicial + dHoraInicial + dTempoResposta > _
      dDataInicial + dHora_Sa�da Then 'Checa se o tempo avan�a para dia seguinte
        dTempoResposta = dTempoResposta - (dHora_Sa�da - dHoraInicial) 'Desconta o tempo do dia
        dHoraInicial = 0 'Zera hora inicial para definir novamente inicio do expediente ao voltar o la�o
      Else 'Terminou o chamado no dia
        dREPOSTA = dDataInicial + dHoraInicial + dTempoResposta 'Define o tempo SLA
        dTempoResposta = CDate("00:00") 'Zera tempo resposta para sair do la�o
      End If
    End If
  End If
                
Pr�ximoDia:
  dDataInicial = dDataInicial + 1 'Adicionar um dia para rodar novamente o la�o
Loop 'Retornar o la�o para o dia seguinte
    
calcular_dataSLA = dREPOSTA
                
End Function

Function verificar_arq_atual(sDir As String) As String
'Verifica �ltimo arquivo atualizado em uma pasta com v�rios arquivos

Dim oFSO As Object
Dim oPasta As Object
Dim oArquivo As Object
Dim dtDataArq As Date
Dim sArq As String

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oPasta = oFSO.GetFolder(sDir)

'Verifica cada arquivo dentro de uma pasta
For Each oArquivo In oPasta.Files
    If oArquivo.datelastmodified > dtDataArq And oArquivo.Name Like "*.xls*" Then
        dtDataArq = oArquivo.datelastmodified
        sArq = oArquivo
    End If
Next oArquivo

verificar_arq_atual = sArq

'Limpa atribui��o
Set oFSO = Nothing
Set oPasta = Nothing
Set oArquivo = Nothing

End Function


'Function IsFileOpen(filename As String)
''Fun��o verifica arquivo aberto
'
'    Dim filenum As Integer, errnum As Integer
'
'    On Error Resume Next   ' Turn error checking off.
'    filenum = FreeFile()   ' Get a free file number.
'    ' Attempt to open the file and lock it.
'    Open filename For Input Lock Read As #filenum
'    Close filenum          ' Close the file.
'    errnum = Err           ' Save the error number that occurred.
'    On Error GoTo 0        ' Turn error checking back on.
'
'    ' Check to see which error occurred.
'    Select Case errnum
'
'        ' No error occurred.
'        ' File is NOT already open by another user.
'        Case 0
'         IsFileOpen = False
'
'        ' Error number for "Permission Denied."
'        ' File is already opened by another user.
'        Case 70
'            IsFileOpen = True
'
'        ' Another error occurred.
'        Case Else
'            Error errnum
'    End Select
'
'End Function

