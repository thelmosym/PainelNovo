Function calcular_dataSLA(dDataInicial As Date, dHoraInicial As Date, vTempoResposta As Variant, sMunicipio As String, sUF As String) As Date
'Calcula prazo referente data de abertura da OS (não executa cálculo com data final de atendimento)


'Declaração de variáveis
Dim dTempoResposta As Date
Dim dREPOSTA As Date
Dim bAlmoço As Boolean
Dim dHora_Entrada As Date
Dim dHora_Saída As Date
Dim dHora_Ini_Almoço
Dim dHora_Fim_Almoço
Dim bFeriado As Boolean
        
Dim Entrada_Seg As Date, Ini_Almoço_Seg As Date, fim_Almoço_Seg As Date, Saída_Seg As Date
Dim Entrada_Ter As Date, Ini_Almoço_Ter As Date, fim_Almoço_Ter As Date, Saída_Ter As Date
Dim Entrada_Qua As Date, Ini_Almoço_Qua As Date, fim_Almoço_Qua As Date, Saída_Qua As Date
Dim Entrada_Qui As Date, Ini_Almoço_Qui As Date, fim_Almoço_Qui As Date, Saída_Qui As Date
Dim Entrada_Sex As Date, Ini_Almoço_Sex As Date, fim_Almoço_Sex As Date, Saída_Sex As Date
Dim Entrada_Sab As Date, Ini_Almoço_Sab As Date, fim_Almoço_Sab As Date, Saída_Sab As Date
Dim Entrada_Dom As Date, Ini_Almoço_Dom As Date, fim_Almoço_Dom As Date, Saída_Dom As Date
    
Dim Expediente_Seg As Boolean, Almoço_Na_Seg As Boolean
Dim Expediente_Ter As Boolean, Almoço_Na_Ter As Boolean
Dim Expediente_Qua As Boolean, Almoço_Na_Qua As Boolean
Dim Expediente_Qui As Boolean, Almoço_Na_Qui As Boolean
Dim Expediente_Sex As Boolean, Almoço_Na_Sex As Boolean
Dim Expediente_Sab As Boolean, Almoço_No_Sab As Boolean
Dim Expediente_Dom As Boolean, Almoço_No_Dom As Boolean
    
'Converter o Tempo de resposta para número formato numérico
dTempoResposta = CDbl(CDate(vTempoResposta))
    
'Definidir dias que haverão expedientes: True -> Tem expediente, False -> Não tem Expediente
Expediente_Seg = True:  Almoço_Na_Seg = True
Expediente_Ter = True:  Almoço_Na_Ter = True
Expediente_Qua = True:  Almoço_Na_Qua = True
Expediente_Qui = True:  Almoço_Na_Qui = True
Expediente_Sex = True:  Almoço_Na_Sex = True
Expediente_Sab = False: Almoço_No_Sab = True
Expediente_Dom = False: Almoço_No_Dom = True
    
'Definir Horário de Expediente
Entrada_Seg = "8:00": Saída_Seg = "17:00"
Entrada_Ter = "8:00": Saída_Ter = "17:00"
Entrada_Qua = "8:00": Saída_Qua = "17:00"
Entrada_Qui = "8:00": Saída_Qui = "17:00"
Entrada_Sex = "8:00": Saída_Sex = "17:00"
Entrada_Sab = "8:00": Saída_Sab = "17:00"
Entrada_Dom = "8:00": Saída_Dom = "17:00"
     
'Definir início e fim do almoço
Ini_Almoço_Seg = "12:00": fim_Almoço_Seg = "13:00"
Ini_Almoço_Ter = "12:00": fim_Almoço_Ter = "13:00"
Ini_Almoço_Qua = "12:00": fim_Almoço_Qua = "13:00"
Ini_Almoço_Qui = "12:00": fim_Almoço_Qui = "13:00"
Ini_Almoço_Sex = "12:00": fim_Almoço_Sex = "13:00"
Ini_Almoço_Sab = "12:00": fim_Almoço_Sab = "13:00"
Ini_Almoço_Dom = "12:00": fim_Almoço_Dom = "13:00"
              
Do Until dTempoResposta = "00:00"

  bFeriado = pesquisar_EFeriado(dDataInicial, sMunicipio, sUF)
  
  'Validar Feriado. Se for, pula o dia
  If bFeriado Then
      dHoraInicial = 0
      GoTo PróximoDia
  End If
        
  'VALIDAR DIA DA SEMANA. DEFINE HORÁRIOS DE ENTRADA, SAÍDA E ALMOÇO NAS VARIÁVEIS DE CONTROLE
  Select Case WorksheetFunction.Weekday(dDataInicial, 2)
    Case 1 'Segunda-feira
      If Expediente_Seg Then
        dHora_Entrada = Entrada_Seg: dHora_Saída = Saída_Seg
        dHora_Ini_Almoço = Ini_Almoço_Seg: dHora_Fim_Almoço = fim_Almoço_Seg
        bAlmoço = Almoço_Na_Seg
      Else
        GoTo PróximoDia
      End If
    Case 2 'Terça-Feira
      If Expediente_Ter Then
        dHora_Entrada = Entrada_Ter: dHora_Saída = Saída_Ter
        dHora_Ini_Almoço = Ini_Almoço_Ter: dHora_Fim_Almoço = fim_Almoço_Ter
        bAlmoço = Almoço_Na_Ter
      Else
        GoTo PróximoDia
      End If
    Case 3 'Quarta-feira
      If Expediente_Qua Then
        dHora_Entrada = Entrada_Qua: dHora_Saída = Saída_Qua
        dHora_Ini_Almoço = Ini_Almoço_Qua: dHora_Fim_Almoço = fim_Almoço_Qua
        bAlmoço = Almoço_Na_Qua
      Else
        GoTo PróximoDia
      End If
    Case 4 'Quinta-feira
      If Expediente_Qui Then
        dHora_Entrada = Entrada_Qui: dHora_Saída = Saída_Qui
        dHora_Ini_Almoço = Ini_Almoço_Qui: dHora_Fim_Almoço = fim_Almoço_Qui
        bAlmoço = Almoço_Na_Qui
      Else
        GoTo PróximoDia
      End If
    Case 5 'Sexta-feira
      If Expediente_Sex Then
        dHora_Entrada = Entrada_Sex: dHora_Saída = Saída_Sex
        dHora_Ini_Almoço = Ini_Almoço_Sex: dHora_Fim_Almoço = fim_Almoço_Sex
        bAlmoço = Almoço_Na_Sex
      Else
        GoTo PróximoDia
      End If
    Case 6 'Sábado
      If Expediente_Sab Then
        dHora_Entrada = Entrada_Sab: dHora_Saída = Saída_Sab
        dHora_Ini_Almoço = Ini_Almoço_Sab: dHora_Fim_Almoço = fim_Almoço_Sab
        bAlmoço = Almoço_No_Sab
      Else
        GoTo PróximoDia
      End If
    Case 7 'Domingo
      If Expediente_Dom Then
        dHora_Entrada = Entrada_Dom: dHora_Saída = Saída_Dom
        dHora_Ini_Almoço = Ini_Almoço_Dom: dHora_Fim_Almoço = fim_Almoço_Dom
        bAlmoço = Almoço_No_Dom
      Else
        GoTo PróximoDia
      End If
  End Select
                    
  'VALIDAR dHoraInicial PARA DENTRO DO EXPEDIENTE DO DIA
  If dHoraInicial > dHora_Saída Then 'Caso hora inicial esteja após termino do expediente
    dHoraInicial = dHora_Saída
    
  ElseIf dHoraInicial < dHora_Entrada Then 'Caso hora inicial esteja antes do início do expediente
    dHoraInicial = dHora_Entrada
    
  ElseIf dHoraInicial > dHora_Ini_Almoço And _
  dHoraInicial < dHora_Fim_Almoço Then 'Caso, no horario de almoço
    dHoraInicial = dHora_Fim_Almoço
    
  End If
                        
  'INÍCIO DO CÁLCULO DO SLA DESCONTANDO O TEMPO ATÉ ZERAR dTempoResposta
  If bAlmoço Then 'SE HORARIO ALMOÇO = TRUE
    If dHoraInicial <= dHora_Ini_Almoço Then 'Inicia contagem antes do almoço
      If dDataInicial + dHoraInicial + dTempoResposta > _
      dDataInicial + dHora_Ini_Almoço Then 'SE não zerou o dTempoResposta
        dTempoResposta = dTempoResposta - (dHora_Ini_Almoço - dHoraInicial) 'Desconta tempo da manhã
        dHoraInicial = dHora_Fim_Almoço 'Definir novamente horário inicial
        dDataInicial = dDataInicial - 1 'Reduzir um dia para rodar novamente o laço e cair no mesmo dia
      Else 'Terminou o chamado no dia
        dREPOSTA = dDataInicial + dHoraInicial + dTempoResposta 'Define o tempo SLA
        dTempoResposta = CDate("00:00") 'Zera tempo resposta para sair do laço
      End If
    Else 'Inicia contagem após o almoço
      If dDataInicial + dHoraInicial + dTempoResposta > _
      dDataInicial + dHora_Saída Then 'Checa se ultrapassa o dia atual
        dTempoResposta = dTempoResposta - (dHora_Saída - dHoraInicial) 'Desconta o tempo do dia
        dHoraInicial = 0 'Zera hora inicial para definir novamente início do expediente ao voltar o laço
      Else 'Terminou o chamado no dia
        dREPOSTA = dDataInicial + dHoraInicial + dTempoResposta 'Define o tempo SLA
        dTempoResposta = CDate("00:00") 'Zera tempo resposta para sair do laço
      End If
    End If
  Else 'SE HORARIO ALMOÇO = FALSE
    If dHoraInicial <= dHora_Saída Then 'Checa se hora de início está dentro do expediente
      If dDataInicial + dHoraInicial + dTempoResposta > _
      dDataInicial + dHora_Saída Then 'Checa se o tempo avança para dia seguinte
        dTempoResposta = dTempoResposta - (dHora_Saída - dHoraInicial) 'Desconta o tempo do dia
        dHoraInicial = 0 'Zera hora inicial para definir novamente inicio do expediente ao voltar o laço
      Else 'Terminou o chamado no dia
        dREPOSTA = dDataInicial + dHoraInicial + dTempoResposta 'Define o tempo SLA
        dTempoResposta = CDate("00:00") 'Zera tempo resposta para sair do laço
      End If
    End If
  End If
                
PróximoDia:
  dDataInicial = dDataInicial + 1 'Adicionar um dia para rodar novamente o laço
Loop 'Retornar o laço para o dia seguinte
    
calcular_dataSLA = dREPOSTA
                
End Function


Function gerar_periodo(iPerMes As Integer, iPerAno As Integer) As String
'Função para gerar período entre datas

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

gerar_periodo = sPerAnt & "  à  " & 25 & "/" & iPerMes & "/" & iPerAno

End Function