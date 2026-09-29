'Funcionalidade KM adicional
                If validar_localizacao(wsTempl_DADOS.Cells(iULin, 2)) Then
                    sCentro = pesquisar_centro(wsTempl_DADOS.Cells(iULin, 19))
                    dKM = calcular_KMAdicional(sCentro, wsTempl_DADOS.Cells(iULin, 22))
                    If dKM > 0 Then
                        wsTempl_DADOS.Cells(iULin, 25) = dKM
                    End If
                Else
                    sCentro = pesquisar_centro(wsTempl_DADOS.Cells(iULin, 19))
                End If


 'Funcionalidade QExec
                dQExec = _
                calcular_QExec(wsTempl_DADOS.Cells(iULin, 2), wsTempl_DADOS.Cells(iULin, 3), wsTempl_DADOS.Cells(iULin, 13), wsTempl_DADOS.Cells(iULin, 14))
                'Verifica o fator de ajuste inserido
                If wsTempl_DADOS.Cells(iULin, 18).Value = Empty Then
                    dFA = 1
                Else
                    dFA = wsTempl_DADOS.Cells(iULin, 18).Value
                End If
                wsTempl_DADOS.Cells(iULin, 27) = dQExec * dFA
                
                'Funcionalidade Unidade
                sUnidade = pesquisar_unidade(wsTempl_DADOS.Cells(iULin, 2))
                
                'Funcionalidade Linha PPU
                sLinhaPPU = pesquisar_linhaPPU(wsTempl_DADOS.Cells(iULin, 2))
                
                wsTempl_DADOS.Cells(iULin, 28) = sUnidade
                
                'Funcionalidade Centro
                wsTempl_DADOS.Cells(iULin, 29) = sCentro
                
                'Funcionalidade Código Centro
                'Obs.: Variável sCentro utilizada da pesquisa anterior
                wsTempl_DADOS.Cells(iULin, 30) = pesquisar_codCentro(sCentro)
                
                'Funcionalidade Linha de serviço PPU
                wsTempl_DADOS.Cells(iULin, 31) = sLinhaPPU





EVOLUCAO DE EMPRESTIMO
COLETA DE EMBALAGEM
COLETA DE ITEM AVULSO
ENTREGA DE EMBALAGEM EXPRESSO
ENTREGA DE EMBALAGEM NORMAL
ENTREGA DE ITEM EXPRESSO
ENTREGA DE ITEM NORMAL

COLETA DE EMBALAGEM
ENTREGA DE EMBALAGEM NORMAL
ENTREGA DE EMBALAGEM EXPRESSO
DEVOLUCAO DE EMPRESTIMO


 GBA
 GDF
 GES
 GRJ
 GSP