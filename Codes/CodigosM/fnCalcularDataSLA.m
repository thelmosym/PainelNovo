let
    fnCalcularDataSLA = (dataInicial as date, horaInicial as time, tempoResposta as duration, optional municipio as text, optional uf as text, optional tbFeriados as table) as datetime =>
        let
            HoraEntrada = #time(8, 0, 0),
            HoraSaida = #time(17, 0, 0),
            HoraIniAlmoco = #time(12, 0, 0),
            HoraFimAlmoco = #time(13, 0, 0),
            TemExpediente = (data as date) as logical =>
                let
                    dow = Date.DayOfWeek(data, Day.Monday),
                    ehDiaUtil = dow <= 4,
                    ehFeriado = if tbFeriados = null then false else
                        let
                            munStr = if municipio = null then "" else Text.Upper(municipio),
                            ufStr = if uf = null then "" else Text.Upper(uf),
                            feriadosData = Table.SelectRows(tbFeriados, each [Data] = data),
                            matchFeriado = Table.RowCount(
                                Table.SelectRows(feriadosData, each
                                    ([UF] = "Todos" or [UF] = null or Text.Upper([UF]) = ufStr) and
                                    ([Município] = "Todos" or [Município] = null or Text.Upper([Município]) = munStr)
                                )
                            ) > 0
                        in
                            matchFeriado
                in
                    ehDiaUtil and not ehFeriado,
            ProcessarDia = (estado as record) as record =>
                let
                    data0 = estado[Data],
                    hora0 = estado[Hora],
                    restante0 = estado[Restante],
                    Resultado =
                        if not TemExpediente(data0) then
                            [ Data = Date.AddDays(data0, 1), Hora = #time(0,0,0), Restante = restante0 ]
                        else
                            let
                                HoraAjustada =
                                    if hora0 > HoraSaida then HoraSaida
                                    else if hora0 < HoraEntrada then HoraEntrada
                                    else if hora0 > HoraIniAlmoco and hora0 < HoraFimAlmoco then HoraFimAlmoco
                                    else hora0,
                                DisponivelManha = if HoraAjustada < HoraIniAlmoco then HoraIniAlmoco - HoraAjustada else #duration(0,0,0,0),
                                CabeNaManha = HoraAjustada < HoraIniAlmoco and restante0 <= DisponivelManha,
                                HoraInicioTarde = if HoraAjustada < HoraIniAlmoco then HoraFimAlmoco else HoraAjustada,
                                RestanteNaTarde = if HoraAjustada < HoraIniAlmoco then restante0 - DisponivelManha else restante0,
                                DisponivelTarde = HoraSaida - HoraInicioTarde,
                                CabeNaTarde = RestanteNaTarde <= DisponivelTarde,
                                ResultadoDia =
                                    if CabeNaManha then
                                        [ Data = data0, Hora = HoraAjustada + restante0, Restante = #duration(0,0,0,0) ]
                                    else if CabeNaTarde then
                                        [ Data = data0, Hora = HoraInicioTarde + RestanteNaTarde, Restante = #duration(0,0,0,0) ]
                                    else
                                        [ Data = Date.AddDays(data0, 1), Hora = #time(0,0,0), Restante = RestanteNaTarde - DisponivelTarde ]
                            in
                                ResultadoDia
                in
                    Resultado,
            EstadoInicial = [ Data = dataInicial, Hora = horaInicial, Restante = tempoResposta ],
            Historico = List.Generate(
                () => EstadoInicial,
                (e) => e[Restante] > #duration(0,0,0,0),
                (e) => ProcessarDia(e)
            ),
            UltimoPendente = if List.Count(Historico) = 0 then EstadoInicial else List.Last(Historico),
            EstadoFinal = if tempoResposta <= #duration(0,0,0,0) then EstadoInicial else ProcessarDia(UltimoPendente)
        in
            // ? CORRECAO: time - time = duration (conversao valida)
            DateTime.From(EstadoFinal[Data]) + (EstadoFinal[Hora] - #time(0,0,0))
in
    fnCalcularDataSLA
