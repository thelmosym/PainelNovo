let
    fnCalcularPrazoSLA = (dPrazo as number, dataAbertura as date, horaAbertura as time) as nullable datetime =>
        if dPrazo = 0 then null
        else if dPrazo = 0.5 then
            let
                Base = fnCalcularDataSLA(dataAbertura, horaAbertura, #duration(0,4,0,0)),
                DataBase = Date.From(Base),
                Resultado =
                    if horaAbertura >= #time(8,0,0) and horaAbertura <= #time(11,0,0) then
                        DateTime.From(DataBase) + #duration(0,17,0,0)
                    else if horaAbertura > #time(11,0,0) and horaAbertura <= #time(16,0,0) then
                        fnCalcularDataSLA(DataBase, #time(12,0,0), #duration(0,8,0,0))
                    else
                        DateTime.From(DataBase) + #duration(0,17,0,0)
            in
                Resultado
        else if dPrazo = 1 then fnCalcularDataSLA(dataAbertura, horaAbertura, #duration(0,8,0,0))
        else if dPrazo = 2 then fnCalcularDataSLA(dataAbertura, horaAbertura, #duration(0,16,0,0))
        else if dPrazo = 3 then fnCalcularDataSLA(dataAbertura, horaAbertura, #duration(1,0,0,0))
        else if dPrazo = 4 then fnCalcularDataSLA(dataAbertura, horaAbertura, #duration(1,8,0,0))
        else if dPrazo = 5 then fnCalcularDataSLA(dataAbertura, horaAbertura, #duration(1,16,0,0))
        else null
in
    fnCalcularPrazoSLA
