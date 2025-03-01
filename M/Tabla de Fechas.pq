let CreaTablaFechas = (Fecha_Inicio as date, Fecha_Fin as date) as table =>
  let
    NumeroDeDias = Duration.Days(Duration.From(Fecha_Fin - Fecha_Inicio)),
    Origen = List.Dates(Fecha_Inicio,NumeroDeDias,#duration(1,0,0,0)),
    Tabla= Table.FromList(Origen, Splitter.SplitByNothing()),
    CambiATipoFecha= Table.TransformColumnTypes(Tabla,{{"Column1", type date}}),
    CambiarNombreCampo = Table.RenameColumns(CambiATipoFecha,{{"Column1", "Fecha"}}),
    InsertaAnyo = Table.AddColumn(CambiarNombreCampo, "Año", each Date.Year([Fecha])),
		InsertaSemana=Table.AddColumn(InsertaAnyo, "Semana", each Date.WeekOfYear([Fecha], Day.Monday)),
    InsertaTrimestre = Table.AddColumn(InsertaSemana, "Trimestre", each Date.QuarterOfYear([Fecha])),
    InsertaMes = Table.AddColumn(InsertaTrimestre, "Mes", each Date.Month([Fecha])),
    InsertaBimestre = Table.AddColumn(InsertaMes, "Bimestre", each Number.RoundUp(Value.Divide([Mes],2),0)),
    InsertaCuatrimestre = Table.AddColumn(InsertaBimestre, "Cuatrimestre", each Number.RoundUp(Value.Divide([Mes],4),0)),
    InsertaSemestre = Table.AddColumn(InsertaCuatrimestre, "Semestre", each Number.RoundUp(Value.Divide([Mes],6),0)),
    InsertaDia = Table.AddColumn(InsertaSemestre, "Día", each Date.Day([Fecha])),
    InsertDiaInt = Table.AddColumn(InsertaDia, "Año-Mes-Día", each [Año] * 10000 + [Mes] * 100 + [Día]),
    InsertaNombreMes = Table.AddColumn(InsertDiaInt, "Nombre del mes tmp", each Date.ToText([Fecha], "MMMM"), type text),
    CambiaNombreMes = Table.AddColumn(InsertaNombreMes, "Nombre del mes", each Text.Upper(Text.Range([Nombre del mes tmp], 0, 1)) & Text.Range([Nombre del mes tmp], 1, Text.Length([Nombre del mes tmp])-1)),
    EliminarMesTmp = Table.RemoveColumns(CambiaNombreMes,{"Nombre del mes tmp"}),
    InsertaMesAnyo = Table.AddColumn(EliminarMesTmp, "MesAño", each (try(Text.Range([Nombre del mes],0,3)) otherwise [Nombre del mes]) & " " & Number.ToText([Año])),
    InsertaNombreTrim = Table.AddColumn(InsertaMesAnyo, "Nombre Trimestre", each "Q" & Number.ToText([Trimestre]) & " " & Number.ToText([Año])),
    InsertaNombreBim = Table.AddColumn(InsertaNombreTrim, "Nombre Bimestre", each "B" & Number.ToText([Bimestre]) & " " & Number.ToText([Año])),
    InsertaNombreCua = Table.AddColumn(InsertaNombreBim, "Nombre Cuatrimestre", each "C" & Number.ToText([Cuatrimestre]) & " " & Number.ToText([Año])),
    InsertaNombreSem = Table.AddColumn(InsertaNombreCua, "Nombre Semestre", each "S" & Number.ToText([Semestre]) & " " & Number.ToText([Año])),
    InsertaDiaSemana = Table.AddColumn(InsertaNombreSem, "Num. Dia de la semana", each Date.DayOfWeek([Fecha])),
    InsertaNombreDiaSem = Table.AddColumn(InsertaDiaSemana, "Día de la semana tmp", each Date.ToText([Fecha], "dddd"), type text),
    CambiaNombreDiaSem = Table.AddColumn(InsertaNombreDiaSem, "Día de la semana", each Text.Upper(Text.Range([Día de la semana tmp], 0, 1)) & Text.Range([Día de la semana tmp], 1, Text.Length([Día de la semana tmp])-1)),
    EliminarColTmp = Table.RemoveColumns(CambiaNombreDiaSem,{"Día de la semana tmp"}),
    OrdenMes = Table.AddColumn(EliminarColTmp, "OrdenMes", each Number.FromText(Number.ToText([Año]) & Date.ToText([Fecha], "MM")))
  in
    OrdenMes 
in
CreaTablaFechas
