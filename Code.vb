Public Sync_Status As Boolean
Public WB_Macro As Workbook
Dim Tiempo

Sub Ejecutar_SQL_en_AWS()
'-------------------------
'   AWS_CONNECTION_CODE
'-------------------------

    'Apagamos la actualizacion grafica para reducir recursos durante la ejecucion
     Application.ScreenUpdating = False
     Application.DisplayAlerts = False

    'Inicialización de variables.
     Dim con As New ADODB.Connection
     Dim rs As New ADODB.Recordset
     Dim ConnectionString As String
     Dim sql, FechaFin, FechaIni, FechaActual, FechaLimite As String
     Dim MesPrimerReg, y As Integer

    'Almacenamos el libro actual en variable
     Set WB_Macro = Application.ThisWorkbook 
    
'---- Construimos conexion ODBC ----

    'Cadena de conexión con una ODBC de Athena AWS, parametros: https://athena-downloads.s3.amazonaws.com/drivers/ODBC/athena-preview/SimbaAthenaODBC_1.1.0_preview/Simba+Athena+ODBC+Install+and+Configuration+Guide.pdf
         ConnectionString = "Driver={Simba Athena ODBC Driver};" & _
                            "AwsRegion=us-west-2;" & _
                            "S3OutputLocation=s3://aws-athena-query-results-992974280925-us-west-2/;" & _
                            "AuthenticationType=IAM Credentials;" & _
                            '"UID=Your_User;" & _
                            '"PWD=Your_Password;"

    'Abrirmos conexión con la BBDD.
     con.Open ConnectionString

    'Timeout en segundos para ejecutar la SQL completa antes de reportar un error.
     con.CommandTimeout = 900

    'Comprueba el estado de la conexion, si resulta 1 = objeto abierto o 0 = objeto cerrado segun W3School "ADO State Property"
         Dim x As String
             Select Case con.State
                 Case 0:
                     x = "No se ha podido establecer conexion con AWS."
                 Case 1:
                     x = "Conexion con AWS ha sido Exitosa."
                 Case 2:
                     x = "Conexion con AWS aun esta en proceso."
                 Case 4:
                     x = "Conexion con AWS está ejecutando un comando."
                 Case 8:
                     x = "AWS esta recuperando filas."
             End Select
         MsgBox x, vbInformation, "Estado de Conexion"

'---- Construimos query SQL inteligente ----

    'Obtiene la fecha actual.
     FechaActual = Format(Now(), "yyyy-MM-dd hh:mm:ss")
    'Obtiene la fecha limite.
     FechaLimite = Format(Now(), "yyyy-MM-26 23:59:59") 'UTC: "yyyy-MM-26 23:59:59".

    'Si la fecha actual es menor a la limite, variable FechaIni = ultimo registro de tabla (ultima sincronización).
        If FechaActual < FechaLimite Then

                'Obtenemos la cantidad de filas existentes en la tabla.
                    FechUltReg = WB_Macro.Sheets(1).ListObjects("AWS_Table").ListRows.Count

                'Declaramos variable con inicio del marco de tiempo con la ultima fecha en la tabla para realizar query SQL.
                    FechaIni = WB_Macro.Sheets(1).ListObjects("AWS_Table").DataBodyRange(FechUltReg, 4).Value

                'Elimina el ultimo registro para no ser duplicado por query.
                    y = WB_Macro.Sheets(1).ListObjects("AWS_Table").ListRows.Count
                    WB_Macro.Sheets(1).ListObjects("AWS_Table").ListRows(y).Range.Delete

        Else

                'Obtienemos el mes del primer registro en la tabla.
                 MesPrimerReg = (Month(WB_Macro.Sheets(1).Range("E3").Value))

                
                    'Si el primer registro de la tabla es del mes pasado, eliminamos tabla y comenzamos nuevo registro.
                    If (Month(FechaActual)) <> MesPrimerReg Then
                        
                        'Limpiamos datos antiguos de la tabla.
                            WB_Macro.Sheets(1).ListObjects("AWS_Table").DataBodyRange.Delete

                        'Declaramos variable con inicio del marco de tiempo con primer dia del mes presente para realizar query SQL.
                            FechaIni = Format(Now(), "yyyy-MM-27 05:00:00") 'Para UTC: "yyyy-MM-27 00:00:00" pero hemos restado 5 Hrs.
                        
                    'Si el primer registro de la tabla es del mes actual, variable FechaIni = ultimo registro de tabla (ultima sincronización).
                    Else

                        'Obtenemos la cantidad de filas existentes en la tabla.
                            FechUltReg = WB_Macro.Sheets(1).ListObjects("AWS_Table").ListRows.Count

                        'Declaramos variable con inicio del marco de tiempo con la ultima fecha en la tabla para realizar query SQL.
                            FechaIni = WB_Macro.Sheets(1).ListObjects("AWS_Table").DataBodyRange(FechUltReg, 4).Value

                        'Elimina el ultimo registro para no ser duplicado por query.
                            y = WB_Macro.Sheets(1).ListObjects("AWS_Table").ListRows.Count
                            WB_Macro.Sheets(1).ListObjects("AWS_Table").ListRows(y).Range.Delete

                    End If

        End If

'---- Ejecutamos query SQL ----

    'Declaramos variable con fin del marco de tiempo de datos que queremos obtener con query SQL.
    FechaFin = Format(DateAdd("h", 5, Now()), "yyyy-MM-dd hh:mm:ss")

    'Estructuramos nuestro query SQL.
     sql = "SELECT    queue.name," & _
        "initiationmethod, " & _
        "CASE WHEN agent is not null THEN 1 ELSE 0 END as Agente, " & _
        "replace(replace(queue.enqueuetimestamp,'T',' '),'Z','') as enqueuetimestamp, " & _
        "replace(replace(agent.connectedtoagenttimestamp,'T',' '),'Z','') as connectedtoagenttimestamp, " & _
        "replace(replace(disconnecttimestamp,'T',' '),'Z','') as disconnecttimestamp, " & _
        "queue.duration, " & _
        "agent.agentinteractionduration, " & _
        "agent.username " & _
        "FROM AwsDataCatalog.asd_amr_db.all_ctr_data " & _
        "WHERE substring(queue.name,1,5) = 'Femsa' " & _
        "And cast(replace(replace(queue.enqueuetimestamp,'T',' '),'Z','') as timestamp) between cast('" & FechaIni & "' as timestamp) and cast('" & FechaFin & "' as timestamp) " & _
        "order by cast(replace(replace(queue.enqueuetimestamp,'T',' '),'Z','') as timestamp)"
 
    'Lanzamos el SQL.
     rs.Open sql, con

    'Averiguamos donde pegar los datos obtenidos por el query, posteriormente los pegamos.
     y = WB_Macro.Sheets(1).ListObjects("AWS_Table").ListRows.Count + 3
     Sheets(1).Range("B" & y).CopyFromRecordset rs

    'Cerramos las conexiones.
     rs.Close
     con.Close

    'Calculamos nueva columna para cantidad de tiempo en responder una llamada.
     WB_Macro.Sheets(1).ListObjects("AWS_Table").ListColumns("ResponseTime").DataBodyRange.FormulaR1C1 = "=IF([@Agente]=1,[@connectedtoagenttimestamp]-[@enqueuetimestamp],"""")"

    'Calculamos nueva columna para cantidad de tiempo dentro de llamada.
     WB_Macro.Sheets(1).ListObjects("AWS_Table").ListColumns("CallTime").DataBodyRange.FormulaR1C1 = "=[@disconnecttimestamp]-[@connectedtoagenttimestamp]"

    'Calculamos nueva columna para dias (como control para tablas dinamicas).
     WB_Macro.Sheets(1).ListObjects("AWS_Table").ListColumns("Days").DataBodyRange.FormulaR1C1 = "=TRUNC([@enqueuetimestamp])"
    
    'Calculamos nueva columna para UTC-5 como horario Guadalajara (como control para tablas dinamicas).
     WB_Macro.Sheets(1).ListObjects("AWS_Table").ListColumns("UTC-5").DataBodyRange.FormulaR1C1 = "=MOD([@enqueuetimestamp]-[@Days]+(-5/24),1)"

    'Calculamos nueva columna para clasificar por horas horario UTC-5 (como control para tablas dinamicas).
     WB_Macro.Sheets(1).ListObjects("AWS_Table").ListColumns("Hours").DataBodyRange.FormulaR1C1 = "=TIME(HOUR([@[UTC-5]]),FLOOR(MINUTE([@[UTC-5]]),30),0)"

    'Calculamos nueva columna para dias (como control para tablas dinamicas).
     WB_Macro.Sheets(1).ListObjects("AWS_Table").ListColumns("Country").DataBodyRange.FormulaR1C1 = "=IFNA(VLOOKUP([@name],Table_Country[[name]:[Pais]],2,FALSE),""Not found"")"
    
    'Actualizamos tablas dinamicas con cambios.
     WB_Macro.RefreshAll
     
    'Programamos repetir la ejecucion de esta macro (cada 5 min).
     Tiempo = VBA.DateAdd("n", 5, Time)
     Application.OnTime EarliestTime:=Tiempo, Procedure:="Ejecutar_SQL_en_AWS"

    'Encendemos la actualizacion grafica.
     Application.DisplayAlerts = True
     Application.ScreenUpdating = True
     
End Sub

Sub CancelarMacro()

Application.OnTime EarliestTime:=Tiempo, Procedure:="Ejecutar_SQL_en_AWS", Schedule:=False

End Sub

Sub Button_Sync()
 
    If Sync_Status = False Then 
        With Sheets(3).Shapes("TextBox333")
            .Fill.ForeColor.RGB = RGB(55, 213, 135) 'azul 85, 142, 213
            .Line.BackColor.RGB = RGB(198, 217, 241)
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .TextFrame.Characters.Text = "Sincronizacion Activada"
        End With
            Call Ejecutar_SQL_en_AWS
            Sync_Status = True
    Else
        With Sheets(3).Shapes("TextBox333")
            .Fill.ForeColor.RGB = RGB(226, 184, 54)
            .Line.BackColor.RGB = RGB(198, 217, 241)
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .TextFrame.Characters.Text = "Sincronizacion Desactivada"
        End With
            Call CancelarMacro
            Sync_Status = False
    End If

End Sub

'-------------------------
'   Workbook_Open (Al iniciar el libro creamos la variable de estado de sincronizacion para la funcion Button_Sync)
'-------------------------
Private Sub Workbook_Open()
    Sync_Status = False
End Sub




