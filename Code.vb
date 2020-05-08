'-------------------------
'   AWS_CONNECTION_CODE
'-------------------------
Sub Ejecutar_SQL_en_AWS()
    
    'Apagamos la actualizacion grafica para reducir recursos durante la ejecucion
    Application.ScreenUpdating = False
   
    'Inicialización de variables
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim ConnectionString As String
    Dim sql, FechaFin, FechaIni As String
    
    'Cadena de conexión con una BBDD de Athena AWS, parametros: https://athena-downloads.s3.amazonaws.com/drivers/ODBC/athena-preview/SimbaAthenaODBC_1.1.0_preview/Simba+Athena+ODBC+Install+and+Configuration+Guide.pdf
         ConnectionString = "Driver={Simba Athena ODBC Driver};" & _
                            "AwsRegion=us-west-2;" & _
                            "S3OutputLocation=s3://aws-athena-query-results-992974280925-us-west-2/;" & _
                            "AuthenticationType=IAM Credentials;" & _
                            "UID=Your_User;" & _
                            "PWD=Your_Password;"

    'Abrir conexión con la BBDD
    con.Open ConnectionString

    'Timeout en segundos para ejecutar la SQL completa antes de reportar un error
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
    
    'Limpia datos previos de la tabla excel
    ActiveWorkbook.Sheets(1).ListObjects("Colin_Table").DataBodyRange.ClearContents

    'Declaramos variables con el marco de tiempo de datos que queremos obtener en el query SQL
    FechaIni = Format(Now(), "yyyy-MM-dd 00:00:00")
    FechaFin = Format(Now(), "yyyy-MM-dd hh:mm:ss")

    'Esta es la SQL que queremos consultar
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
 
    'Lanzamos la SQL
    rs.Open sql, con
    
    'Copiamos los resultados de la SQL sobre la primera hoja del Excel en la celda A2 (perteneciente a la tabla)
    Sheets(1).Range("A2").CopyFromRecordset rs
    
    'Cerramos las conexiones
    rs.Close
    con.Close
    
    'Encendemos la actualizacion grafica
    Application.ScreenUpdating = True
End Sub