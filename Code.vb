'-----------------------
'      VB_CODE
'-----------------------
Sub Ejecutar_SQL_en_AWS()

    'Inicialización de variables
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim ConnectionString As String
    Dim sql As String


    'Cadena de conexión con una BBDD de Athena AWS, parametros: https://athena-downloads.s3.amazonaws.com/drivers/ODBC/athena-preview/SimbaAthenaODBC_1.1.0_preview/Simba+Athena+ODBC+Install+and+Configuration+Guide.pdf
         ConnectionString = "Driver={Simba Athena ODBC Driver};" & _
                            "AwsRegion=us-west-2;" & _
                            "S3OutputLocation=s3://aws-athena-query-results-992974280925-us-west-2/;" & _
                            "AuthenticationType=IAM Credentials;" & _
                            "UID=AKIA6OMPBDDOTPQCWMO6;" & _
                            "PWD=Gb3H3PB7x2aqQsAxK3Biw6SU3VjhgzwsoiAZl5U4;"

    'Abrir conexión con la BBDD
    con.Open ConnectionString

    'Comprueba el estado de la conexion, si resulta 1 = objeto abierto o 0 = objeto cerrado segun W3School "ADO State Property"
    MsgBox (con.State)

    'Timeout en segundos para ejecutar la SQL completa antes de reportar un error
    con.CommandTimeout = 900

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
    "And cast(replace(replace(queue.enqueuetimestamp,'T',' '),'Z','') as timestamp) between cast('2020-04-28 05:00:00' as timestamp) and cast('2020-04-29 04:59:59' as timestamp) " & _
    "order by cast(replace(replace(queue.enqueuetimestamp,'T',' '),'Z','') as timestamp)"

    'Lanzamos la SQL
    rs.Open sql, con
    
    'Copiamos los resultados de la SQL sobre la primera hoja del Excel en la celda A2
    Sheets(1).Range("A2").CopyFromRecordset rs
    
    'Cerramos las conexiones
    rs.Close
    con.Close
    
End Sub





 'Crear tabla en Excel VBA
    Sheet1.ListObjects.Add (xlSrcRange, Range ("A1: D10"),, xlYes) .Name = "myTable1"

Sheets("InteraccionesDATA").Select
    Range("A1").Select
    
    Dim UltCol As Integer
    
UltCol = Cells(1, Cells.Columns.Count).End(xlToLeft).Column
Dim UltRow As Integer
    
    'ActiveSheet.Range(Cells(1, 1), Cells(5, 5)).Select
    
    UltRow = ActiveSheet.Columns("A").Find("*", _
        searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1", Cells(UltRow, UltCol)), , xlYes).Name = _
        "Table1"
Sheets("DATA").Select

LibroOr_Horario.Activate
 Valid1 = Range("B27", Range("B27").End(xlDown)).Count
 LibroOr_Tickets.Activate
 Valid2 = Range("B20", Range("B20").End(xlDown)).Count + 19


 'Limpia Libro
    Set B = Range("B8", Range("B8").End(xlDown))
    Set DEF = Range("D8", Range("F8").End(xlDown))
    Set H = Range("H8", Range("H8").End(xlDown))
    
    Union(B, DEF, H).ClearContents
    Union(B, DEF, H).Interior.ColorIndex = 0
    Columns("J:J").ColumnWidth = 25.43
    Columns("K:K").ColumnWidth = 14
