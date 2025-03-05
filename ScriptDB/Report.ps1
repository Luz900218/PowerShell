$accessDB = "C:\T\Data\Operations\Operations_YIELDS.mdb"
$connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$accessDB;"

$conn = New-Object -ComObject ADODB.Connection
$conn.Open($connString)

if ($conn.State -eq 1) {
    Write-Host "Conexión exitosa a Access (.mdb)"
} else {
    Write-Host "Error de conexión"
    exit
}

$query ="SELECT * FROM Operations where DateValue([START TIME]) = Date()-1";

$command = New-Object -ComObject ADODB.Command
$command.ActiveConnection = $conn
$command.CommandText = $query
$recordset = $command.Execute()

if ($recordset.EOF -and $recordset.BOF) {
    Write-Host "No se encontraron datos para la fecha $FechaFiltro"
    $conn.Close()
    exit
}

$results = @()

$columns = @()
for ($i = 0; $i -lt $recordset.Fields.Count; $i++) {
    $columns += $recordset.Fields.Item($i).Name
}

while (-not $recordset.EOF) {
    $row = New-Object PSObject
    foreach ($column in $columns) {
        $value = $recordset.Fields.Item($column).Value

        if ($value -is [DateTime]) {
            $value = $value.ToString("MM/dd/yyyy hh:mm:ss tt")
        }

        $row | Add-Member -MemberType NoteProperty -Name $column -Value $value
    }
    $results += $row
    $recordset.MoveNext()
}

$recordset.Close()
$conn.Close()

$csvPath = "C:\Users\Sensym\Downloads\Resultados.csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "Consulta exportada a CSV exitosamente: $csvPath"