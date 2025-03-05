
#$accessDB = "C:\Users\Sensym\Downloads\YIELDS.mdb"

# Cadena de conexión para archivos .mdb
$connString = "Driver={Microsoft Access Driver (*.mdb)};DBQ=C:\Users\Sensym\Downloads\YIELDS.mdb";
#$connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$accessDB;"

# Crear conexión ADO
$conn = New-Object -ComObject ADODB.Connection
$conn.Open($connString)

# Verificar si la conexión fue exitosa
if ($conn.State -eq 1) {
    Write-Host "Conexión exitosa a Access (.mdb)"
} else {
    Write-Host "Error de conexión"
}

# Cerrar la conexión
$conn.Close()