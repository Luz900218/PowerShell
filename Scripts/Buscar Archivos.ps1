Import-Module ImportExcel

$origen = "\\129.30.181.254\t$\HONEYWELL TEST DATA\FTD-2024"
$destino = "C:\TEMP\Files"
$cadena = "13U0500PAXK"

# Crear la carpeta destino si no existe
if (!(Test-Path $destino)) { New-Item -ItemType Directory -Path $destino }

# Obtener todos los archivos de Excel en la carpeta de origen
$archivos = Get-ChildItem -Path $origen -Recurse -Include *.xlsx, *.xls -File

$archivosEncontrados = @()

foreach ($archivo in $archivos) {
    $datos = Import-Excel -Path $archivo.FullName

    if ($datos -match $cadena) {
        Write-Host "Archivo encontrado: $($archivo.FullName)"
        $archivosEncontrados += $archivo.FullName
        Copy-Item -Path $archivo.FullName -Destination $destino
    }
}

# Mostrar resultado final
if ($archivosEncontrados.Count -eq 0) {
    Write-Host "No se encontraron archivos de Excel con la cadena '$cadena'."
} else {
    Write-Host "Se copiaron los siguientes archivos:"
    $archivosEncontrados | ForEach-Object { Write-Host $_ }
}