# Configuración de rutas
$origen = "\\129.30.181.254\t$\HONEYWELL TEST DATA\FTD-2024"   # Ruta de la carpeta compartida
$destino = "C:\TEMP\Files"                     # Ruta donde se copiarán los archivos
$cadena = "13U0500PAXK"                  # Texto a buscar en los archivos
$lista = "$destino\lista_archivos.txt"      # Archivo para guardar la lista de archivos encontrados

# Crear la carpeta destino si no existe
if (!(Test-Path $destino)) { New-Item -ItemType Directory -Path $destino }

# Buscar archivos que contengan la cadena y guardar la lista
Get-ChildItem -Path $origen -Recurse -File | 
Where-Object { Select-String -Path $_.FullName -Pattern $cadena -Quiet } | 
Select-Object -ExpandProperty FullName | Out-File -FilePath $lista

# Copiar los archivos encontrados
Get-Content $lista | ForEach-Object { Copy-Item -Path $_ -Destination $destino }