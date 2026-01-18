# Script PowerShell para exportar DESCREP con saltos de línea preservados
# Ajusta la cadena de conexión según tu configuración

$serverInstance = "localhost"  # O tu servidor SQL
$database = "IMPLANPROD"
$codigoReparacion = "974000710"
$outputFile = "C:\temp\descrep_974000710.txt"

# Crear directorio si no existe
$outputDir = Split-Path $outputFile -Parent
if (!(Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# Query SQL
$query = @"
SELECT Especificacion 
FROM DESCREP 
WHERE CodigoReparacion = '$codigoReparacion'
"@

try {
    # Conectar a SQL Server
    $connectionString = "Server=$serverInstance;Database=$database;Integrated Security=True;TrustServerCertificate=True;"
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $connection.Open()
    
    # Ejecutar query
    $command = New-Object System.Data.SqlClient.SqlCommand($query, $connection)
    $reader = $command.ExecuteReader()
    
    if ($reader.Read()) {
        $especificacion = $reader["Especificacion"]
        
        # Guardar a archivo con encoding UTF-8
        [System.IO.File]::WriteAllText($outputFile, $especificacion, [System.Text.Encoding]::UTF8)
        
        Write-Host "✓ Archivo exportado exitosamente a: $outputFile" -ForegroundColor Green
        Write-Host ""
        Write-Host "=== CONTENIDO ===" -ForegroundColor Cyan
        Write-Host $especificacion
        Write-Host ""
        Write-Host "=== ESTADISTICAS ===" -ForegroundColor Yellow
        Write-Host "Longitud total: $($especificacion.Length) caracteres"
        
        $cantidadCR = ($especificacion.ToCharArray() | Where-Object {$_ -eq [char]13}).Count
        $cantidadLF = ($especificacion.ToCharArray() | Where-Object {$_ -eq [char]10}).Count
        
        Write-Host "Cantidad de CR: $cantidadCR"
        Write-Host "Cantidad de LF: $cantidadLF"
        Write-Host "Cantidad de líneas: $(($especificacion -split "`r`n").Count)"
        
        # Abrir archivo en Notepad
        Write-Host ""
        Write-Host "Abriendo archivo en Notepad..." -ForegroundColor Cyan
        Start-Process notepad.exe -ArgumentList $outputFile
    }
    else {
        Write-Host "✗ No se encontró el código de reparación: $codigoReparacion" -ForegroundColor Red
    }
    
    $reader.Close()
    $connection.Close()
}
catch {
    Write-Host "✗ Error: $($_.Exception.Message)" -ForegroundColor Red
}
