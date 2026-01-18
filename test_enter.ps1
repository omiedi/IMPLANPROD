$texto = @'
Desarmar bomba, limpiar, controlar, armar, ensayar y  acondicionar para el despacho. Incluye:  Piñon y corona nuevos.  Rectificar, cromar y rectificar muñones y extremo cigueñal  (ALTERNATIVA CIGUEÑAL NUEVO), Embujar Cpo.Motriz alojamiento  rodamiento y caja porta empaq. y reemplazar piezas s/lista.  NOTA: Si el cliente opta por recuperar el cigueñal NO  EXISTIRIA GARANTIA DE TRABAJO. (El cigueñal No es SIAM ni  NATIONAL).
'@

Write-Host "=== ANALISIS DE TEXTO ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Longitud total: $($texto.Length) caracteres"
Write-Host ""

$contieneCarriageReturn = $texto.Contains([char]13)
$contieneLineFeed = $texto.Contains([char]10)

Write-Host "Contiene CR (Carriage Return \r): $contieneCarriageReturn" -ForegroundColor $(if($contieneCarriageReturn){'Green'}else{'Red'})
Write-Host "Contiene LF (Line Feed \n): $contieneLineFeed" -ForegroundColor $(if($contieneLineFeed){'Green'}else{'Red'})
Write-Host ""

$cantidadCR = ($texto.ToCharArray() | Where-Object {$_ -eq [char]13}).Count
$cantidadLF = ($texto.ToCharArray() | Where-Object {$_ -eq [char]10}).Count

Write-Host "Cantidad de CR: $cantidadCR"
Write-Host "Cantidad de LF: $cantidadLF"
Write-Host ""

if($cantidadCR -gt 0 -or $cantidadLF -gt 0) {
    Write-Host "=== POSICIONES DE SALTOS DE LINEA ===" -ForegroundColor Yellow
    for($i=0; $i -lt $texto.Length; $i++) {
        if($texto[$i] -eq [char]13) {
            $contexto = $texto.Substring([Math]::Max(0, $i-20), [Math]::Min(40, $texto.Length - [Math]::Max(0, $i-20)))
            Write-Host "Posicion $i : CR" -ForegroundColor Green
            Write-Host "  Contexto: '$contexto'" -ForegroundColor Gray
        }
        if($texto[$i] -eq [char]10) {
            Write-Host "Posicion $i : LF" -ForegroundColor Green
        }
    }
} else {
    Write-Host "=== NO SE ENCONTRARON SALTOS DE LINEA ===" -ForegroundColor Red
    Write-Host "El texto es una sola linea continua"
}

Write-Host ""
Write-Host "=== VISTA DEL TEXTO ===" -ForegroundColor Cyan
Write-Host $texto
