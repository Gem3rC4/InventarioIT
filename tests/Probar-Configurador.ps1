#requires -Version 5.1
<#
.SYNOPSIS
Prueba la interpretacion de reservas HTTP (netsh) sin modificar HTTP.sys ni el firewall.
#>
[CmdletBinding()]
param([string]$ModulePath)
$ErrorActionPreference = 'Stop'
if (-not $ModulePath) { $ModulePath = Join-Path (Split-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) -Parent) 'Inventario.Funciones.psm1' }
Import-Module $ModulePath -Force
$script:Passed = 0
$script:Failures = New-Object 'System.Collections.Generic.List[string]'
function Assert-Equal {
    param([AllowNull()][object]$Actual,[AllowNull()][object]$Expected,[string]$Message)
    if ($Actual -cne $Expected) { throw "$Message (esperado: $Expected; recibido: $Actual)" }
}
function Invoke-Test {
    param([string]$Name,[scriptblock]$Action)
    try { & $Action; $script:Passed++; Write-Host "OK: $Name" -ForegroundColor Green }
    catch { $script:Failures.Add(("{0}: {1}" -f $Name,$_.Exception.Message)); Write-Host ("FALLO: {0}: {1}" -f $Name,$_.Exception.Message) -ForegroundColor Red }
}

$url = 'http://172.16.74.11:8765/'
$sid = 'S-1-5-21-1111111111-2222222222-3333333333-500'
$otherSid = 'S-1-5-21-1111111111-2222222222-3333333333-1001'

# Salida real de Windows en espanol cuando NO existe la reserva (codigo de salida 0).
$emptySpanish = "`r`nReservas de direcciones URL:`r`n---------------------------`r`n`r`n"
$emptyEnglish = "`r`nURL Reservations:`r`n-----------------`r`n`r`n"
function New-Reservation {
    param([string]$Label,[string]$Url,[string]$Sddl)
    "`r`nURL Reservations:`r`n-----------------`r`n`r`n    $Label            : $Url`r`n        User: EQUIPO\Usuario`r`n            Listen: Yes`r`n            Delegate: No`r`n            SDDL: $Sddl`r`n`r`n"
}

Invoke-Test 'Lista vacia en espanol se interpreta como sin reserva' {
    Assert-Equal (Get-UrlAclState -NetshOutput $emptySpanish -UrlPrefix $url -Sid $sid) 'Missing' 'Confundio una lista vacia con otra identidad.'
}
Invoke-Test 'Lista vacia en ingles y salida nula son sin reserva' {
    Assert-Equal (Get-UrlAclState -NetshOutput $emptyEnglish -UrlPrefix $url -Sid $sid) 'Missing' 'Ingles vacio.'
    Assert-Equal (Get-UrlAclState -NetshOutput $null -UrlPrefix $url -Sid $sid) 'Missing' 'Salida nula.'
}
Invoke-Test 'Reserva del mismo SID se conserva' {
    $out = New-Reservation 'Reserved URL' $url "D:(A;;GX;;;$sid)"
    Assert-Equal (Get-UrlAclState -NetshOutput $out -UrlPrefix $url -Sid $sid) 'Owned' 'No reconocio la reserva propia.'
}
Invoke-Test 'Reserva en espanol con tildes y sin barra final' {
    $out = New-Reservation ('Direcci' + [char]0x00F3 + 'n URL reservada') 'HTTP://172.16.74.11:8765' "D:(A;;GA;;;$sid)"
    Assert-Equal (Get-UrlAclState -NetshOutput $out -UrlPrefix $url -Sid $sid) 'Owned' 'No normalizo el prefijo.'
}
Invoke-Test 'Reserva de otra cuenta se informa como ajena' {
    $out = New-Reservation 'Reserved URL' $url "D:(A;;GX;;;$otherSid)"
    Assert-Equal (Get-UrlAclState -NetshOutput $out -UrlPrefix $url -Sid $sid) 'Foreign' 'Acepto una reserva ajena.'
}
Invoke-Test 'Una regla de denegacion para el SID no cuenta como propia' {
    $out = New-Reservation 'Reserved URL' $url "D:(D;;GX;;;$sid)"
    Assert-Equal (Get-UrlAclState -NetshOutput $out -UrlPrefix $url -Sid $sid) 'Foreign' 'Acepto una denegacion.'
}
Invoke-Test 'Otro prefijo en la salida no se confunde con el propio' {
    $out = New-Reservation 'Reserved URL' 'http://172.16.74.11:9999/' "D:(A;;GX;;;$sid)"
    Assert-Equal (Get-UrlAclState -NetshOutput $out -UrlPrefix $url -Sid $sid) 'Missing' 'Uso la reserva de otro puerto.'
}
Invoke-Test 'Alias LA del Administrador integrado se reconoce como propio' {
    # Windows muestra "SDDL: D:(A;;GX;;;LA)" para la cuenta Administrador (RID 500) del equipo.
    $localAdmin = ([Security.AccessControl.RawSecurityDescriptor]::new('D:(A;;GX;;;LA)')).DiscretionaryAcl[0].SecurityIdentifier.Value
    Assert-Equal ($localAdmin -match '^S-1-5-21-.+-500$') $true 'LA no corresponde al RID 500.'
    $out = New-Reservation ('Direcci' + [char]0x00F3 + 'n URL reservada') $url 'D:(A;;GX;;;LA)'
    Assert-Equal (Get-UrlAclState -NetshOutput $out -UrlPrefix $url -Sid $localAdmin) 'Owned' 'No reconocio el alias LA.'
    Assert-Equal (Get-UrlAclState -NetshOutput $out -UrlPrefix $url -Sid $otherSid) 'Foreign' 'Asigno LA a otra cuenta.'
}
Invoke-Test 'Alias de grupo BU se convierte a su SID' {
    $out = New-Reservation 'Reserved URL' $url 'D:(A;;GX;;;BU)(A;;GX;;;LS)'
    Assert-Equal (Get-UrlAclState -NetshOutput $out -UrlPrefix $url -Sid 'S-1-5-32-545') 'Owned' 'No convirtio BU.'
}
Invoke-Test 'Permiso sin escucha no cuenta como propio' {
    $out = New-Reservation 'Reserved URL' $url "D:(A;;GR;;;$sid)"
    Assert-Equal (Get-UrlAclState -NetshOutput $out -UrlPrefix $url -Sid $sid) 'Foreign' 'Acepto un permiso sin GX.'
}
Invoke-Test 'Configurador y preparador se analizan sin errores' {
    $root = Split-Path $ModulePath -Parent
    foreach ($file in @('Configurar-Servidor.ps1','Servidor-Inventario.ps1','Inventario-Equipo.ps1')) {
        $tokens = $null; $errors = $null
        $null = [Management.Automation.Language.Parser]::ParseFile((Join-Path $root $file),[ref]$tokens,[ref]$errors)
        Assert-Equal $errors.Count 0 "Errores de sintaxis en $file"
    }
}

Write-Host ("Pruebas configurador correctas: {0}; fallidas: {1}." -f $script:Passed,$script:Failures.Count)
if ($script:Failures.Count) { throw ($script:Failures -join [Environment]::NewLine) }
