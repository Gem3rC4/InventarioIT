#requires -Version 5.1
<#
.SYNOPSIS
Prepara el receptor central y genera el comando que ejecutaran los demas equipos.
.DESCRIPTION
Use -PrepararRed una vez en una consola elevada para la reserva HTTP y el firewall de
la red local. Sin ese parametro solo genera configuracion y comandos, sin cambiar Windows.
#>
[CmdletBinding()]
param(
    [string]$ExcelUrl,
    [string]$Tabla,
    [string]$DireccionIP,
    [ValidateRange(1024,65535)][int]$Puerto = 8765,
    [string]$TenantId,
    [string]$ClientId,
    [string]$ConfigPath = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'InventarioEquipos\servidor.config.json'),
    [switch]$PrepararRed
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Test-PrivateIPv4 {
    param([string]$Address)
    $parsedAddress = $null
    if (-not [Net.IPAddress]::TryParse($Address,[ref]$parsedAddress) -or $parsedAddress.AddressFamily -ne [Net.Sockets.AddressFamily]::InterNetwork) { return $false }
    $bytes = $parsedAddress.GetAddressBytes()
    return ($bytes[0] -eq 10 -or ($bytes[0] -eq 172 -and $bytes[1] -ge 16 -and $bytes[1] -le 31) -or ($bytes[0] -eq 192 -and $bytes[1] -eq 168))
}

function New-ReceiverKey {
    $bytes = [byte[]]::new(32)
    $generator = [Security.Cryptography.RandomNumberGenerator]::Create()
    try { $generator.GetBytes($bytes) } finally { $generator.Dispose() }
    return [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+','-').Replace('/','_')
}

function Invoke-ReceiverSetup {
    param([string]$ExcelUrl,[string]$Tabla,[string]$DireccionIP,[int]$Puerto,[string]$TenantId,[string]$ClientId,[string]$ConfigPath,[switch]$PrepararRed)
    if ($env:OS -ne 'Windows_NT') { throw 'Ejecute el configurador en el equipo Windows central.' }
    if ($PrepararRed) {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = [Security.Principal.WindowsPrincipal]::new($identity)
        if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'PrepararRed necesita PowerShell como administrador. Se usa para la reserva HTTP y una regla limitada a la subred local.' }
    }
    if (-not $DireccionIP) {
        $addresses = @(Get-NetIPConfiguration -ErrorAction Stop | Where-Object { $_.IPv4DefaultGateway } | ForEach-Object { $_.IPv4Address.IPAddress } | Where-Object { Test-PrivateIPv4 $_ } | Sort-Object -Unique)
        if ($addresses.Count -eq 1) { $DireccionIP = $addresses[0] }
        else {
            Write-Host ('Direcciones locales detectadas: ' + ($addresses -join ', '))
            $DireccionIP = Read-Host 'Indique la IP privada del equipo central accesible desde los demas equipos'
        }
    }
    if (-not (Test-PrivateIPv4 $DireccionIP)) { throw 'Se requiere una IP privada IPv4 (10.x, 172.16-31.x o 192.168.x).' }
    if (-not $ExcelUrl) { $ExcelUrl = Read-Host 'Pegue el enlace completo del Excel de SharePoint' }
    Import-Module (Join-Path $PSScriptRoot 'Inventario.Funciones.psm1') -Force -ErrorAction Stop
    $null = ConvertTo-SharingToken $ExcelUrl
    $configDirectory = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($ConfigPath))
    $null = New-Item -ItemType Directory -Path $configDirectory -Force
    $key = $null
    if (Test-Path -LiteralPath $ConfigPath -PathType Leaf) {
        $existing = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $key = [string](Get-OptionalProperty $existing 'ApiKey')
    }
    if (-not $key -or $key.Length -lt 32 -or $key -notmatch '^[A-Za-z0-9_-]+$') { $key = New-ReceiverKey }
    $listenUri = 'http://{0}:{1}/' -f $DireccionIP,$Puerto
    $queueDirectory = Join-Path $configDirectory 'cola'
    $configuration = [ordered]@{
        ExcelUrl = $ExcelUrl; Tabla = $Tabla; TenantId = $TenantId; ClientId = $ClientId
        ApiKey = $key; ListenUri = $listenUri; QueueDirectory = $queueDirectory
    }
    $configuration | ConvertTo-Json | Set-Content -LiteralPath $ConfigPath -Encoding UTF8
    @{ ServidorUrl = $listenUri; ClaveApi = $key } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $configDirectory 'cliente.config.json') -Encoding UTF8
    $sourceUrl = 'https://raw.githubusercontent.com/Gem3rC4/InventarioIT/main/Inventario-Equipo.ps1'
    $clientCommand = '[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12; $codigoInventario = Invoke-RestMethod -Uri ''' + $sourceUrl + '''; & ([scriptblock]::Create($codigoInventario)) -ServidorUrl ''' + $listenUri + ''' -ClaveApi ''' + $key + ''''
    $commandPath = Join-Path $configDirectory 'Comando-Equipos.txt'
    $clientCommand | Set-Content -LiteralPath $commandPath -Encoding UTF8
    $serverScript = Join-Path $PSScriptRoot 'Servidor-Inventario.ps1'
    $startCommand = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' + $serverScript + '" -ConfigPath "' + [IO.Path]::GetFullPath($ConfigPath) + '" -InstalarDependencias'
    $cmdContent = "@echo off`r`nchcp 65001 >nul`r`n" + $startCommand + "`r`n"
    [IO.File]::WriteAllText((Join-Path $configDirectory 'Iniciar-Servidor.cmd'),$cmdContent,[Text.UTF8Encoding]::new($false))
    if ($PrepararRed) {
        # Reserva HTTP.sys para el SID del usuario del receptor. "netsh show" devuelve codigo 0
        # aunque no exista la reserva, por eso se interpreta la salida por prefijo y SDDL.
        $receiverIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
        try {
            $account = $receiverIdentity.Name
            $sid = $receiverIdentity.User.Value
        } finally { $receiverIdentity.Dispose() }
        Write-Host ('Preparando recepcion HTTP: {0} - usuario de Windows: {1}' -f $listenUri,$account)
        $showOutput = & netsh.exe http show urlacl ("url=$listenUri") | Out-String
        if ($LASTEXITCODE -ne 0) { throw ('No se pudo consultar la reserva HTTP de {0}. Detalle: {1}' -f $listenUri,$showOutput.Trim()) }
        $reservationState = Get-UrlAclState -NetshOutput $showOutput -UrlPrefix $listenUri -Sid $sid
        if ($reservationState -eq 'Missing') {
            $addOutput = & netsh.exe http add urlacl ("url=$listenUri") ("sddl=D:(A;;GX;;;$sid)") | Out-String
            if ($LASTEXITCODE -ne 0) { throw ('No se pudo reservar {0}. Detalle: {1}' -f $listenUri,$addOutput.Trim()) }
            $verifyOutput = & netsh.exe http show urlacl ("url=$listenUri") | Out-String
            if ((Get-UrlAclState -NetshOutput $verifyOutput -UrlPrefix $listenUri -Sid $sid) -ne 'Owned') { throw ('La reserva de {0} no quedo asignada a {1}. Revise: netsh http show urlacl url={0}' -f $listenUri,$account) }
            Write-Host 'La reserva HTTP quedo creada para este usuario.' -ForegroundColor Green
        } elseif ($reservationState -eq 'Owned') {
            Write-Host 'La reserva HTTP ya pertenece al usuario del receptor; se conservo.'
        } else {
            $message = "La direccion $listenUri ya esta reservada para otra cuenta de Windows y no se modifico.`n" +
                "Opcion 1: use otro puerto, por ejemplo agregando -Puerto $($Puerto + 1).`n" +
                "Opcion 2: si esa reserva es de una prueba anterior de InventarioIT, eliminela y repita: netsh http delete urlacl url=$listenUri"
            throw $message
        }
        $ruleName = 'InventarioIT-Central-{0}' -f $Puerto
        $existingRule = Get-NetFirewallRule -Name $ruleName -ErrorAction SilentlyContinue
        if (-not $existingRule) {
            $null = New-NetFirewallRule -Name $ruleName -Group 'InventarioIT' -DisplayName ('Inventario IT central - puerto ' + $Puerto) -Direction Inbound -Action Allow -Protocol TCP -LocalAddress $DireccionIP -LocalPort $Puerto -RemoteAddress LocalSubnet -Profile Domain,Private
        } else {
            if ([string]$existingRule.Group -ne 'InventarioIT') { throw 'Existe una regla con el mismo nombre que no pertenece a InventarioIT. Su administrador debe revisarla; no se modifico.' }
            $null = Set-NetFirewallRule -Name $ruleName -Enabled True -Direction Inbound -Action Allow -Profile Domain,Private
            $existingRule | Get-NetFirewallPortFilter | Set-NetFirewallPortFilter -Protocol TCP -LocalPort $Puerto | Out-Null
            $existingRule | Get-NetFirewallAddressFilter | Set-NetFirewallAddressFilter -LocalAddress $DireccionIP -RemoteAddress LocalSubnet | Out-Null
            Write-Host 'Se actualizo la regla de InventarioIT para esta IP y la subred local.'
        }
    }
    Write-Host "Configuracion del central: $ConfigPath" -ForegroundColor Green
    Write-Host "Comando para los equipos: $commandPath" -ForegroundColor Green
    Write-Host 'El comando generado contiene la clave de recepcion. Compartalo solo con el personal que ejecuta el inventario.'
    Write-Host 'Inicie el receptor en este equipo con el archivo Iniciar-Servidor.cmd generado. Microsoft se autentica solo en el central.'
    Write-Host 'Mantenga una IP fija o una reserva DHCP para este equipo. HTTP se limita a esta red local; para redes no confiables configure HTTPS.'
    return [pscustomobject]@{ ConfigPath = [IO.Path]::GetFullPath($ConfigPath); ListenUri = $listenUri; Comando = $commandPath; Inicio = Join-Path $configDirectory 'Iniciar-Servidor.cmd' }
}

if ($MyInvocation.InvocationName -ne '.') {
    $setupArguments = @{} + $PSBoundParameters
    if (-not $setupArguments.ContainsKey('Puerto')) { $setupArguments.Puerto = $Puerto }
    if (-not $setupArguments.ContainsKey('ConfigPath')) { $setupArguments.ConfigPath = $ConfigPath }
    Invoke-ReceiverSetup @setupArguments
}
