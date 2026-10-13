# ============================================================
# 0. CONFIGURACION
# ============================================================
$PingDurationSeconds = 60
$PingIntervalMs = 1000
$PingTimeoutMs = 800
$TcpTimeoutMs = 4000

# Ookla Speedtest CLI (opcional). Mantener desactivado hasta que la empresa
# confirme que su uso corporativo esta autorizado conforme a sus terminos.
$EnableOoklaSpeedtest = $true
$SpeedtestCliPath = Join-Path $PSScriptRoot 'tools\speedtest.exe'
$SpeedtestTimeoutSeconds = 180
$DnsTestNames = @(
    'www.microsoft.com',
    'www.cloudflare.com',
    'www.google.com'
)
$PublicPingTargets = @(
    [PSCustomObject]@{ Name = 'Cloudflare'; Target = '1.1.1.1' },
    [PSCustomObject]@{ Name = 'Google DNS'; Target = '8.8.8.8' }
)


function Measure-PingBatch {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Targets,

        [ValidateRange(5, 600)]
        [int]$DurationSeconds = 60,

        [ValidateRange(250, 5000)]
        [int]$IntervalMs = 1000,

        [ValidateRange(100, 5000)]
        [int]$TimeoutMs = 800
    )

    # Cada destino tiene su propio objeto Ping y su lista de muestras.
    $probes = @(
        foreach ($targetInfo in $Targets) {
            [PSCustomObject]@{
                Name      = [string]$targetInfo.Name
                Layer     = [string]$targetInfo.Layer
                Interface = [string]$targetInfo.Interface
                Target    = [string]$targetInfo.Target
                Ping      = (New-Object System.Net.NetworkInformation.Ping)
                Samples   = (New-Object System.Collections.Generic.List[object])
            }
        }
    )

    if ($probes.Count -eq 0) {
        return @()
    }

    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    $nextProbeMs = 0L

    try {
        while ($clock.Elapsed.TotalSeconds -lt $DurationSeconds) {
            $waitMs = [int]($nextProbeMs - $clock.ElapsedMilliseconds)
            if ($waitMs -gt 0) {
                Start-Sleep -Milliseconds $waitMs
            }

            if ($clock.Elapsed.TotalSeconds -ge $DurationSeconds) {
                break
            }

            $timestamp = Get-Date
            $pending = @()

            # Iniciar todos los pings del ciclo antes de esperar resultados.
            # Esto permite medir los destinos concurrentemente.
            foreach ($probe in $probes) {
                try {
                    $task = $probe.Ping.SendPingAsync($probe.Target, $TimeoutMs)
                    $pending += [PSCustomObject]@{
                        Probe     = $probe
                        Task      = $task
                        Timestamp = $timestamp
                    }
                }
                catch {
                    $probe.Samples.Add([PSCustomObject]@{
                        Timestamp      = $timestamp
                        Success        = $false
                        ResponseTimeMs = $null
                        Status         = 'Error al iniciar ping'
                    })
                }
            }

            # Recoger las respuestas ya iniciadas.
            foreach ($request in $pending) {
                try {
                    $reply = $request.Task.GetAwaiter().GetResult()
                    $success = (
                        $reply.Status -eq
                        [System.Net.NetworkInformation.IPStatus]::Success
                    )
                    $responseTime = if ($success) {
                        [double]$reply.RoundtripTime
                    } else {
                        $null
                    }
                    $status = $reply.Status.ToString()
                }
                catch {
                    $success = $false
                    $responseTime = $null
                    $status = 'Error'
                }

                $request.Probe.Samples.Add([PSCustomObject]@{
                    Timestamp      = $request.Timestamp
                    Success        = $success
                    ResponseTimeMs = $responseTime
                    Status         = $status
                })
            }

            $nextProbeMs += $IntervalMs
        }
    }
    finally {
        $clock.Stop()
        foreach ($probe in $probes) {
            $probe.Ping.Dispose()
        }
    }

    # Calcular estadisticas para cada destino.
    $results = foreach ($probe in $probes) {
        $successfulSamples = @($probe.Samples | Where-Object { $_.Success })
        $sent = $probe.Samples.Count
        $received = $successfulSamples.Count
        $lost = $sent - $received
        $lossPercent = if ($sent -gt 0) {
            [math]::Round(($lost / $sent) * 100, 2)
        } else {
            $null
        }

        $times = @($successfulSamples | ForEach-Object { [double]$_.ResponseTimeMs })
        $minimum = $null
        $average = $null
        $maximum = $null
        $jitter = $null

        if ($times.Count -gt 0) {
            $stats = $times | Measure-Object -Minimum -Average -Maximum
            $minimum = [math]::Round($stats.Minimum, 2)
            $average = [math]::Round($stats.Average, 2)
            $maximum = [math]::Round($stats.Maximum, 2)
        }

        # Variacion media absoluta entre muestras exitosas consecutivas.
        if ($times.Count -gt 1) {
            $differences = @()
            for ($i = 1; $i -lt $times.Count; $i++) {
                $differences += [math]::Abs($times[$i] - $times[$i - 1])
            }
            $jitter = [math]::Round(($differences | Measure-Object -Average).Average, 2)
        }

        [PSCustomObject]@{
            Name            = $probe.Name
            Layer           = $probe.Layer
            Interface       = $probe.Interface
            Target          = $probe.Target
            DurationSeconds = [math]::Round($clock.Elapsed.TotalSeconds, 1)
            Sent            = $sent
            Received        = $received
            Lost            = $lost
            LossPercent     = $lossPercent
            MinimumMs       = $minimum
            AverageMs       = $average
            MaximumMs       = $maximum
            JitterApproxMs  = $jitter
        }
    }

    return @($results)
}

function Test-DnsResolution {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DnsServer,

        [Parameter(Mandatory = $true)]
        [string]$QueryName,

        [string]$ConfiguredOn = 'No especificado'
    )

    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    $success = $false
    $addresses = @()
    $errorMessage = $null

    try {
        $answer = Resolve-DnsName `
            -Name $QueryName `
            -Server $DnsServer `
            -Type A `
            -DnsOnly `
            -ErrorAction Stop

        $addresses = @(
            $answer |
                Where-Object { $_.Type -eq 'A' -and $_.IPAddress } |
                Select-Object -ExpandProperty IPAddress -Unique
        )
        $success = $addresses.Count -gt 0
        if (-not $success) {
            $errorMessage = 'La consulta no devolvio registros IPv4.'
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
    }
    finally {
        $clock.Stop()
    }

    [PSCustomObject]@{
        ConfiguredOn = $ConfiguredOn
        DnsServer    = $DnsServer
        QueryName    = $QueryName
        Success      = $success
        QueryTimeMs  = [math]::Round($clock.Elapsed.TotalMilliseconds, 2)
        Addresses    = $addresses -join ', '
        Error        = $errorMessage
    }
}

function Test-TcpPort {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Target,

        [int]$Port = 443,
        [int]$TimeoutMs = 4000
    )

    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    $client = New-Object System.Net.Sockets.TcpClient
    $asyncResult = $null
    $success = $false
    $status = 'Failed'
    $errorMessage = $null

    try {
        $asyncResult = $client.BeginConnect($Target, $Port, $null, $null)
        $completed = $asyncResult.AsyncWaitHandle.WaitOne($TimeoutMs, $false)

        if ($completed) {
            $client.EndConnect($asyncResult)
            $success = $true
            $status = 'Connected'
        }
        else {
            $status = 'Timeout'
            $errorMessage = "No se completo la conexion en $TimeoutMs ms."
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        $status = 'Error'
    }
    finally {
        $clock.Stop()
        if ($null -ne $asyncResult) {
            try { $asyncResult.AsyncWaitHandle.Close() } catch { }
        }
        $client.Close()
    }

    [PSCustomObject]@{
        Target      = $Target
        Port        = $Port
        Success     = $success
        Status      = $status
        ConnectTimeMs = [math]::Round($clock.Elapsed.TotalMilliseconds, 2)
        Error       = $errorMessage
    }
}

function Test-HttpsEndpoint {
    param(
        [string]$Url = 'https://www.microsoft.com/',
        [int]$TimeoutSeconds = 15
    )

    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    $responseReceived = $false
    $statusCode = $null
    $errorMessage = $null

    try {
        # Usa la configuracion de proxy del sistema/usuario que aplique a PowerShell.
        $response = Invoke-WebRequest `
            -Uri $Url `
            -Method Head `
            -TimeoutSec $TimeoutSeconds `
            -MaximumRedirection 5 `
            -UseBasicParsing `
            -ErrorAction Stop

        $statusCode = [int]$response.StatusCode
        $responseReceived = $true
    }
    catch {
        # Un HTTP 4xx/5xx tambien demuestra que se recibio una respuesta HTTP.
        if ($null -ne $_.Exception.Response) {
            try {
                $statusCode = [int]$_.Exception.Response.StatusCode
                $responseReceived = $true
            } catch { }
        }
        if (-not $responseReceived) {
            $errorMessage = $_.Exception.Message
        }
    }
    finally {
        $clock.Stop()
    }

    [PSCustomObject]@{
        Url             = $Url
        HttpResponseReceived = $responseReceived
        StatusCode      = $statusCode
        DurationMs      = [math]::Round($clock.Elapsed.TotalMilliseconds, 2)
        Error           = $errorMessage
    }
}

function Invoke-OoklaSpeedtest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExecutablePath,

        [ValidateRange(30, 600)]
        [int]$TimeoutSeconds = 180
    )

    # No descargar ni instalar ejecutables automaticamente.
    if (-not (Test-Path -LiteralPath $ExecutablePath -PathType Leaf)) {
        return [PSCustomObject]@{
            Status             = 'No instalado'
            Server             = $null
            DownloadMbps       = $null
            UploadMbps         = $null
            LatencyMs           = $null
            JitterMs            = $null
            PacketLossPercent  = $null
            ResultUrl           = $null
            Error               = "No se encontro speedtest.exe en: $ExecutablePath"
            RawJson             = $null
        }
    }

    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()
    $process = $null

    try {
        # --format=json permite procesar los resultados sin analizar texto visual.
        # No aceptamos la licencia automaticamente: debe revisarse previamente.
        $process = Start-Process `
            -FilePath $ExecutablePath `
            -ArgumentList '--format=json --progress=no' `
            -NoNewWindow `
            -PassThru `
            -RedirectStandardOutput $stdoutPath `
            -RedirectStandardError $stderrPath

        $finished = $process.WaitForExit($TimeoutSeconds * 1000)
        if (-not $finished) {
            try { $process.Kill() } catch { }
            return [PSCustomObject]@{
                Status             = 'Timeout'
                Server             = $null
                DownloadMbps       = $null
                UploadMbps         = $null
                LatencyMs           = $null
                JitterMs            = $null
                PacketLossPercent  = $null
                ResultUrl           = $null
                Error               = "Speedtest supero el limite de $TimeoutSeconds segundos."
                RawJson             = $null
            }
        }

        $process.Refresh()
        $exitCode = $process.ExitCode
        $outputText = Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue
        $errorText = Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue

        if ($exitCode -ne 0) {
            $detail = if (-not [string]::IsNullOrWhiteSpace($errorText)) { $errorText.Trim() } else { 'Sin detalle de error en stderr.' }
            throw "Speedtest finalizo con codigo $exitCode. $detail"
        }

        if ([string]::IsNullOrWhiteSpace($outputText)) {
            throw 'Speedtest no devolvio salida JSON.'
        }

        # Convertir el JSON de Ookla en un objeto PowerShell.
        $jsonStart = $outputText.IndexOf('{')
        if ($jsonStart -lt 0) {
            throw "No se encontro un objeto JSON en la salida. $($outputText.Trim())"
        }
        $jsonText = $outputText.Substring($jsonStart).Trim()
        $data = $jsonText | ConvertFrom-Json -ErrorAction Stop

        # Ookla expresa bandwidth en bytes/segundo; convertir a megabits/segundo.
        $downloadMbps = $null
        if ($null -ne $data.download -and $null -ne $data.download.bandwidth) {
            $downloadMbps = [math]::Round(([double]$data.download.bandwidth * 8) / 1000000, 2)
        }

        $uploadMbps = $null
        if ($null -ne $data.upload -and $null -ne $data.upload.bandwidth) {
            $uploadMbps = [math]::Round(([double]$data.upload.bandwidth * 8) / 1000000, 2)
        }

        $latencyMs = $null
        $jitterMs = $null
        if ($null -ne $data.ping) {
            if ($null -ne $data.ping.latency) { $latencyMs = [math]::Round([double]$data.ping.latency, 2) }
            if ($null -ne $data.ping.jitter)  { $jitterMs = [math]::Round([double]$data.ping.jitter, 2) }
        }

        $packetLoss = $null
        if ($null -ne $data.packetLoss) {
            $packetLoss = [math]::Round([double]$data.packetLoss, 2)
        }

        $serverName = $null
        if ($null -ne $data.server) {
            $serverName = '{0} - {1}, {2}' -f $data.server.name, $data.server.location, $data.server.country
        }

        [PSCustomObject]@{
            Status             = 'Completado'
            Server             = $serverName
            DownloadMbps       = $downloadMbps
            UploadMbps         = $uploadMbps
            LatencyMs           = $latencyMs
            JitterMs            = $jitterMs
            PacketLossPercent  = $packetLoss
            ResultUrl           = if ($null -ne $data.result) { $data.result.url } else { $null }
            Error               = $null
            RawJson             = $jsonText
        }
    }
    catch {
        [PSCustomObject]@{
            Status             = 'Error'
            Server             = $null
            DownloadMbps       = $null
            UploadMbps         = $null
            LatencyMs           = $null
            JitterMs            = $null
            PacketLossPercent  = $null
            ResultUrl           = $null
            Error               = $_.Exception.Message
            RawJson             = $null
        }
    }
    finally {
        if ($null -ne $process) {
            try { $process.Dispose() } catch { }
        }
        Remove-Item -LiteralPath $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

# ============================================================
# RECOLECCION DE INFORMACION
# ============================================================
$networkInterfaces = foreach ($name in @('Wi-Fi', 'Ethernet')) {
    $adapter = Get-NetAdapter -Name $name -ErrorAction SilentlyContinue
    if ($null -eq $adapter) {
        Write-Warning "No se encontro el adaptador '$name'. Se omitira."
        continue
    }

    $interfaceRoute = $null
    if ($adapter.Status -eq 'Up') {
        $interfaceRoute = Get-NetRoute `
            -DestinationPrefix '0.0.0.0/0' `
            -AddressFamily IPv4 `
            -InterfaceIndex $adapter.ifIndex `
            -ErrorAction SilentlyContinue |
            Sort-Object RouteMetric |
            Select-Object -First 1
    }

    if ($adapter.Status -eq 'Up') {
        $config = Get-NetIPConfiguration -InterfaceIndex $adapter.ifIndex -ErrorAction SilentlyContinue
        $ip = @($config.IPv4Address | ForEach-Object { $_.IPAddress }) -join ', '
        $dnsServers = @(
            Get-DnsClientServerAddress `
                -InterfaceIndex $adapter.ifIndex `
                -AddressFamily IPv4 `
                -ErrorAction SilentlyContinue |
                ForEach-Object { $_.ServerAddresses }
        )
    }
    else {
        # No mostrar IP/DNS que puedan permanecer en la configuracion de un adaptador desconectado.
        $ip = ''
        $dnsServers = @()
    }

    [PSCustomObject]@{
        Interface      = $adapter.Name
        Status         = $adapter.Status
        LinkSpeed      = if ($adapter.Status -eq 'Up') { $adapter.LinkSpeed } else { 'N/A (inactivo)' }
        InterfaceIndex = $adapter.ifIndex
        IPv4           = $ip
        DefaultGateway = if ($interfaceRoute) { $interfaceRoute.NextHop } else { $null }
        DNS            = $dnsServers -join ', '
        DNSServers     = @($dnsServers)
        HasRoute       = [bool]$interfaceRoute
    }
}

# Seleccionar la ruta predeterminada efectiva entre interfaces activas.
$defaultRouteCandidates = @(
    foreach ($route in (Get-NetRoute -DestinationPrefix '0.0.0.0/0' -AddressFamily IPv4 -ErrorAction SilentlyContinue)) {
        if ([string]::IsNullOrWhiteSpace([string]$route.NextHop) -or $route.NextHop -eq '0.0.0.0') {
            continue
        }
        $adapter = Get-NetAdapter -InterfaceIndex $route.InterfaceIndex -ErrorAction SilentlyContinue
        if ($null -eq $adapter -or $adapter.Status -ne 'Up') {
            continue
        }
        $ipInterface = Get-NetIPInterface -InterfaceIndex $route.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1
        $interfaceMetric = if ($ipInterface) { [int]$ipInterface.InterfaceMetric } else { 0 }
        [PSCustomObject]@{
            Interface       = $adapter.Name
            InterfaceIndex  = $route.InterfaceIndex
            Gateway         = $route.NextHop
            RouteMetric     = [int]$route.RouteMetric
            InterfaceMetric = $interfaceMetric
            EffectiveMetric = ([int]$route.RouteMetric + $interfaceMetric)
        }
    }
)
$primaryRoute = $defaultRouteCandidates | Sort-Object EffectiveMetric | Select-Object -First 1

# ============================================================
# PREPARAR REPORTES
# ============================================================
$computerName = $env:COMPUTERNAME
$startTime = Get-Date
$reportDirectory = Join-Path $PSScriptRoot 'Reportes'
New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
$runId = '{0}_{1}' -f $computerName, $startTime.ToString('yyyyMMdd_HHmmss')
$reportBase = Join-Path $reportDirectory ("NetworkDiagnostics_{0}" -f $runId)
$transcriptStarted = $false
try {
    Start-Transcript -Path ($reportBase + '.txt') -Force | Out-Null
    $transcriptStarted = $true
}
catch {
    Write-Warning "No se pudo iniciar el registro TXT: $($_.Exception.Message)"
}

# ============================================================
# ENCABEZADO
# ============================================================
Write-Host ''
Write-Host '############################################################' -ForegroundColor Cyan
Write-Host '             DIAGNOSTICO DE RED CORPORATIVA' -ForegroundColor Cyan
Write-Host '############################################################' -ForegroundColor Cyan
Write-Host ("Equipo : {0}" -f $computerName)
Write-Host ("Fecha  : {0}" -f $startTime.ToString('yyyy-MM-dd HH:mm:ss'))
Write-Host ''
Write-Host '--- 1. INVENTARIO DE RED ---' -ForegroundColor Cyan
$networkInterfaces |
    Select-Object Interface, Status, LinkSpeed, IPv4, DefaultGateway, DNS, HasRoute |
    Format-Table -Wrap -AutoSize

if ($primaryRoute) {
    Write-Host ("Ruta predeterminada seleccionada por metrica efectiva: {0}; gateway {1}; metrica {2}" -f $primaryRoute.Interface, $primaryRoute.Gateway, $primaryRoute.EffectiveMetric) -ForegroundColor Gray
}
else {
    Write-Host 'No se encontro una ruta IPv4 predeterminada utilizable en una interfaz activa.' -ForegroundColor Red
}

# ============================================================
#  PING SOSTENIDO
# ============================================================
$pingTargets = @()
foreach ($network in $networkInterfaces) {
    if ($network.Status -eq 'Up' -and $network.HasRoute -and
        -not [string]::IsNullOrWhiteSpace([string]$network.DefaultGateway)) {
        $pingTargets += [PSCustomObject]@{
            Name      = "Gateway ($($network.Interface))"
            Layer     = 'Gateway'
            Interface = $network.Interface
            Target    = $network.DefaultGateway
        }
    }
}

if ($primaryRoute) {
    foreach ($publicTarget in $PublicPingTargets) {
        $pingTargets += [PSCustomObject]@{
            Name      = "Internet ICMP ($($publicTarget.Name))"
            Layer     = 'Internet ICMP'
            Interface = $primaryRoute.Interface
            Target    = $publicTarget.Target
        }
    }
}

Write-Host ''
Write-Host ("--- 2. PING SOSTENIDO CONCURRENTE ({0} segundos) ---" -f $PingDurationSeconds) -ForegroundColor Cyan
Write-Host 'Los destinos comienzan sus mediciones en paralelo; las pruebas de ping pueden tardar cerca de un minuto.'
$pingResults = @(
    Measure-PingBatch `
        -Targets $pingTargets `
        -DurationSeconds $PingDurationSeconds `
        -IntervalMs $PingIntervalMs `
        -TimeoutMs $PingTimeoutMs
)

if ($pingResults.Count -gt 0) {
    $pingResults |
        Select-Object Name, Layer, Interface, Target, Sent, Received, Lost, LossPercent, MinimumMs, AverageMs, MaximumMs, JitterApproxMs |
        Format-Table -Wrap -AutoSize
}
else {
    Write-Host 'No fue posible ejecutar pings: no se detectaron destinos validos.' -ForegroundColor Yellow
}

# ============================================================
# DNS
# ============================================================
Write-Host ''
Write-Host '--- 3. PRUEBA FUNCIONAL DNS ---' -ForegroundColor Cyan
$dnsResults = @()
foreach ($network in $networkInterfaces) {
    if ($network.Status -ne 'Up') { continue }

    $servers = @($network.DNSServers | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
    if ($servers.Count -eq 0) {
        Write-Host ("{0}: no tiene servidores DNS IPv4 configurados." -f $network.Interface) -ForegroundColor Yellow
        continue
    }

    foreach ($server in $servers) {
        foreach ($queryName in $DnsTestNames) {
            Write-Host ("DNS {0} en {1}: consultando {2}..." -f $server, $network.Interface, $queryName)
            $dnsResults += Test-DnsResolution -DnsServer $server -QueryName $queryName -ConfiguredOn $network.Interface
        }
    }
}

if ($dnsResults.Count -gt 0) {
    $dnsResults |
        Select-Object ConfiguredOn, DnsServer, QueryName, Success, QueryTimeMs, Addresses |
        Format-Table -Wrap -AutoSize
}
else {
    Write-Host 'No se realizaron consultas DNS en interfaces activas.' -ForegroundColor Yellow
}

# ============================================================
# CONECTIVIDAD 
# ============================================================
Write-Host ''
Write-Host '--- 4. CONECTIVIDAD TCP/443 ---' -ForegroundColor Cyan
$tcpTargets = @(
    [PSCustomObject]@{ Target = '1.1.1.1'; Port = 443 },
    [PSCustomObject]@{ Target = '8.8.8.8'; Port = 443 }
)
$tcpResults = @(
    foreach ($tcpTarget in $tcpTargets) {
        Write-Host ("Probando TCP {0}:{1}..." -f $tcpTarget.Target, $tcpTarget.Port)
        Test-TcpPort -Target $tcpTarget.Target -Port $tcpTarget.Port -TimeoutMs $TcpTimeoutMs
    }
)
$tcpResults | Format-Table Target, Port, Success, Status, ConnectTimeMs, Error -Wrap -AutoSize

# ============================================================
# PRUEBA HTTPS REAL
# ============================================================
Write-Host ''
Write-Host '--- 5. SOLICITUD HTTPS ---' -ForegroundColor Cyan
Write-Host 'Intentando obtener respuesta HTTP(S) de https://www.microsoft.com/ ...'
$httpsResult = Test-HttpsEndpoint -Url 'https://www.microsoft.com/' -TimeoutSeconds 15
$httpsResult | Format-List Url, HttpResponseReceived, StatusCode, DurationMs, Error

# ============================================================
# TRACERT 
# ============================================================
$tcpSuccessCount = @($tcpResults | Where-Object { $_.Success }).Count
if (-not $httpsResult.HttpResponseReceived -and $tcpSuccessCount -eq 0) {
    Write-Host ''
    Write-Host '--- 6. TRACERT (se ejecuta porque no hubo conectividad TCP/HTTPS comprobada) ---' -ForegroundColor Yellow
    Write-Host 'Nota: los saltos que no responden pueden filtrar ICMP; un asterisco no demuestra por si solo que ese router este fallando.'
    $traceOutput = & "$env:SystemRoot\System32\tracert.exe" -d -h 12 -w 700 1.1.1.1 2>&1
    $traceOutput | ForEach-Object { Write-Host $_ }
}
else {
    Write-Host ''
    Write-Host 'TRACERT omitido: se obtuvo al menos una evidencia de conectividad TCP/HTTPS. Puede ejecutarse manualmente si se necesita investigar una ruta.' -ForegroundColor Gray
}

# ============================================================
# PRUEBA DE VELOCIDAD (OPCIONAL)
# ============================================================
$speedtestResult = $null
Write-Host ''
Write-Host '--- 7. PRUEBA DE VELOCIDAD (OOKLA SPEEDTEST CLI, OPCIONAL) ---' -ForegroundColor Cyan

if (-not $EnableOoklaSpeedtest) {
    Write-Host 'Omitida por configuracion. Solo habilitar si el uso corporativo esta autorizado.' -ForegroundColor Yellow
}
else {
    Write-Host 'La prueba genera trafico significativo y puede afectar temporalmente otras mediciones.' -ForegroundColor Yellow
    Write-Host ("Ejecutable: {0}" -f $SpeedtestCliPath)
    Write-Host 'Iniciando Speedtest'

    $speedtestResult = Invoke-OoklaSpeedtest `
        -ExecutablePath $SpeedtestCliPath `
        -TimeoutSeconds $SpeedtestTimeoutSeconds

    $speedtestResult |
        Select-Object Status, Server, DownloadMbps, UploadMbps, LatencyMs, JitterMs, PacketLossPercent, ResultUrl, Error |
        Format-List

    if ($speedtestResult.Status -eq 'Completado') {
        Write-Host '[OK] Speedtest finalizado.' -ForegroundColor Green
    }
    else {
        Write-Host ("[ADVERTENCIA] Speedtest no completo: {0}" -f $speedtestResult.Status) -ForegroundColor Red
    }
}

Write-Host ''
Write-Host '############################################################' -ForegroundColor Cyan
Write-Host '             INTERPRETACION PRELIMINAR' -ForegroundColor Cyan
Write-Host '############################################################' -ForegroundColor Cyan

$primaryGatewayResult = $null
if ($primaryRoute) {
    $primaryGatewayResult = $pingResults |
        Where-Object { $_.Layer -eq 'Gateway' -and $_.Interface -eq $primaryRoute.Interface } |
        Select-Object -First 1
}

if (-not $primaryRoute) {
    Write-Host '[CRITICO] No se detecto una ruta IPv4 predeterminada en una interfaz activa.' -ForegroundColor Red
}
elseif (-not $primaryGatewayResult) {
    Write-Host '[ADVERTENCIA] No se obtuvo resultado del ping al gateway de la ruta principal.' -ForegroundColor Yellow
}
elseif ($primaryGatewayResult.LossPercent -gt 0) {
    Write-Host ("[ADVERTENCIA] Se observo {0}% de perdida ICMP hacia el gateway. Revisar LAN, NIC/cable, Wi-Fi si aplica, VLAN, switch y gateway; repetir para confirmar." -f $primaryGatewayResult.LossPercent) -ForegroundColor Yellow
}
else {
    Write-Host '[OK] No se observo perdida ICMP hacia el gateway durante la muestra.' -ForegroundColor Green
    if ($primaryGatewayResult.MaximumMs -ge 20) {
        Write-Host ("[OBSERVACION] El RTT maximo al gateway fue {0} ms; investigar el pico si coincide con la lentitud reportada." -f $primaryGatewayResult.MaximumMs) -ForegroundColor Yellow
    }
}

$dnsSuccessCount = @($dnsResults | Where-Object { $_.Success }).Count
if ($dnsResults.Count -eq 0) {
    Write-Host '[SIN DATOS] No se pudo evaluar el funcionamiento DNS.' -ForegroundColor Yellow
}
elseif ($dnsSuccessCount -eq 0) {
    Write-Host '[SOSPECHA DNS] Ninguno de los servidores configurados resolvio los nombres de prueba. Verificar alcance a los DNS, politicas y configuracion.' -ForegroundColor Red
}
elseif ($dnsSuccessCount -lt $dnsResults.Count) {
    Write-Host ("[ADVERTENCIA DNS] {0} de {1} consultas funcionaron; revisar los servidores/nombres fallidos y repetir la prueba." -f $dnsSuccessCount, $dnsResults.Count) -ForegroundColor Yellow
}
else {
    Write-Host ("[OK DNS] Las {0} consultas realizadas devolvieron direcciones IPv4." -f $dnsResults.Count) -ForegroundColor Green
}

$internetPingResults = @($pingResults | Where-Object { $_.Layer -eq 'Internet ICMP' })
$externalPingRespondedCount = @($internetPingResults | Where-Object { $_.Received -gt 0 }).Count
$allExternalIcmpFailed = ($internetPingResults.Count -gt 0) -and ($externalPingRespondedCount -eq 0)
if ($allExternalIcmpFailed -and (($tcpSuccessCount -gt 0) -or $httpsResult.HttpResponseReceived)) {
    Write-Host '[NOTA ICMP] Los destinos publicos no respondieron ping, pero hay evidencia TCP/HTTPS. ICMP puede estar filtrado; no clasificarlo como caida de Internet solo por ping.' -ForegroundColor Yellow
}

if ($httpsResult.HttpResponseReceived) {
    Write-Host ("[OK HTTPS] Se recibio una respuesta HTTP desde el endpoint; codigo {0}; tiempo total {1} ms." -f $httpsResult.StatusCode, $httpsResult.DurationMs) -ForegroundColor Green
}
elseif ($tcpSuccessCount -gt 0) {
    Write-Host '[ADVERTENCIA WEB] TCP/443 funciona hacia al menos un destino, pero no se obtuvo respuesta HTTPS del endpoint probado. Investigar DNS del sistema, proxy, TLS, filtrado web o el servicio.' -ForegroundColor Yellow
}
else {
    Write-Host '[SOSPECHA SALIDA] No hubo respuesta TCP/443 hacia los IP probados ni respuesta HTTPS. Revisar ruta predeterminada, firewall/proxy y salida WAN; el tracert puede aportar pistas.' -ForegroundColor Red
}

if ($primaryGatewayResult -and $primaryGatewayResult.LossPercent -eq 0 -and $dnsSuccessCount -gt 0 -and $httpsResult.HttpResponseReceived) {
    Write-Host '[RESUMEN] Gateway, al menos una consulta DNS y el endpoint HTTPS respondieron. Si persiste la lentitud, falta medir ancho de banda y probar el servicio concreto afectado.' -ForegroundColor Cyan
}

Write-Host ''
Write-Host 'Limites: estas pruebas localizan indicios en este equipo y en estos destinos; no identifican por si solas el switch/puerto, firewall o tramo del ISP responsable.' -ForegroundColor Gray

# ============================================================
# 12. EXPORTAR DATOS Y CERRAR REPORTE
# ============================================================
try {
    if ($pingResults.Count -gt 0) {
        $pingResults | Export-Csv -Path ($reportBase + '_Ping.csv') -NoTypeInformation -Encoding UTF8
    }
    if ($dnsResults.Count -gt 0) {
        $dnsResults | Export-Csv -Path ($reportBase + '_DNS.csv') -NoTypeInformation -Encoding UTF8
    }
    if ($tcpResults.Count -gt 0) {
        $tcpResults | Export-Csv -Path ($reportBase + '_TCP443.csv') -NoTypeInformation -Encoding UTF8
    }
    @($httpsResult) | Export-Csv -Path ($reportBase + '_HTTPS.csv') -NoTypeInformation -Encoding UTF8
    if ($null -ne $speedtestResult) {
        $speedtestResult | Select-Object Status, Server, DownloadMbps, UploadMbps, LatencyMs, JitterMs, PacketLossPercent, ResultUrl, Error |
            Export-Csv -Path ($reportBase + '_Speedtest.csv') -NoTypeInformation -Encoding UTF8
        if (-not [string]::IsNullOrWhiteSpace([string]$speedtestResult.RawJson)) {
            Set-Content -LiteralPath ($reportBase + '_Speedtest.json') -Value $speedtestResult.RawJson -Encoding UTF8
        }
    }
}
catch {
    Write-Warning "No se pudieron exportar uno o mas CSV: $($_.Exception.Message)"
}

Write-Host ''
Write-Host '--- ARCHIVOS GENERADOS ---' -ForegroundColor Cyan
Write-Host ("Carpeta: {0}" -f $reportDirectory)
Write-Host ("Base de archivos: {0}" -f $reportBase)
Write-Host 'El TXT contiene la salida de consola; los CSV guardan las metricas por tipo de prueba.'
Write-Host 'Si Speedtest se habilita y termina correctamente, tambien se guardan Speedtest.csv y la salida original Speedtest.json.'
Write-Host ''
Write-Host '############################################################' -ForegroundColor Cyan
Write-Host '             FIN DEL DIAGNOSTICO' -ForegroundColor Cyan
Write-Host '############################################################' -ForegroundColor Cyan

if ($transcriptStarted) {
    Stop-Transcript | Out-Null
}
