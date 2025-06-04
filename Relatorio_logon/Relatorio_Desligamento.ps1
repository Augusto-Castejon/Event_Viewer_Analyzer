# Caminho do relatório de desligamentos (completo)
$outputPath = "C:\Temp\Relatorio_logon\Relatorio_Desligamento.html"
$startTime = (Get-Date).AddDays(-7)

# Coleta eventos 1074, 1076, 41
$events = Get-WinEvent -FilterHashtable @{
    LogName = 'System';
    Id = 1074, 1076, 41;
    StartTime = $startTime
} | ForEach-Object {
    $xml = [xml]$_.ToXml()
    $eventId = $_.Id
    $props = $_.Properties

    $user = ""
    $proc = ""
    $tipo = ""
    $motivo = ""
    $comentario = ""

    switch ($eventId) {
        1074 {
            $user = $props[6].Value
            $proc = $props[0].Value
            $tipo = if ($props[3].Value -ne $null -and $props[3].Value -ne "") { $props[3].Value } else { "N/A" }
            $motivo = if ($props[4].Value -ne $null -and $props[4].Value -ne "") { $props[4].Value } else { "N/A" }
            $comentario = if ($props[5].Value -ne $null -and $props[5].Value -ne "") { $props[5].Value } else { "" }
        }
        1076 {
            $user = $props[1].Value
            $proc = "Manual (registro de motivo)"
            $tipo = if ($props[3].Value -ne $null -and $props[3].Value -ne "") { $props[3].Value } else { "N/A" }
            $motivo = if ($props[2].Value -ne $null -and $props[2].Value -ne "") { $props[2].Value } else { "N/A" }
            $comentario = if ($props[4].Value -ne $null -and $props[4].Value -ne "") { $props[4].Value } else { "" }
        }
        41 {
            $user = "Sistema"
            $proc = "Kernel-Power"
            $tipo = "Desligamento inesperado"
            $motivo = "Sem resposta do sistema"
            $comentario = "Sistema não foi desligado corretamente"
        }
    }

    [PSCustomObject]@{
        TimeCreated = $_.TimeCreated
        Usuario = $user
        Processo = $proc
        Tipo = $tipo
        Motivo = $motivo
        Comentario = $comentario
        EventID = $eventId
    }
}

# Gera HTML
$html = @"
<html>
<head>
<title>Relatório de Desligamentos</title>
<style>
body { font-family: Arial; margin: 20px; }
h1 { color: #2a3f5f; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
th { background-color: #f2f2f2; }
</style>
</head>
<body>
<img src="images.jpg" alt="Logo Geset" style="height: 80px; margin-bottom: 20px;">
<h1>Relatório de Desligamentos - Últimos 7 dias</h1>
<table>
<tr><th>Data/Hora</th><th>Usuário</th><th>Processo</th><th>Tipo</th><th>Motivo</th><th>Comentário</th><th>Evento</th></tr>
"@

foreach ($event in $events | Sort-Object TimeCreated -Descending) {
    $html += "<tr><td>$($event.TimeCreated)</td><td>$($event.Usuario)</td><td>$($event.Processo)</td><td>$($event.Tipo)</td><td>$($event.Motivo)</td><td>$($event.Comentario)</td><td>$($event.EventID)</td></tr>"
}

$html += "</table></body></html>"
$html | Out-File -Encoding UTF8 $outputPath
Start-Process $outputPath
