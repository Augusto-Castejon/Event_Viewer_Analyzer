# Caminho do relatório
$outputPath = "C:\Temp\Relatorio_logon\Relatorio_RDP_Filtravel.html"
$startTime = (Get-Date).AddDays(-7)
$eventIds = 21, 24, 25

# Coleta os eventos do log de sessões RDP
$logs = Get-WinEvent -FilterHashtable @{
    LogName = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational';
    Id = $eventIds;
    StartTime = $startTime
} | ForEach-Object {
    $xml = [xml]$_.ToXml()

    # Extração inteligente do nome de usuário
    $dataFields = @{}
    foreach ($d in $xml.Event.EventData.Data) {
        if ($d.Name) { $dataFields[$d.Name] = $d.'#text' }
    }

    $userRaw = $dataFields['TargetUserName']
    if (-not $userRaw -or $userRaw -eq "") { $userRaw = $dataFields['UserName'] }
    if (-not $userRaw -or $userRaw -eq "") { $userRaw = $dataFields['User'] }

    if (-not $userRaw -or $userRaw -eq "") {
        $alt = $_.Properties | Where-Object { $_.Value -match "\\" -or $_.Value -match "@" }
        if ($alt) { $userRaw = $alt[0].Value }
    }

    if (-not $userRaw -or $userRaw -eq "") {
        $sessionID = $_.Properties[0].Value
        $user = "Desconhecido (SessionID: $sessionID)"
    } else {
        $user = $userRaw
    }

    [PSCustomObject]@{
        Date = $_.TimeCreated.Date
        TimeCreated = $_.TimeCreated
        EventID = $_.Id
        Description = switch ($_.Id) {
            21 { "Logon (Sessão iniciada)" }
            24 { "Desconexão (Sessão RDP desconectada)" }
            25 { "Reconexão (Sessão RDP reconectada)" }
            default { "Outro" }
        }
        User = $user
        SessionID = $dataFields['SessionID']
        Address = $dataFields['Address']
    }
}

# Agrupa os dados por dia
$summary = $logs | Group-Object Date | ForEach-Object {
    $day = Get-Date $_.Name -Format "yyyy-MM-dd"
    $logons = ($_.Group | Where-Object { $_.EventID -eq 21 }).Count
    $discos = ($_.Group | Where-Object { $_.EventID -eq 24 }).Count
    $recons = ($_.Group | Where-Object { $_.EventID -eq 25 }).Count
    [PSCustomObject]@{
        Date = $day
        Logons = $logons
        Desconexoes = $discos
        Reconexoes = $recons
    }
}

# HTML inicial
$htmlHeader = @"
<html>
<head>
<title>Relatório RDP (últimos 7 dias) - Geset</title>
<style>
body { font-family: Arial; margin: 20px; }
h1 { color: #2a3f5f; }
input { margin: 5px; padding: 5px; }
canvas { display: block; max-width: 100%; height: auto; }
table { border-collapse: collapse; width: 100%; margin-bottom: 40px; }
th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
th { background-color: #f2f2f2; }
</style>
</head>
<body>
<img src="images.jpg" alt="Logo Geset" style="height: 80px; margin-bottom: 20px;">
<h1>Relatório de Sessões RDP - Últimos 7 dias - Geset</h1>
"@

$labels = ($summary | ForEach-Object { "'$($_.Date)'" }) -join ","
$logons = ($summary | ForEach-Object { $_.Logons }) -join ","
$disc = ($summary | ForEach-Object { $_.Desconexoes }) -join ","
$recon = ($summary | ForEach-Object { $_.Reconexoes }) -join ","

$htmlChart = @"
<div style='width: 100%; max-width: 900px; height: 400px; margin: 0 auto 40px auto;'>
  <canvas id='myChart'></canvas>
</div>
<script src='https://cdn.jsdelivr.net/npm/chart.js'></script>
<script>
const ctx = document.getElementById('myChart').getContext('2d');
const myChart = new Chart(ctx, {
    type: 'bar',
    data: {
        labels: [$labels],
        datasets: [
            {
                label: 'Logons',
                data: [$logons],
                backgroundColor: 'rgba(75, 192, 192, 0.7)'
            },
            {
                label: 'Desconexões',
                data: [$disc],
                backgroundColor: 'rgba(255, 159, 64, 0.7)'
            },
            {
                label: 'Reconexões',
                data: [$recon],
                backgroundColor: 'rgba(153, 102, 255, 0.7)'
            }
        ]
    },
    options: {
        responsive: true,
        maintainAspectRatio: true,
        aspectRatio: 2,
        scales: {
            y: { beginAtZero: true }
        }
    }
});
</script>
"@

$htmlFilters = @"
<h2>Filtrar Detalhes</h2>
<label for='userFilter'>Filtrar por usuário:</label>
<input type='text' id='userFilter' onkeyup='filterTable()' placeholder='Digite o nome do usuário'>

<label for='dateFilter'>Filtrar por data:</label>
<input type='text' id='dateFilter' onkeyup='filterTable()' placeholder='Ex: 06/01/2025'>

<script>
function filterTable() {
    let userInput = document.getElementById('userFilter').value.toLowerCase();
    let dateInput = document.getElementById('dateFilter').value.toLowerCase();
    let rows = document.querySelectorAll("#details tbody tr");

    rows.forEach(row => {
        let user = row.cells[1].innerText.toLowerCase();
        let date = row.cells[0].innerText.toLowerCase();
        if (user.includes(userInput) && date.includes(dateInput)) {
            row.style.display = '';
        } else {
            row.style.display = 'none';
        }
    });
}
</script>
"@

$htmlDetail = "<h2>Detalhamento de Eventos</h2>"
$htmlDetail += "<table id='details'><thead><tr><th>Data/Hora</th><th>Usuário</th><th>Evento</th><th>SessionID</th><th>IP</th></tr></thead><tbody>"

foreach ($log in $logs | Sort-Object TimeCreated -Descending) {
    $htmlDetail += "<tr><td>$($log.TimeCreated)</td><td>$($log.User)</td><td>$($log.Description)</td><td>$($log.SessionID)</td><td>$($log.Address)</td></tr>"
}
$htmlDetail += "</tbody></table>"

$htmlFooter = "</body></html>"

$html = $htmlHeader + $htmlChart + $htmlFilters + $htmlDetail + $htmlFooter
$html | Out-File -Encoding UTF8 $outputPath
