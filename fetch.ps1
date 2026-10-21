# Сборщик тендеров для GNB Line.
# Берёт актуальные закупки (этап «подача заявок») по 44-ФЗ и 223-ФЗ из ГосПлан API,
# оставляет только подходящие по смыслу (ГНБ, прокол, бестраншейная прокладка)
# и сохраняет их в data/tenders.json. Этот файл читает страница index.html.
#
# Работает в Windows PowerShell 5.1 и в PowerShell 7 (GitHub Actions).
# Бесплатный тестовый сервер ГосПлан: 10 запросов в минуту, ключ не нужен.
# Если купите ключ (3 000 ₽/мес), задайте переменную окружения GOSPLAN_API_KEY —
# сборщик сам переключится на рабочий сервер и станет быстрее.

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Root     = Split-Path -Parent $MyInvocation.MyCommand.Path
$DataFile = Join-Path $Root 'data\tenders.json'
$ApiKey   = $env:GOSPLAN_API_KEY
if ($ApiKey) { $Base = 'https://v2.gosplan.info'; $Pause = 1 } else { $Base = 'https://v2test.gosplan.info'; $Pause = 7.5 }

# Слова для поиска в ГосПлан (поиск там идёт «по любому слову», поэтому слова короткие и характерные).
$SearchWords = @('ГНБ', 'прокол', 'проколом', 'бестраншейной', 'бестраншейная', 'горизонтально', 'продавливанием', 'микротоннелирование', 'футляре')

# Что считаем «нашей» работой (проверяем по названию закупки и её позициям).
$Relevant = '(?i)(\bГНБ\b|горизонтальн\w*[\s-]*направленн|наклонно[\s-]*направленн|бестраншейн|прокол|продавлив|микротоннел|закрыт\w* способ|методом бурения|управляем\w* бурени)'
# Явный мусор, который попадается по слову «прокол» и т.п.
$Junk = '(?i)(перчат|игл[аы]|катетер|куртк|комбинезон|обув|сапог|костюм|одежд|\bшин|покрышк|камер[аы] для|ткан|пункц|медицин|лекарств|биопс|стомат|ушей|пирсинг|ветеринар|протектор|аптечк|скважин\w* на воду|ГБУЗ|«ГНБ»|^\s*поставк|маркетингов\w* исследов|запасн\w* част|расходн\w* материал)'

$Utf8 = New-Object System.Text.UTF8Encoding($false)

function Invoke-Api([string]$Path) {
    $url = "$Base$Path"
    $headers = @{}
    if ($ApiKey) { $headers['apikey'] = $ApiKey; $headers['Authorization'] = "Bearer $ApiKey" }
    for ($try = 1; $try -le 4; $try++) {
        Start-Sleep -Milliseconds ([int]($Pause * 1000))
        try {
            $r = Invoke-WebRequest -Uri $url -Headers $headers -UseBasicParsing -TimeoutSec 60
            # Ответ всегда в UTF-8; декодируем сами, иначе PowerShell 5.1 портит кириллицу.
            $text = [System.Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray())
            return ($text | ConvertFrom-Json)
        } catch {
            $code = $null
            if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
            Write-Host "  ! $code $($_.Exception.Message) — попытка $try"
            if ($code -eq 404) { return $null }
            Start-Sleep -Seconds (20 * $try)
        }
    }
    return $null
}

function Enc([string]$s) { [uri]::EscapeDataString($s) }

function Get-Prop($obj, [string[]]$path) {
    $cur = $obj
    foreach ($p in $path) {
        if ($null -eq $cur) { return $null }
        $cur = $cur.$p
    }
    return $cur
}

function As-Array($x) { if ($null -eq $x) { @() } elseif ($x -is [array]) { $x } else { @($x) } }

# --- Кэш прошлых запусков: подробности по закупке грузим только один раз ---
$cache = @{}
if (Test-Path $DataFile) {
    try {
        $old = [IO.File]::ReadAllText($DataFile, [Text.Encoding]::UTF8) | ConvertFrom-Json
        foreach ($t in (As-Array $old.tenders)) { if ($t.detailsLoaded) { $cache[$t.id] = $t } }
    } catch { Write-Host 'Старый файл данных не прочитан, начинаем заново.' }
}

# --- 1. Поиск по словам ---
$found = @{}
foreach ($law in @('fz44', 'fz223')) {
    foreach ($w in $SearchWords) {
        $skip = 0
        do {
            Write-Host "Поиск $law «$w» (с $skip)"
            $page = As-Array (Invoke-Api "/$law/purchases?object_info=$(Enc $w)&stage=1&limit=100&skip=$skip&published_forpast=3m")
            foreach ($p in $page) {
                $text = "$($p.object_info)"
                if ($text -notmatch $Relevant -or $text -match $Junk) { continue }
                $id = "$law-$($p.purchase_number)"
                if (-not $found.ContainsKey($id)) { $found[$id] = @{ law = $law; row = $p } }
            }
            $skip += 100
        } while ($page.Count -eq 100 -and $skip -lt 500)
    }
}
Write-Host "Найдено подходящих: $($found.Count)"

# --- 2. Подробности по каждой закупке ---
$now = (Get-Date).ToUniversalTime()
$result = New-Object System.Collections.ArrayList
foreach ($id in $found.Keys) {
    $law = $found[$id].law
    $p   = $found[$id].row
    $num = "$($p.purchase_number)"
    $deadline = if ($law -eq 'fz44') { $p.collecting_finished_at } else { $p.submission_close_at }

    $t = [ordered]@{
        id            = $id
        law           = if ($law -eq 'fz44') { '44-ФЗ' } else { '223-ФЗ' }
        number        = $num
        title         = "$($p.object_info)"
        price         = [double]$p.max_price
        currency      = $p.currency_code
        region        = $p.region
        published     = $p.published_at
        deadline      = $deadline
        okpd2         = As-Array $p.okpd2
        contractGuaranteePart   = $p.contract_guarantee_part
        contractGuaranteeAmount = $p.contract_guarantee_amount
        url           = "https://zakupki.gov.ru/epz/order/extendedsearch/results.html?searchString=$num"
        detailsLoaded = $false
    }

    if ($cache.ContainsKey($id)) {
        $c = $cache[$id]
        foreach ($k in @('customer', 'regionText', 'method', 'platform', 'address', 'termDays', 'appGuaranteeAmount', 'appGuaranteePart', 'objects', 'docs', 'url', 'detailsLoaded')) {
            if ($null -ne $c.$k) { $t[$k] = $c.$k }
        }
    } else {
        Write-Host "Подробности $id"
        if ($law -eq 'fz44') {
            $n = Invoke-Api "/fz44/purchases/$num/notification"
            if ($n) {
                $s = $n.source
                $t.url      = if ($s.commonInfo.href) { $s.commonInfo.href } else { $t.url }
                $t.method   = Get-Prop $s @('commonInfo', 'placingWay', 'name')
                $t.platform = Get-Prop $s @('commonInfo', 'ETP', 'name')
                $t.customer = Get-Prop $s @('purchaseResponsibleInfo', 'responsibleOrgInfo', 'fullName')
                $req = (As-Array (Get-Prop $s @('notificationInfo', 'customerRequirementsInfo', 'customerRequirementInfo')))[0]
                if ($req) {
                    $t.appGuaranteeAmount = Get-Prop $req @('applicationGuarantee', 'amount')
                    $t.appGuaranteePart   = Get-Prop $req @('applicationGuarantee', 'part')
                    $cg = Get-Prop $req @('contractGuarantee', 'part')
                    if ($cg) { $t.contractGuaranteePart = [double]$cg }
                    $cga = Get-Prop $req @('contractGuarantee', 'amount')
                    if ($cga) { $t.contractGuaranteeAmount = [double]$cga }
                    $term = Get-Prop $req @('contractConditionsInfo', 'contractExecutionPaymentPlan', 'contractExecutionTermsInfo', 'relativeTermsInfo')
                    if ($term -and $term.term) {
                        $days = [int]$term.term
                        if ($term.termDayType -eq 'working') { $days = [int][math]::Ceiling($days * 7 / 5) }
                        $t.termDays = $days
                    }
                    $place = (As-Array (Get-Prop $req @('contractConditionsInfo', 'deliveryPlacesInfo', 'byGARInfo')))[0]
                    if ($place) { $t.address = (("$($place.GARInfo.GARAddress) $($place.deliveryPlace)").Trim()) }
                }
                $objs = @()
                foreach ($o in (As-Array (Get-Prop $s @('notificationInfo', 'purchaseObjectsInfo', 'notDrugPurchaseObjectsInfo', 'purchaseObject')))) {
                    $unit = if ($o.OKEI) { $o.OKEI.nationalCode } else { $null }
                    $qty  = if ($o.quantity -and $o.quantity.value) { [double]($o.quantity.value -replace ',', '.') } else { $null }
                    $objs += [ordered]@{ name = "$($o.name)"; unit = $unit; qty = $qty; sum = $o.sum }
                }
                $t.objects = $objs
                $docs = @()
                foreach ($a in (As-Array (Get-Prop $s @('attachmentsInfo', 'attachmentInfo')))) {
                    $docs += [ordered]@{ name = "$($a.fileName)"; url = "$($a.url)" }
                }
                $t.docs = $docs
                $t.detailsLoaded = $true
            }
        } else {
            $d = Invoke-Api "/fz223/purchases/$num"
            if ($d) {
                $s = (As-Array $d.docs | Where-Object { $_.source } | Select-Object -Last 1).source
                if ($s) {
                    if ($s.urlEIS) { $t.url = $t.url }  # ссылка urlEIS ведёт в личный кабинет, оставляем публичный поиск
                    $t.method     = $s.purchaseCodeName
                    $t.platform   = $s.urlVSRZ
                    $t.customer   = Get-Prop $s @('customer', 'mainInfo', 'fullName')
                    $t.regionText = Get-Prop $s @('customer', 'mainInfo', 'region')
                    $docs = @()
                    foreach ($a in (As-Array (Get-Prop $s @('attachments', 'document')))) {
                        if ($a.url) { $docs += [ordered]@{ name = "$($a.fileName)"; url = "$($a.url)" } }
                    }
                    $t.docs = $docs
                }
                $t.detailsLoaded = $true
            }
        }
    }
    [void]$result.Add($t)
}

$out = [ordered]@{
    updated = $now.ToString('yyyy-MM-ddTHH:mm:ssZ')
    source  = 'ГосПлан API (данные ЕИС zakupki.gov.ru), 44-ФЗ и 223-ФЗ, этап подачи заявок'
    tenders = @($result | Sort-Object { $_.deadline })
}
$dir = Split-Path -Parent $DataFile
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
[IO.File]::WriteAllText($DataFile, ($out | ConvertTo-Json -Depth 8), $Utf8)
Write-Host "Сохранено: $($result.Count) закупок -> $DataFile"
