[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$RepositoryRoot,
  [Parameter(Mandatory = $true)]
  [string]$EvidenceRoot
)

$ErrorActionPreference = 'Stop'
$env:PGPASSWORD = 'root'
$Psql = Join-Path $env:PGBIN 'psql.exe'
$CreateDb = Join-Path $env:PGBIN 'createdb.exe'
$DropDb = Join-Path $env:PGBIN 'dropdb.exe'
$ArtifactsRoot = Join-Path $RepositoryRoot 'artifacts'
$ScratchRoot = Join-Path $env:RUNNER_TEMP ('ale10030-' + [guid]::NewGuid().ToString('N'))
$ReportNames = @(
  'event_decision.csv',
  'claim_wave.csv',
  'final_dispatch_state.csv',
  'endpoint_summary.csv'
)

New-Item -ItemType Directory -Force -Path $EvidenceRoot | Out-Null
New-Item -ItemType Directory -Force -Path $ScratchRoot | Out-Null

function Assert-True {
  param([bool]$Condition, [string]$Message)
  if (-not $Condition) {
    throw $Message
  }
}

function Get-Sha256 {
  param([string]$PathValue)
  return (Get-FileHash -Algorithm SHA256 -LiteralPath $PathValue).Hash.ToLowerInvariant()
}

function Get-TreeHash {
  param([string]$Root)
  $lines = Get-ChildItem -LiteralPath $Root -File -Recurse |
    ForEach-Object {
      $relative = [System.IO.Path]::GetRelativePath($Root, $_.FullName).Replace('\', '/')
      $relative + [char]0 + (Get-Sha256 $_.FullName)
    } |
    Sort-Object
  $text = [string]::Join("`n", $lines)
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
  $hash = [System.Security.Cryptography.SHA256]::HashData($bytes)
  return [Convert]::ToHexString($hash).ToLowerInvariant()
}

function Get-ZipFileEntries {
  param([string]$ArchivePath)
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
  try {
    return @($archive.Entries |
      Where-Object { -not [string]::IsNullOrEmpty($_.Name) } |
      ForEach-Object { $_.FullName.Replace('\', '/') } |
      Sort-Object)
  } finally {
    $archive.Dispose()
  }
}

function Assert-SequenceEqual {
  param([object[]]$Actual, [object[]]$Expected, [string]$Label)
  $difference = Compare-Object -ReferenceObject $Expected -DifferenceObject $Actual
  if ($difference) {
    throw "$Label member list mismatch: $($difference | Out-String)"
  }
}

function Get-WorkbookSheetNames {
  param([string]$WorkbookPath)
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)
  try {
    $entry = $archive.GetEntry('xl/workbook.xml')
    Assert-True ($null -ne $entry) 'workbook.xml missing'
    $reader = [System.IO.StreamReader]::new($entry.Open())
    try {
      [xml]$xml = $reader.ReadToEnd()
    } finally {
      $reader.Dispose()
    }
    return @($xml.SelectNodes("//*[local-name()='sheet']") | ForEach-Object { $_.GetAttribute('name') })
  } finally {
    $archive.Dispose()
  }
}

function Get-WorkbookText {
  param([string]$WorkbookPath)
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)
  $values = [System.Collections.Generic.List[string]]::new()
  try {
    foreach ($entry in $archive.Entries) {
      if (-not $entry.FullName.StartsWith('xl/worksheets/') -or -not $entry.FullName.EndsWith('.xml')) {
        continue
      }
      $reader = [System.IO.StreamReader]::new($entry.Open())
      try {
        [xml]$xml = $reader.ReadToEnd()
      } finally {
        $reader.Dispose()
      }
      foreach ($cell in $xml.SelectNodes("//*[local-name()='c']")) {
        $cellType = $cell.GetAttribute('t')
        if ($cellType -eq 'str') {
          $node = $cell.SelectSingleNode("./*[local-name()='v']")
          if ($null -ne $node -and -not [string]::IsNullOrEmpty($node.InnerText)) {
            $values.Add($node.InnerText)
          }
        } elseif ($cellType -eq 'inlineStr') {
          $parts = $cell.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText }
          if ($parts.Count -gt 0) {
            $values.Add([string]::Join('', $parts))
          }
        }
      }
    }
  } finally {
    $archive.Dispose()
  }
  return @($values)
}

function Assert-NaturalText {
  param([string[]]$Texts, [string]$Label)
  $quoteCharacters = @(
    [char]34, [char]39, [char]96, '“', '”', '‘', '’', '＂', '＇',
    '「', '」', '『', '』', '«', '»', '‹', '›', '〝', '〞', '〟', '《', '》', '〈', '〉'
  )
  $space = '[\t \u00a0\u1680\u2000-\u200a\u202f\u205f\u3000]+'
  $han = '[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]'
  $boundary = "(?:$han$space[A-Za-z0-9]|[A-Za-z0-9]$space$han|[A-Za-z]$space[0-9]|[0-9]$space[A-Za-z])"
  $riskTerms = @('此外', '至关重要', '深入探讨', '彰显', '赋能', '无缝', '不断演变的格局', '不仅', '不只是', '值得注意的是', '专家认为', '行业报告显示', '观察者指出', '未来展望', '挑战与未来', '——')
  $processTerms = @('制题返修', '去AI', '修改题目', '规则调整', 'Windows复现', 'GitHub Actions', '双干净目录', '动态变化', '负例', '附件哈希', '飞书回读', 'record_id', 'file_token')
  foreach ($text in $Texts) {
    foreach ($character in $quoteCharacters) {
      Assert-True (-not $text.Contains([string]$character)) "$Label contains a forbidden quote"
    }
    Assert-True (-not [regex]::IsMatch($text, $boundary)) "$Label contains a mixed boundary space"
    foreach ($term in ($riskTerms + $processTerms)) {
      Assert-True (-not $text.Contains($term)) "$Label contains forbidden term $term"
    }
  }
}

function Expand-TaskWorkspace {
  param([string]$Name)
  $workspace = Join-Path $ScratchRoot $Name
  New-Item -ItemType Directory -Force -Path $workspace | Out-Null
  Expand-Archive -LiteralPath (Join-Path $ArtifactsRoot '输入数据包.zip') -DestinationPath $workspace
  Expand-Archive -LiteralPath (Join-Path $ArtifactsRoot 'reference.zip') -DestinationPath $workspace
  return $workspace
}

function Save-ExpectedReports {
  param([string]$Workspace)
  $expectedRoot = Join-Path $Workspace 'expected_reports'
  New-Item -ItemType Directory -Force -Path $expectedRoot | Out-Null
  foreach ($name in $ReportNames) {
    $source = Join-Path $Workspace "output/reports/$name"
    Copy-Item -LiteralPath $source -Destination (Join-Path $expectedRoot $name)
    Remove-Item -LiteralPath $source
  }
  return $expectedRoot
}

function New-TestDatabase {
  $database = 'ale10030_' + [guid]::NewGuid().ToString('N').Substring(0, 12)
  & $CreateDb -h 127.0.0.1 -p 5432 -U postgres $database | Out-Host
  if ($LASTEXITCODE -ne 0) {
    throw 'createdb failed'
  }
  return $database
}

function Remove-TestDatabase {
  param([string]$Database)
  & $DropDb --if-exists -h 127.0.0.1 -p 5432 -U postgres $Database | Out-Host
}

function Invoke-Psql {
  param([string]$Database, [string[]]$Arguments)
  & $Psql -X -v ON_ERROR_STOP=1 -h 127.0.0.1 -p 5432 -U postgres -d $Database @Arguments | Out-Host
  if ($LASTEXITCODE -ne 0) {
    throw "psql failed with exit code $LASTEXITCODE"
  }
}

function Get-ReferenceStaticHashes {
  param([string]$Workspace)
  $hashes = [ordered]@{}
  foreach ($relative in @('output/run.ps1', 'output/sql/schema.sql', 'output/sql/outbox_runtime.sql')) {
    $hashes[$relative] = Get-Sha256 (Join-Path $Workspace $relative)
  }
  return $hashes
}

function Assert-ReferenceStaticHashes {
  param([string]$Workspace, [object]$Expected)
  foreach ($relative in $Expected.Keys) {
    Assert-True ((Get-Sha256 (Join-Path $Workspace $relative)) -eq $Expected[$relative]) "$relative changed during execution"
  }
}

function Get-CsvSemanticRows {
  param([string]$CsvPath)
  $rows = Import-Csv -LiteralPath $CsvPath
  return @($rows | ForEach-Object { $_ | ConvertTo-Json -Compress } | Sort-Object)
}

function Compare-Reports {
  param([string]$Workspace, [string]$ExpectedRoot)
  $hashes = [ordered]@{}
  foreach ($name in $ReportNames) {
    $actualPath = Join-Path $Workspace "output/reports/$name"
    $expectedPath = Join-Path $ExpectedRoot $name
    Assert-True (Test-Path -LiteralPath $actualPath) "$name missing"
    $actualRows = Get-CsvSemanticRows $actualPath
    $expectedRows = Get-CsvSemanticRows $expectedPath
    Assert-SequenceEqual $actualRows $expectedRows $name
    $hashes[$name] = Get-Sha256 $actualPath
  }
  return $hashes
}

function Invoke-CleanRun {
  param([string]$Name)
  $workspace = Expand-TaskWorkspace $Name
  $expectedRoot = Save-ExpectedReports $workspace
  $staticHashes = Get-ReferenceStaticHashes $workspace
  $inputRoot = Join-Path $workspace 'input_data'
  $beforeHash = Get-TreeHash $inputRoot
  $database = New-TestDatabase
  try {
    & (Join-Path $workspace 'output/run.ps1') `
      -InputRoot $inputRoot `
      -DatabaseName $database `
      -PgHostName 127.0.0.1 `
      -PgPort 5432 `
      -PgUser postgres `
      -PgPassword root `
      -PsqlPath $Psql | Out-Host
    if ($LASTEXITCODE -ne 0) {
      throw 'run.ps1 returned a nonzero exit code'
    }
  } finally {
    Remove-TestDatabase $database
  }
  $afterHash = Get-TreeHash $inputRoot
  Assert-True ($beforeHash -eq $afterHash) 'input files changed during execution'
  Assert-ReferenceStaticHashes $workspace $staticHashes
  return [ordered]@{
    directory_name = $Name
    input_tree_sha256 = $beforeHash
    report_sha256 = Compare-Reports $workspace $expectedRoot
    exit_code = 0
  }
}

function Invoke-MutationRun {
  $workspace = Expand-TaskWorkspace '规则变化 中文 空格'
  Save-ExpectedReports $workspace | Out-Null
  $policyPath = Join-Path $workspace 'input_data/data/endpoint_policy.csv'
  $source = Get-Content -LiteralPath $policyPath -Raw
  $changed = $source.Replace('webhook-risk,risk.alert,true,1,5|20,webhook', 'webhook-risk,risk.alert,true,1,5|25,webhook')
  Assert-True ($changed -ne $source) 'mutation target not found'
  Set-Content -LiteralPath $policyPath -Value $changed -Encoding utf8NoBOM -NoNewline
  $database = New-TestDatabase
  try {
    & (Join-Path $workspace 'output/run.ps1') `
      -InputRoot (Join-Path $workspace 'input_data') `
      -DatabaseName $database `
      -PgHostName 127.0.0.1 `
      -PgPort 5432 `
      -PgUser postgres `
      -PgPassword root `
      -PsqlPath $Psql | Out-Host
    if ($LASTEXITCODE -ne 0) {
      throw 'mutation run returned a nonzero exit code'
    }
  } finally {
    Remove-TestDatabase $database
  }
  $row = Import-Csv -LiteralPath (Join-Path $workspace 'output/reports/final_dispatch_state.csv') |
    Where-Object { $_.endpoint_id -eq 'webhook-risk' -and $_.dedupe_key -eq 'risk:9201' }
  Assert-True ($row.next_due_at_utc -eq '2026-07-30T10:26:00Z') 'mutation did not change retry time'
  return [ordered]@{
    rule = 'webhook-risk第二次退避由20分钟改为25分钟'
    expected_effect = 'risk:9201的next_due_at_utc变为2026-07-30T10:26:00Z'
    observed = $true
    exit_code = 0
  }
}

function Invoke-NegativeRun {
  $workspace = Expand-TaskWorkspace '无效输入 中文 空格'
  Save-ExpectedReports $workspace | Out-Null
  Remove-Item -LiteralPath (Join-Path $workspace 'input_data/data/events.csv')
  $database = New-TestDatabase
  $failed = $false
  try {
    try {
      & (Join-Path $workspace 'output/run.ps1') `
        -InputRoot (Join-Path $workspace 'input_data') `
        -DatabaseName $database `
        -PgHostName 127.0.0.1 `
        -PgPort 5432 `
        -PgUser postgres `
        -PgPassword root `
        -PsqlPath $Psql | Out-Host
    } catch {
      $failed = $true
    }
  } finally {
    Remove-TestDatabase $database
  }
  $residual = @($ReportNames | Where-Object { Test-Path -LiteralPath (Join-Path $workspace "output/reports/$_") })
  Assert-True $failed 'missing events.csv did not fail'
  Assert-True ($residual.Count -eq 0) 'negative run left report files'
  return [ordered]@{
    case = '缺少events.csv'
    nonzero_exit = $true
    residual_reports = $residual
    pass = $true
  }
}

function Initialize-ConcurrencyDatabase {
  param([string]$Workspace, [string]$Database)
  Push-Location (Join-Path $Workspace 'output/sql')
  try {
    Invoke-Psql $Database @('-f', 'schema.sql')
    Invoke-Psql $Database @('-f', 'outbox_runtime.sql')
  } finally {
    Pop-Location
  }
  $copies = @(
    @('dispatch.source_event(event_id,topic,dedupe_key,created_at_utc,priority,payload_valid,email_opt_in,payload_bytes)', 'events.csv'),
    @('dispatch.source_endpoint_policy(endpoint_id,topic,active,max_in_flight,retry_profile_text,endpoint_kind)', 'endpoint_policy.csv'),
    @('dispatch.source_existing_state(endpoint_id,dedupe_key,state,attempt_no,lease_owner,lease_until_utc,next_due_at_utc,provider_message_id)', 'existing_dispatch_state.csv'),
    @('dispatch.source_delivery_result(wave_id,worker_id,endpoint_id,dedupe_key,delivered_at_utc,outcome,status_code,provider_message_id)', 'delivery_results.csv')
  )
  Push-Location (Join-Path $Workspace 'input_data/data')
  try {
    foreach ($copy in $copies) {
      Invoke-Psql $Database @('-c', "\copy $($copy[0]) FROM '$($copy[1])' CSV HEADER")
    }
  } finally {
    Pop-Location
  }
  Invoke-Psql $Database @('-c', 'SELECT dispatch.prepare_outbox();')
}

function Invoke-ConcurrencyProbe {
  $workspace = Expand-TaskWorkspace '并发会话 中文 空格'
  Save-ExpectedReports $workspace | Out-Null
  $database = New-TestDatabase
  $lockFile = Join-Path $env:RUNNER_TEMP ('ale-lock-' + [guid]::NewGuid().ToString('N') + '.sql')
  $lockOut = $lockFile + '.out'
  $lockErr = $lockFile + '.err'
  $lockProcess = $null
  try {
    Initialize-ConcurrencyDatabase $workspace $database
    @(
      "SET application_name='ale_lock_holder';",
      'BEGIN;',
      "SELECT endpoint_id,dedupe_key FROM dispatch.dispatch_job WHERE endpoint_id='webhook-risk' AND dedupe_key='risk:9201' FOR UPDATE;",
      'SELECT pg_sleep(15);',
      'COMMIT;'
    ) | Set-Content -LiteralPath $lockFile -Encoding utf8NoBOM
    $arguments = @('-X', '-v', 'ON_ERROR_STOP=1', '-h', '127.0.0.1', '-p', '5432', '-U', 'postgres', '-d', $database, '-f', $lockFile)
    $lockProcess = Start-Process -FilePath $Psql -ArgumentList $arguments -PassThru -RedirectStandardOutput $lockOut -RedirectStandardError $lockErr

    $ready = $false
    for ($attempt = 0; $attempt -lt 60; $attempt += 1) {
      $active = & $Psql -X -A -t -h 127.0.0.1 -p 5432 -U postgres -d $database -c "SELECT count(*) FROM pg_stat_activity WHERE application_name='ale_lock_holder' AND query LIKE 'SELECT pg_sleep%';"
      if ($LASTEXITCODE -eq 0 -and $active.Trim() -eq '1') {
        $ready = $true
        break
      }
      Start-Sleep -Milliseconds 250
    }
    Assert-True $ready 'lock holder did not reach the sleep state'

    $claimed = & $Psql -X -v ON_ERROR_STOP=1 -A -t -F '|' -h 127.0.0.1 -p 5432 -U postgres -d $database -c "SELECT endpoint_id,dedupe_key FROM dispatch.claim_dispatch_batch('C2','worker-c2',1,'2026-07-30T10:00:00Z');"
    Assert-True ($LASTEXITCODE -eq 0) 'concurrent claimant failed'
    Assert-True (($claimed -join "`n").Contains('webhook-risk|risk:9202')) 'concurrent claimant did not skip the locked highest-priority row'

    Wait-Process -Id $lockProcess.Id
    $lockProcess.Refresh()
    Assert-True ($lockProcess.ExitCode -eq 0) 'lock holder failed'
    return [ordered]@{
      held_job = 'webhook-risk与risk:9201'
      concurrent_claim = 'webhook-risk与risk:9202'
      skip_locked_observed = $true
      holder_exit_code = $lockProcess.ExitCode
      claimant_exit_code = 0
    }
  } finally {
    if ($null -ne $lockProcess -and -not $lockProcess.HasExited) {
      Stop-Process -Id $lockProcess.Id -Force
    }
    Remove-TestDatabase $database
    foreach ($file in @($lockFile, $lockOut, $lockErr)) {
      if (Test-Path -LiteralPath $file) {
        Remove-Item -LiteralPath $file
      }
    }
  }
}

try {
  $manifest = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'manifest.json') -Raw | ConvertFrom-Json
  $attachmentHashes = [ordered]@{}
  foreach ($property in $manifest.attachments.PSObject.Properties) {
    $actual = Get-Sha256 (Join-Path $ArtifactsRoot $property.Name)
    Assert-True ($actual -eq $property.Value) "$($property.Name) hash mismatch"
    $attachmentHashes[$property.Name] = $actual
  }

  Assert-SequenceEqual (Get-ZipFileEntries (Join-Path $ArtifactsRoot '输入数据包.zip')) @(
    'input_data/README.md',
    'input_data/data/delivery_results.csv',
    'input_data/data/endpoint_policy.csv',
    'input_data/data/events.csv',
    'input_data/data/existing_dispatch_state.csv',
    'input_data/rules/outbox_contract.md',
    'input_data/starter_sql/outbox_starter.sql'
  ) 'input archive'
  Assert-SequenceEqual (Get-ZipFileEntries (Join-Path $ArtifactsRoot 'reference.zip')) @(
    'output/reports/claim_wave.csv',
    'output/reports/endpoint_summary.csv',
    'output/reports/event_decision.csv',
    'output/reports/final_dispatch_state.csv',
    'output/run.ps1',
    'output/sql/outbox_runtime.sql',
    'output/sql/schema.sql'
  ) 'reference archive'

  Assert-SequenceEqual (Get-WorkbookSheetNames (Join-Path $ArtifactsRoot '关键标准答案.xlsx')) @(
    '交付物答案清单',
    '固定字段答案',
    '固定集合答案',
    '固定数值答案',
    '允许变体答案'
  ) 'answer workbook sheets'
  Assert-SequenceEqual (Get-WorkbookSheetNames (Join-Path $ArtifactsRoot '任务规格转化.xlsx')) @('任务规格转化') 'specification workbook sheets'

  $textValues = [System.Collections.Generic.List[string]]::new()
  foreach ($file in Get-ChildItem -LiteralPath (Join-Path $RepositoryRoot 'task') -File) {
    $textValues.Add((Get-Content -LiteralPath $file.FullName -Raw))
  }
  $scanWorkspace = Expand-TaskWorkspace '文字扫描 中文 空格'
  $textValues.Add((Get-Content -LiteralPath (Join-Path $scanWorkspace 'input_data/README.md') -Raw))
  $textValues.Add((Get-Content -LiteralPath (Join-Path $scanWorkspace 'input_data/rules/outbox_contract.md') -Raw))
  foreach ($workbookName in @('关键标准答案.xlsx', '任务规格转化.xlsx')) {
    foreach ($value in Get-WorkbookText (Join-Path $ArtifactsRoot $workbookName)) {
      $textValues.Add($value)
    }
  }
  Assert-NaturalText @($textValues) 'task natural language'

  $humanizer = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'qa/humanizer_review.json') -Raw | ConvertFrom-Json
  Assert-True ($humanizer.pass -eq $true) 'humanizer review did not pass'
  Assert-True ($humanizer.scores.total -ge 45) 'humanizer score below 45'
  Assert-True ($humanizer.online_ai_detection_used_as_conclusion -eq $false) 'online AI field was used as a conclusion'

  $runOne = Invoke-CleanRun '第一次 中文 空格'
  $runTwo = Invoke-CleanRun '第二次 中文 空格'
  Assert-True (($runOne.report_sha256 | ConvertTo-Json -Compress) -eq ($runTwo.report_sha256 | ConvertTo-Json -Compress)) 'clean runs differ'

  $evidence = [ordered]@{
    task_asset_id = $manifest.task_asset_id
    commit_sha = $env:GITHUB_SHA
    workflow_run_id = $env:GITHUB_RUN_ID
    runner_os = $env:RUNNER_OS
    runner_image = $env:ImageOS
    runner_image_version = $env:ImageVersion
    powershell_version = $PSVersionTable.PSVersion.ToString()
    psql_version = (& $Psql --version).Trim()
    attachment_sha256 = $attachmentHashes
    natural_language_gate = [ordered]@{
      reviewed_text_values = $textValues.Count
      quote_hits = 0
      boundary_space_hits = 0
      humanizer_risk_hits = 0
      humanizer_score = $humanizer.scores.total
      online_ai_field_used_as_conclusion = $false
      pass = $true
    }
    clean_runs = @($runOne, $runTwo)
    mutation = (Invoke-MutationRun)
    negative = (Invoke-NegativeRun)
    concurrency = (Invoke-ConcurrencyProbe)
    pass = $true
  }
  $evidence | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'evidence.json') -Encoding utf8NoBOM
  & $Psql --version | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'software-version.txt') -Encoding utf8NoBOM
  Write-Host ($evidence | ConvertTo-Json -Depth 12)
} catch {
  [ordered]@{
    commit_sha = $env:GITHUB_SHA
    workflow_run_id = $env:GITHUB_RUN_ID
    pass = $false
    error = $_.Exception.Message
  } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'evidence.json') -Encoding utf8NoBOM
  throw
} finally {
  if (Test-Path -LiteralPath $ScratchRoot) {
    Remove-Item -LiteralPath $ScratchRoot -Recurse -Force
  }
}
