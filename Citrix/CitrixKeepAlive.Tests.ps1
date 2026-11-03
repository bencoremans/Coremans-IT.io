param([string]$ScriptPath = (Join-Path $PSScriptRoot 'CitrixKeepAlive.ps1'))
$ErrorActionPreference = 'Stop'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }

# Load only pure detection helpers; do not run the tray application or touch registry.
foreach ($name in 'Select-CitrixComRegistration', 'Get-CitrixPowerShellPath') {
    $functionAst = $ast.Find({ param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    }, $true)
    . ([scriptblock]::Create($functionAst.Extent.Text))
}
function Assert-Equal($Actual, $Expected, $Label) {
    if ($Actual -ne $Expected) { throw "$Label : expected '$Expected', got '$Actual'" }
}
function Assert-Throws([scriptblock]$Action, $Label) {
    $threw = $false
    try { & $Action | Out-Null } catch { $threw = $true }
    if (-not $threw) { throw "$Label : expected a terminating error" }
}
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString())
try {
    foreach ($dir in 'x86', 'x64', 'broken') {
        $path = Join-Path $testRoot $dir
        New-Item -ItemType Directory -Path $path -Force | Out-Null
        New-Item -ItemType File -Path (Join-Path $path 'wfica.ocx') | Out-Null
        if ($dir -ne 'broken') { New-Item -ItemType File -Path (Join-Path $path 'WFICALib.dll') | Out-Null }
    }
    $x86 = [pscustomobject]@{ View = 'Registry32'; ServerPath = (Join-Path $testRoot 'x86/wfica.ocx') }
    $x64 = [pscustomobject]@{ View = 'Registry64'; ServerPath = (Join-Path $testRoot 'x64/wfica.ocx') }
    $broken = [pscustomobject]@{ View = 'Registry32'; ServerPath = (Join-Path $testRoot 'broken/wfica.ocx') }
    Assert-Equal (Select-CitrixComRegistration @($x64) 'Registry32').View 'Registry64' '32-bit host with only 64-bit Citrix'
    Assert-Equal (Select-CitrixComRegistration @($x86) 'Registry64').View 'Registry32' '64-bit host with only 32-bit Citrix'
    foreach ($view in 'Registry32', 'Registry64') {
        Assert-Equal (Select-CitrixComRegistration @($x86, $x64) $view).View $view 'Both registrations preserve current host'
    }
    Assert-Equal (Select-CitrixComRegistration @($broken, $x64) 'Registry32').View 'Registry64' 'Broken matching installation is skipped'
    Assert-Throws { Select-CitrixComRegistration @() 'Registry64' } 'Missing registration'
    Assert-Throws { Select-CitrixComRegistration @($broken) 'Registry32' } 'Missing adjacent DLL'
    Remove-Item -LiteralPath $x86.ServerPath
    Assert-Throws { Select-CitrixComRegistration @($x86) 'Registry32' } 'Stale OCX path'
    foreach ($case in @(
        @('Registry64', $true, $false, 'Sysnative'),
        @('Registry64', $true, $true, 'System32'),
        @('Registry32', $true, $true, 'SysWOW64'),
        @('Registry32', $false, $false, 'System32')
    )) {
        $path = Get-CitrixPowerShellPath $case[0] $case[1] $case[2] $testRoot
        $expected = Join-Path $testRoot "$($case[3])/WindowsPowerShell/v1.0/powershell.exe"
        Assert-Equal ($path.Replace('\', '/')) ($expected.Replace('\', '/')) 'PowerShell host path'
    }
    Assert-Throws { Get-CitrixPowerShellPath 'Registry64' $false $false $testRoot } '64-bit target on 32-bit OS'

    # Execute the actual background job for missing and invalid interop assemblies.
    $jobCommand = $ast.Find({ param($node)
        $node -is [System.Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Start-Job'
    }, $true)
    $jobBlock = $jobCommand.CommandElements | Where-Object { $_ -is [System.Management.Automation.Language.ScriptBlockExpressionAst] } | Select-Object -First 1
    $body = $jobBlock.ScriptBlock.GetScriptBlock()
    foreach ($dir in 'broken', 'x64') {
        $job = Start-Job -ScriptBlock $body -ArgumentList 15000, 126, (Join-Path $testRoot 'test.log'), 'Info', 'unused', (Join-Path $testRoot $dir)
        try {
            $finished = Wait-Job $job -Timeout 20
            if (-not $finished) { throw 'Background failure was not propagated within 20 seconds.' }
            Assert-Equal $job.State 'Failed' "Background job failure for $dir"
        } finally {
            Stop-Job $job
            Remove-Job $job
        }
    }
    Write-Output 'PASS: syntax, architecture selection, stale/missing files, host paths, and fatal job propagation.'
} finally {
    Remove-Item -LiteralPath $testRoot -Recurse -Force
}
