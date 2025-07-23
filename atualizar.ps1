# Requer execução como administrador
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Este script precisa ser executado como administrador." -ForegroundColor Red
    Write-Host "Tentando reiniciar como administrador..." -ForegroundColor Yellow
    Start-Process pwsh "-NoProfile -ExecutionPolicy Bypass -Command `"$($MyInvocation.MyCommand.Definition)`"" -Verb RunAs
    Exit
}

# Verificar política de execução
if ((Get-ExecutionPolicy -Scope CurrentUser) -eq 'Restricted') {
    Write-Host "A política de execução do PowerShell está restrita. Configurando para Bypass..." -ForegroundColor Yellow
    Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
}

$global:errors = @()
function Log-Error { param($msg) $global:errors += "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $msg"; Write-Host "ERRO: $msg" -ForegroundColor Red }

# Verificar internet
if (-not (Test-Connection 8.8.8.8 -Count 2 -Quiet)) {
    Log-Error "Sem conexão com internet. Verifique sua conexão e tente novamente."
    Write-Host "Fechando em 10 segundos..." -ForegroundColor Yellow
    Start-Sleep 10
    exit
}

Write-Host "Atualizando WinGet..." -ForegroundColor Cyan
try { winget upgrade --id Microsoft.Winget.Source --silent --accept-source-agreements --accept-package-agreements --disable-interactivity 2>&1 | Out-Null } catch { Log-Error "WinGet: $($_.Exception.Message)" }

Write-Host "Atualizando Chocolatey..." -ForegroundColor Cyan
try { choco upgrade chocolatey -y --no-progress 2>&1 | Out-Null } catch { Log-Error "Chocolatey: $($_.Exception.Message)" }

Write-Host "Verificando atualizações da Microsoft Store..." -ForegroundColor Cyan
try {
    $result = Get-CimInstance -Namespace "Root\cimv2\mdm\dmmap" -ClassName "MDM_EnterpriseModernAppManagement_AppManagement01" |
              Invoke-CimMethod -MethodName "UpdateScanMethod"
    if ($result.ReturnValue -ne 0) {
        Log-Error "Falha ao verificar atualizações da Microsoft Store. Código: $($result.ReturnValue)"
    }
} catch {
    Log-Error "Microsoft Store: $($_.Exception.Message)"
}

Write-Host "Verificando Office 365..." -ForegroundColor Cyan
$officeProcesses = "outlook", "winword", "excel", "powerpnt", "msaccess", "mspub", "onenote"
try {
    Get-Process -Name $officeProcesses -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3

    $officePath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun"
    if (Test-Path "$officePath\OfficeC2RClient.exe") {
        $process = Start-Process -FilePath "$officePath\OfficeC2RClient.exe" -ArgumentList "/update user" -Wait -NoNewWindow -PassThru
        if ($process.ExitCode -ne 0) {
            Log-Error "Office 365 falhou com código $($process.ExitCode)"
        }
    } else {
        Log-Error "OfficeC2RClient não encontrado. Office 365 pode não estar instalado."
    }
} catch {
    Log-Error "Office 365: $($_.Exception.Message)"
}

$browsers = @(
    @{Name="Chrome"; Process="chrome"; WingetId="Google.Chrome"; ChocoId="googlechrome"},
    @{Name="Brave"; Process="brave"; WingetId="Brave.Brave"; ChocoId="brave"},
    @{Name="Edge"; Process="msedge"; WingetId="Microsoft.Edge"; ChocoId=$null}
)

foreach ($b in $browsers) {
    Write-Host "Atualizando $($b.Name)..." -ForegroundColor Cyan
    try {
        Get-Process -Name $b.Process -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2

        if ($b.WingetId -and (Get-Command winget -ErrorAction SilentlyContinue)) {
            winget upgrade --id $b.WingetId --silent --accept-package-agreements --accept-source-agreements --disable-interactivity --timeout 300 2>&1 | Out-Null
        }
        if ($b.ChocoId -and (Get-Command choco -ErrorAction SilentlyContinue)) {
            choco upgrade $b.ChocoId -y --no-progress 2>&1 | Out-Null
        }
    } catch {
        Log-Error "Navegador $($b.Name): $($_.Exception.Message)"
    }
}

Write-Host "Verificando atualizações do PowerShell..." -ForegroundColor Cyan
try {
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        winget upgrade --id Microsoft.PowerShell --silent --accept-source-agreements --accept-package-agreements --disable-interactivity 2>&1 | Out-Null
    }
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        choco upgrade powershell-core -y --no-progress 2>&1 | Out-Null
    }
} catch {
    Log-Error "PowerShell: $($_.Exception.Message)"
}

Write-Host "Atualizando módulos do PowerShell..." -ForegroundColor Cyan
try {
    $modules = Get-InstalledModule | Where-Object { $_.Repository -eq 'PSGallery' }
    foreach ($module in $modules) {
        try {
            Update-Module -Name $module.Name -Force
        } catch {
            Log-Error "Módulo PowerShell $($module.Name): $($_.Exception.Message)"
        }
    }
} catch {
    Log-Error "Erro ao obter módulos instalados: $($_.Exception.Message)"
}

Write-Host "Atualizando WSL..." -ForegroundColor Cyan
if (Get-Command wsl -ErrorAction SilentlyContinue) {
    wsl --update

    Write-Host "Atualizando distros Linux e pacotes..." -ForegroundColor Cyan
    try {
        $distrosOutput = wsl --list --verbose
        $distros = $distrosOutput | Select-Object -Skip 1 | ForEach-Object {
            if ($_ -match '^\s*(\*?)\s*(\S+)\s+\S+\s+\d+$') {
                $name = $matches[2]
                if ($name -notmatch 'docker|windows') {
                    $name
                }
            }
        }

        foreach ($distro in $distros) {
            if (-not [string]::IsNullOrWhiteSpace($distro)) {
                Write-Host "Atualizando distro: $distro"
                wsl -d $distro -u root -e sh -c "
                    if command -v apt &> /dev/null; then
                        apt update > /dev/null 2>&1
                        apt upgrade -y > /dev/null 2>&1
                        apt autoremove -y > /dev/null 2>&1
                    elif command -v dnf &> /dev/null; then
                        dnf upgrade -y > /dev/null 2>&1
                        dnf autoremove -y > /dev/null 2>&1
                    elif command -v yum &> /dev/null; then
                        yum update -y > /dev/null 2>&1
                        yum autoremove -y > /dev/null 2>&1
                    elif command -v pacman &> /dev/null; then
                        pacman -Syu --noconfirm > /dev/null 2>&1
                        pacman -Qdtq | pacman -Rs - --noconfirm > /dev/null 2>&1
                    elif command -v apk &> /dev/null; then
                        apk update > /dev/null 2>&1
                        apk upgrade > /dev/null 2>&1
                        apk cache clean > /dev/null 2>&1
                    fi
                " 2>&1 | Out-Null
                if ($LASTEXITCODE -ne 0) {
                    Log-Error "Distro ${distro}: falha na atualização (código $LASTEXITCODE)"
                }
            }
        }
    } catch {
        Log-Error "Erro ao atualizar WSL: $($_.Exception.Message)"
    }
}

$logPath = "$PWD\fail.log"
if ($global:errors.Count -gt 0) {
    $global:errors | Out-File $logPath -Encoding UTF8
    Write-Host "Concluído com erros. Verifique $logPath para detalhes." -ForegroundColor Yellow
} else {
    "Sucesso $(Get-Date)" | Out-File $logPath -Encoding UTF8
    Write-Host "Concluído sem erros!" -ForegroundColor Green
}

Write-Host "Fechando em 10 segundos..."
Start-Sleep 10