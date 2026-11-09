# -*- coding: utf-8 -*-
<#
自签名代码签名脚本（Windows PowerShell 5.1+，无需 signtool、无需联网）

做什么：
  1) 生成（或复用）一张"代码签名"自签名证书（当前用户，3 年有效期）
  2) 导出 .pfx（签名用）与 .cer（发给用户安装信任用）
  3) 把证书加入当前用户的"受信任的根证书颁发机构"（可选，本机不再弹"未知发布者"）
  4) 对 dist 里的 EXE 签名（SHA256；时间戳服务器不通时自动跳过时间戳）
  5) 校验签名状态，并重新打包 zip、生成 SHA256SUMS.txt

用法（在项目根目录）：
  powershell -ExecutionPolicy Bypass -File scripts\sign_exe.ps1
  powershell -ExecutionPolicy Bypass -File scripts\sign_exe.ps1 -Subject "CN=XX单位 表格工具"
  powershell -ExecutionPolicy Bypass -File scripts\sign_exe.ps1 -NoTrust -NoRezip

移除信任（换证书或卸载时）：
  certutil -user -delstore Root <证书指纹>
#>
param(
    [string]$Subject  = "CN=杨子江 by DeepSeek HARNESS",
    [string]$Password = "TableTool@2026",
    [string]$PfxPath  = "",   # 默认 dist\codesign.pfx
    [string]$CerPath  = "",   # 默认 dist\codesign.cer
    [string]$ExePath  = "",   # 默认 dist\表格工具\表格工具.exe
    [switch]$NoTrust,         # 不把证书加入当前用户受信任根
    [switch]$NoTimestamp,     # 不加时间戳
    [switch]$NoRezip          # 签名后不重新打 zip
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not $Root) { $Root = (Get-Location).Path }
Set-Location $Root

if (-not $PfxPath) { $PfxPath = Join-Path $Root "dist\codesign.pfx" }
if (-not $CerPath) { $CerPath = Join-Path $Root "dist\codesign.cer" }
if (-not $ExePath) { $ExePath = Join-Path $Root "dist\表格工具\表格工具.exe" }

function Info($m) { Write-Host "[签名] $m" -ForegroundColor Cyan }
function Warn($m) { Write-Host "[注意] $m" -ForegroundColor Yellow }

if (-not (Test-Path -LiteralPath $ExePath)) {
    throw "找不到要签名的文件：$ExePath（请先运行 scripts\build_exe.py 打包）"
}

# ---------------------------------------------------------------- 1. 证书
$cert = $null
try {
    $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue |
        Where-Object { $_.Subject -eq $Subject } |
        Sort-Object NotAfter -Descending | Select-Object -First 1
} catch { }

if ($cert -and $cert.NotAfter -gt (Get-Date).AddDays(30)) {
    Info ("复用已有证书：{0}（{1} 到期）" -f $cert.Thumbprint, $cert.NotAfter.ToString("yyyy-MM-dd"))
} else {
    Info "生成新的代码签名证书：$Subject"
    $cert = New-SelfSignedCertificate `
        -Type CodeSigningCert `
        -Subject $Subject `
        -KeyUsage DigitalSignature `
        -KeyAlgorithm RSA -KeyLength 2048 `
        -NotAfter (Get-Date).AddYears(3) `
        -CertStoreLocation Cert:\CurrentUser\My
    Info ("已生成：{0}（{1} 到期）" -f $cert.Thumbprint, $cert.NotAfter.ToString("yyyy-MM-dd"))
}

# ---------------------------------------------------------------- 2. 导出 pfx / cer
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $PfxPath) | Out-Null
$sec = ConvertTo-SecureString -String $Password -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath $PfxPath -Password $sec | Out-Null
Export-Certificate -Cert $cert -FilePath $CerPath | Out-Null
Info ("已导出：{0}" -f $PfxPath)
Info ("已导出：{0}" -f $CerPath)

# ---------------------------------------------------------------- 3. 信任（当前用户根）
if (-not $NoTrust) {
    $already = $false
    try {
        $already = [bool](Get-ChildItem Cert:\CurrentUser\Root -ErrorAction SilentlyContinue |
            Where-Object { $_.Thumbprint -eq $cert.Thumbprint })
    } catch { }
    if ($already) {
        Info "证书已在「受信任的根证书颁发机构」，跳过"
    } else {
        Info "把证书加入当前用户「受信任的根证书颁发机构」…"
        $done = $false
        try {
            & certutil -user -addstore -f Root $CerPath | Out-Null
            $done = $LASTEXITCODE -eq 0
        } catch { }
        if (-not $done) {
            try {
                Import-Certificate -FilePath $CerPath -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
                $done = $true
            } catch { }
        }
        if ($done) { Info "已加入信任" } else { Warn "加入信任失败（可选步骤，不影响签名本身）" }
    }
}

# ---------------------------------------------------------------- 4. 签名
function Sign-File([string]$Path) {
    Info ("签名：{0}" -f $Path)
    $args = @{
        FilePath      = $Path
        Certificate   = $cert
        HashAlgorithm = "SHA256"
    }
    if (-not $NoTimestamp) { $args["TimestampServer"] = "http://timestamp.digicert.com" }
    try {
        $sig = Set-AuthenticodeSignature @args
    } catch {
        if (-not $NoTimestamp) {
            Warn "时间戳服务器不可用，改为无时间戳签名"
            $args.Remove("TimestampServer")
            $sig = Set-AuthenticodeSignature @args
        } else { throw }
    }
    if ($sig.Status -ne "Valid") {
        Warn ("签名状态：{0} {1}" -f $sig.Status, $sig.StatusMessage)
    } else {
        Info ("签名完成：{0}" -f $Path)
    }
}

Sign-File $ExePath
# 如有安装包/其它 exe，一并签（当前绿色版只有启动器）
$extra = Get-ChildItem (Join-Path $Root "dist") -Filter "*.exe" -File -ErrorAction SilentlyContinue
foreach ($e in $extra) {
    if ($e.FullName -ne (Resolve-Path -LiteralPath $ExePath).Path) { Sign-File $e.FullName }
}

# ---------------------------------------------------------------- 5. 校验 + 打包 + 校验和
Info "校验签名…"
$check = Get-AuthenticodeSignature -FilePath $ExePath
Write-Host ("  文件：{0}" -f $ExePath)
Write-Host ("  状态：{0}" -f $check.Status)
Write-Host ("  签名者：{0}" -f ($(if ($check.SignerCertificate) { $check.SignerCertificate.Subject } else { "（无）" })))
if ($check.SignerCertificate) {
    Write-Host ("  指纹：{0}" -f $check.SignerCertificate.Thumbprint)
    Write-Host ("  有效期至：{0}" -f $check.SignerCertificate.NotAfter)
}

if (-not $NoRezip) {
    $stage = Join-Path $Root "dist\表格工具"
    $zip = Join-Path $Root "dist\表格工具_v1.0_win64.zip"
    Info "重新打包：$zip"
    if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
    Compress-Archive -Path (Join-Path $stage "*") -DestinationPath $zip -CompressionLevel Optimal
    Info ("已重新打包：{0}（{1:N1} MB）" -f $zip, ((Get-Item $zip).Length / 1MB))
}

# SHA256 校验文件（zip 不能被 Authenticode 签，用哈希给用户比对）
$sums = Join-Path $Root "dist\SHA256SUMS.txt"
$lines = @()
$zipPath = Join-Path $Root "dist\表格工具_v1.0_win64.zip"
foreach ($p in @($ExePath, $zipPath)) {
    if (Test-Path -LiteralPath $p) {
        $h = (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash
        $lines += ("{0}  {1}" -f $h.ToLower(), (Split-Path -Leaf $p))
    }
}
$lines += ""
$lines += "证书指纹: " + $cert.Thumbprint
$lines += "证书主体: " + $cert.Subject
$lines += "有效期至: " + $cert.NotAfter
$lines | Set-Content -LiteralPath $sums -Encoding UTF8
Info ("已写入校验和：{0}" -f $sums)

Write-Host ""
Write-Host "完成。" -ForegroundColor Green
Write-Host "把下面两个文件一起发给用户："
Write-Host ("  1) {0}" -f (Join-Path $Root "dist\表格工具_v1.0_win64.zip"))
Write-Host ("  2) {0}（用户双击 → 安装证书 → 本地计算机 → 受信任的根证书颁发机构）" -f $CerPath)
Write-Host "本机已可直接运行，不再提示「未知发布者」。"
