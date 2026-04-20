<#
.SYNOPSIS
  HWPFrame.HwpObject COM의 AppID 레지스트리 항목을 'Interactive User'로 설정한다.

.DESCRIPTION
  원격 세션/Session 0에서 pyhwpx 사용 시 "서버 실행이 실패했습니다" 에러를 유발하는
  DCOM 라우팅 문제를 해결한다. HKLM 쓰기이므로 관리자 권한 필요. 멱등.

.NOTES
  한글 업데이트로 지워질 수 있다.
#>

$ErrorActionPreference = 'Stop'

$me = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($me)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error 'Administrator 권한이 필요합니다.'
    exit 1
}

$clsid = '{2291CF00-64A1-4877-A9B4-68CFE89612D6}'

$appid32 = "HKLM:\SOFTWARE\WOW6432Node\Classes\AppID\$clsid"
if (-not (Test-Path $appid32)) { New-Item -Path $appid32 -Force | Out-Null }
Set-ItemProperty -Path $appid32 -Name 'RunAs' -Value 'Interactive User' -Type String

$clsidKey32 = "HKLM:\SOFTWARE\WOW6432Node\Classes\CLSID\$clsid"
if (Test-Path $clsidKey32) {
    Set-ItemProperty -Path $clsidKey32 -Name 'AppID' -Value $clsid -Type String
}

$appid64 = "HKLM:\SOFTWARE\Classes\AppID\$clsid"
if (-not (Test-Path $appid64)) { New-Item -Path $appid64 -Force | Out-Null }
Set-ItemProperty -Path $appid64 -Name 'RunAs' -Value 'Interactive User' -Type String

Write-Host 'HWP AppID RunAs=Interactive User 설정 완료.' -ForegroundColor Green
exit 0
