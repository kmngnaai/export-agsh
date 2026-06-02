param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Step,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$CommitMessage
)

$ErrorActionPreference = "Stop"

# Luon chay tu thu muc goc repo de cac duong dan tuong doi hoat dong on dinh.
$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

$BackendPath = "00.Detaisublog_v26.py"

function Fail {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host "FAIL: $Message"
    exit 1
}

function Invoke-NativeStep {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Command
    )

    Write-Host ""
    Write-Host "RUN: $Name"

    try {
        & $Command
        if ($LASTEXITCODE -ne 0) {
            Fail "$Name that bai voi exit code $LASTEXITCODE."
        }
    }
    catch {
        Fail "$Name that bai: $_"
    }

    Write-Host "PASS: $Name"
}

Write-Host "RUN: Kiem tra git working tree sach"
try {
    $GitStatus = git status --porcelain
    if ($LASTEXITCODE -ne 0) {
        Fail "Khong kiem tra duoc git working tree."
    }
}
catch {
    Fail "Khong kiem tra duoc git working tree: $_"
}

# Khong cleanup khi repo dang co thay doi de tranh ghi de cong viec chua commit.
if ($GitStatus) {
    Fail "Git working tree khong sach. Hay xu ly thay doi hien co truoc."
}
Write-Host "PASS: Git working tree sach"

Invoke-NativeStep "Ap dung cleanup step '$Step'" {
    powershell -ExecutionPolicy Bypass -File .\tools\apply-cleanup-step.ps1 -Step $Step
}

Write-Host ""
Write-Host "RUN: Chay regression sau cleanup"
try {
    powershell -ExecutionPolicy Bypass -File .\tools\run-regression.ps1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "RUN: Khoi phuc $BackendPath do regression that bai"
        git restore -- $BackendPath
        if ($LASTEXITCODE -ne 0) {
            Fail "Regression that bai va khong the khoi phuc $BackendPath."
        }
        Fail "Regression that bai. Da khoi phuc $BackendPath."
    }
}
catch {
    Write-Host "RUN: Khoi phuc $BackendPath do regression that bai"
    git restore -- $BackendPath
    if ($LASTEXITCODE -ne 0) {
        Fail "Regression that bai va khong the khoi phuc $BackendPath."
    }
    Fail "Regression that bai. Da khoi phuc $BackendPath. Chi tiet: $_"
}
Write-Host "PASS: Regression sau cleanup"

# Chi commit khi cleanup thuc su lam thay doi backend.
Write-Host ""
Write-Host "RUN: Kiem tra thay doi o $BackendPath"
git diff --quiet -- $BackendPath
if ($LASTEXITCODE -eq 0) {
    Fail "Khong co thay doi o $BackendPath. Khong tao commit."
}
if ($LASTEXITCODE -ne 1) {
    Fail "Khong kiem tra duoc thay doi o $BackendPath."
}
Write-Host "PASS: Co thay doi o $BackendPath"

Invoke-NativeStep "Stage $BackendPath" {
    git add -- $BackendPath
}

Invoke-NativeStep "Commit cleanup" {
    git commit -m $CommitMessage
}

Invoke-NativeStep "Push commit" {
    git push
}

Invoke-NativeStep "Hien thi git status" {
    git status
}

Write-Host ""
Write-Host "SAFE CLEANUP COMMIT PASSED"
