param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Step
)

$ErrorActionPreference = "Stop"

# Luôn chạy từ thư mục gốc repo để các đường dẫn tương đối ổn định.
$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

$BackendPath = "00.Detaisublog_v26.py"
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Fail {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host "FAIL: $Message"
    exit 1
}

function Restore-BackendAndFail {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    git restore -- $BackendPath
    Write-Host "FAIL: $Message"
    Write-Host "RESTORED: $BackendPath"
    exit 1
}

function Invoke-CleanupTest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Command
    )

    try {
        & $Command
        if ($LASTEXITCODE -ne 0) {
            Restore-BackendAndFail "$Name thất bại với exit code $LASTEXITCODE."
        }
    }
    catch {
        Restore-BackendAndFail "$Name thất bại: $_"
    }
}

function Count-LiteralOccurrences {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$Needle
    )

    $Count = 0
    $StartIndex = 0

    while ($true) {
        $Index = $Text.IndexOf($Needle, $StartIndex, [System.StringComparison]::Ordinal)
        if ($Index -lt 0) {
            return $Count
        }

        $Count += 1
        $StartIndex = $Index + $Needle.Length
    }
}

# Chỉ cho phép các cleanup nhỏ đã được phân tích trước.
$AllowedSteps = @(
    "remove-v72-self-assignment-u88",
    "remove-v72-self-assignment-u90"
)

if ($AllowedSteps -notcontains $Step) {
    Fail "Step '$Step' không nằm trong whitelist."
}

# Không sửa khi repo đang có thay đổi chưa commit để tránh ghi đè công việc khác.
$GitStatus = git status --porcelain
if ($LASTEXITCODE -ne 0) {
    Fail "Không kiểm tra được git working tree."
}
if ($GitStatus) {
    Fail "Git working tree không sạch. Hãy xử lý thay đổi hiện có trước."
}

$Text = [System.IO.File]::ReadAllText($BackendPath, $Utf8NoBom)
$NewLine = "`n"
if ($Text.Contains("`r`n")) {
    $NewLine = "`r`n"
}

switch ($Step) {
    "remove-v72-self-assignment-u88" {
        $TargetBlock = [string]::Join($NewLine, @(
            "# Force late references to use cache-enabled builder",
            "# Legacy self-assignment candidate: this assignment is a no-op.",
            "# Keep the cache-enabled function definition above until regression-tested.",
            "_v72_build_detail_bundle = _v72_build_detail_bundle"
        ))
        $Replacement = "# Force late references to use cache-enabled builder"
    }
    "remove-v72-self-assignment-u90" {
        $TargetBlock = [string]::Join($NewLine, @(
            "# Force late references",
            "# Legacy self-assignment candidate: this assignment is a no-op.",
            "# Keep the cache-enabled function definition above until regression-tested.",
            "_v72_build_detail_bundle = _v72_build_detail_bundle"
        ))
        $Replacement = "# Force late references"
    }
}

# Dùng anchor cụ thể để chỉ xóa block mục tiêu, không chạm function hoặc override U108.
$MatchCount = Count-LiteralOccurrences -Text $Text -Needle $TargetBlock

if ($MatchCount -eq 0) {
    Fail "Không tìm thấy block cleanup chính xác cho Step '$Step'."
}
if ($MatchCount -gt 1) {
    Fail "Tìm thấy $MatchCount block cho Step '$Step', nhiều hơn số kỳ vọng là 1."
}

$UpdatedText = $Text.Replace($TargetBlock, $Replacement)
[System.IO.File]::WriteAllText($BackendPath, $UpdatedText, $Utf8NoBom)

Write-Host "APPLIED: $Step"

# Nếu compile hoặc regression lỗi, khôi phục backend về trạng thái trước cleanup.
Invoke-CleanupTest "py_compile" {
    python -m py_compile "00.Detaisublog_v26.py"
}

Invoke-CleanupTest "Regression test" {
    powershell -ExecutionPolicy Bypass -File .\tools\run-regression.ps1
}

Write-Host "CLEANUP STEP PASSED"
