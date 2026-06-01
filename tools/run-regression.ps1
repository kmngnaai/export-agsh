$ErrorActionPreference = "Stop"

# Luôn chạy từ thư mục gốc repo để các đường dẫn tests/... hoạt động ổn định.
$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

# Output test mặc định. Chỉ dùng dữ liệu copy, không dùng dữ liệu thật.
$OutputPath = "D:\01.AutobyNgan\00.Build.App\11.CODEX\02_LOCAL_TEST\export-agsh-test-data\202605test\test_output_fix_vas_hc.xlsx"

function Invoke-RegressionStep {
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
            Write-Host "FAIL: $Name (exit code: $LASTEXITCODE)"
            exit $LASTEXITCODE
        }
    }
    catch {
        Write-Host "FAIL: $Name"
        Write-Host $_
        exit 1
    }

    Write-Host "PASS: $Name"
}

# Chạy lần lượt và dừng ngay nếu một bước thất bại.
Invoke-RegressionStep "smoke import backend" {
    python tests/smoke_import_backend.py
}

Invoke-RegressionStep "output baseline" {
    python tests/check_output_baseline.py $OutputPath
}

Invoke-RegressionStep "Vas/HC multi Cus no" {
    python tests/check_vas_hc_multi_cus.py $OutputPath
}

Invoke-RegressionStep "deleted path absent" {
    python tests/check_deleted_path_absent.py $OutputPath "SUB-INV33333.xls"
}

Invoke-RegressionStep "new path present" {
    python tests/check_path_present.py $OutputPath "SUB-INV3333test.xls"
}

Invoke-RegressionStep "path occurrences" {
    python tests/check_path_occurrences.py $OutputPath "SUB-INV3333test.xls"
}

Invoke-RegressionStep "git status" {
    git status
}

Write-Host ""
Write-Host "ALL REGRESSION TESTS PASSED"
