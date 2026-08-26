$ErrorActionPreference = 'Stop'

Write-Host ''
Write-Host 'GMP Offline OCR - Hardware Check' -ForegroundColor Cyan
Write-Host '================================' -ForegroundColor Cyan

$gpus = Get-CimInstance Win32_VideoController

Write-Host ''
Write-Host 'Detected display adapters:' -ForegroundColor Yellow
foreach ($gpu in $gpus) {
    $vramGb = if ($gpu.AdapterRAM) { [math]::Round(([uint64]$gpu.AdapterRAM / 1GB), 1) } else { 'unknown' }
    Write-Host "- $($gpu.Name) | VRAM: $vramGb GB | Driver: $($gpu.DriverVersion)"
}

$nvidiaAdapters = @($gpus | Where-Object { $_.Name -match 'NVIDIA' })
if ($nvidiaAdapters.Count -eq 0) {
    Write-Host ''
    Write-Host 'Result: no NVIDIA GPU detected.' -ForegroundColor Red
    Write-Host 'The current DeepSeek-OCR server requires an NVIDIA CUDA GPU. Intel, AMD, and integrated GPUs are not supported by this configuration.' -ForegroundColor Yellow
    exit 1
}

$nvidiaSmi = Get-Command nvidia-smi -ErrorAction SilentlyContinue
if (-not $nvidiaSmi) {
    Write-Host ''
    Write-Host 'Result: NVIDIA GPU detected, but nvidia-smi is unavailable.' -ForegroundColor Yellow
    Write-Host 'Install or update the NVIDIA driver, then run this check again.' -ForegroundColor Yellow
    exit 1
}

Write-Host ''
Write-Host 'NVIDIA CUDA devices:' -ForegroundColor Yellow
$deviceInfo = & nvidia-smi --query-gpu=name,memory.total,driver_version --format=csv,noheader 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host $deviceInfo -ForegroundColor Red
    exit 1
}
$deviceInfo | ForEach-Object { Write-Host "- $_" }

$memoryValues = & nvidia-smi --query-gpu=memory.total --format=csv,noheader,nounits
$maxMemoryMb = ($memoryValues | ForEach-Object { [int]$_.Trim() } | Measure-Object -Maximum).Maximum

Write-Host ''
if ($maxMemoryMb -ge 16000) {
    Write-Host "Result: compatible for local DeepSeek-OCR testing ($maxMemoryMb MB VRAM detected)." -ForegroundColor Green
    Write-Host 'Next step: install NVIDIA Container Toolkit, then deploy the local OCR server.'
    exit 0
}

if ($maxMemoryMb -ge 12000) {
    Write-Host "Result: NVIDIA GPU detected with $maxMemoryMb MB VRAM." -ForegroundColor Yellow
    Write-Host 'It may work with reduced OCR resolution, but 16 GB VRAM is the tested target.'
    exit 0
}

Write-Host "Result: $maxMemoryMb MB VRAM detected, which is below the 16 GB tested target." -ForegroundColor Red
Write-Host 'Do not deploy the current model configuration until a lower-memory configuration is validated.'
exit 1
