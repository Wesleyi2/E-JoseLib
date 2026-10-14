python app.py 2>&1 | ForEach-Object {
    $line = $_.ToString()

    if ($line -match '"[^"]*" 2\d\d') {
        Write-Host $line -ForegroundColor Green
    }
    elseif ($line -match '"[^"]*" 3\d\d') {
        Write-Host $line -ForegroundColor Cyan
    }
    elseif ($line -match '"[^"]*" 4\d\d') {
        Write-Host $line -ForegroundColor Yellow
    }
    elseif ($line -match '"[^"]*" 5\d\d') {
        Write-Host $line -ForegroundColor Red
    }
    else {
        Write-Host $line
    }
}