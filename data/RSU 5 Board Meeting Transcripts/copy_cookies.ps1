$src = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Network\Cookies"
$dst = "$env:TEMP\edge_cookies_copy.db"
$fs = [System.IO.File]::Open($src, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
$buf = New-Object byte[] $fs.Length
$null = $fs.Read($buf, 0, $buf.Length)
$fs.Close()
[System.IO.File]::WriteAllBytes($dst, $buf)
Write-Host "Copied $($buf.Length) bytes to $dst"
