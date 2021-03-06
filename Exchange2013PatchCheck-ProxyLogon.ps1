# Code from Reference: https://techcommunity.microsoft.com/t5/core-infrastructure-and-security/how-to-correctly-check-file-versions-with-powershell/ba-p/257642
Update-TypeData -TypeName System.Io.FileInfo -MemberType ScriptProperty -MemberName FileVersionUpdated -Value {
    New-Object System.Version -ArgumentList @(
        $this.VersionInfo.FileMajorPart
        $this.VersionInfo.FileMinorPart
        $this.VersionInfo.FileBuildPart
        $this.VersionInfo.FilePrivatePart
    )
}


$exchangePath = (Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath
$fileVersions = Import-Csv -Path "$PSScriptRoot\ExchangeFilePatchVersions.csv"

foreach ($file in $fileVersions) {
    $filePath = Join-Path -Path $exchangePath -ChildPath "bin" | Join-Path -ChildPath $file.Filename
    if (Test-Path $filePath) {
        if ( (Get-Item -Path $filePath).FileVersionUpdated.toString() -eq $($file.FileVersion)) {
            Write-Host $("{0} is patched" -f $file.Filename) -ForegroundColor Green
        }
        else {
            Write-Host $("{0} is not patched" -f $file.Filename) -ForegroundColor Red
            (Get-Item -Path $filePath).VersionInfo.FileVersion
        }
    }
    else {
        Write-Host $("{0} file not found." -f $file.Filename) -ForegroundColor Yellow
    }
}