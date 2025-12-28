$file = "Microsoft Ultimate Installer.ps1"
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$null, [ref]$errors)
if ($errors) {
    foreach ($e in $errors) {
        Write-Warning "Error at line $($e.Extent.StartLineNumber), column $($e.Extent.StartColumnNumber): $($e.Message)"
    }
}
else {
    Write-Output "No syntax errors found."
}