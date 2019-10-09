# From this answer: https://superuser.com/questions/1185779/is-it-possible-to-install-fonts-in-windows-without-admin-rights#1306464

Add-Type -Name Session -Namespace "" -Member @"
[DllImport("gdi32.dll")]
public static extern int AddFontResource(string filePath);
"@

$Source = "C:\YourCustomFontRespository"
$count = 0
$successes = 0
ForEach($font in Get-ChildItem $Source -Recurse -Include *.ttf, *.otf) {
    $result = [Session]::AddFontResource($font.FullName)
    If ($result = 1) {
        Write-Output "Success: $($font.Name)"
        $successes += 1
    }
    else {
        Write-Output "Failed: $($font.Name)"
    }
    $count += 1
}
Write-Output ""
Write-Output "$successes / $count fonts added."
