$url = "https://unsplash.com/"

$response = Invoke-WebRequest -Uri $url
$images = $response.Images

if ($images.Count -lt 1) {
    throw "No images found in the Unsplash response."
}

$firstImage = $images[0]
$imageFullUrl = $firstImage.Src

$uri = New-Object System.Uri($imageFullUrl)
$imageUrl = $uri.GetLeftPart([System.UriPartial]::Path)
$fileName = "unsplash4lockscreen-" + [System.IO.Path]::GetFileName($uri.LocalPath) + ".jpg"

$downloadsFolderPath = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$imagePath = Join-Path -Path $downloadsFolderPath -ChildPath $fileName

Invoke-WebRequest -Uri $imageUrl -OutFile $imagePath

# Reference: https://github.com/nccgroup/Change-Lockscreen/blob/master/Change-Lockscreen.ps1
[Windows.System.UserProfile.LockScreen, Windows.System.UserProfile, ContentType = WindowsRuntime] | Out-Null
Add-Type -AssemblyName System.Runtime.WindowsRuntime

Function Await($WinRtTask, $ResultType) {
    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}

Function AwaitAction($WinRtAction) {
    $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and !$_.IsGenericMethod })[0]
    $netTask = $asTask.Invoke($null, @($WinRtAction))
    $netTask.Wait(-1) | Out-Null
}


[Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
$image = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($imagePath)) ([Windows.Storage.StorageFile])
AwaitAction ([Windows.System.UserProfile.LockScreen]::SetImageFileAsync($image))

Remove-Item $imagePath

Write-Host "Done."
