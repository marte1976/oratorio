param(
    [Parameter(Mandatory = $true)]
    [string]$CatalogPath,
    [Parameter(Mandatory = $true)]
    [string]$CourseKey,
    [Parameter(Mandatory = $true)]
    [string]$BaseDir,
    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Runtime.WindowsRuntime
[void][Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
[void][Windows.Storage.StorageFolder, Windows.Storage, ContentType=WindowsRuntime]
[void][Windows.Storage.FileProperties.MusicProperties, Windows.Storage, ContentType=WindowsRuntime]
[void][Windows.Media.Editing.MediaComposition, Windows.Media.Editing, ContentType=WindowsRuntime]
[void][Windows.Media.Editing.MediaClip, Windows.Media.Editing, ContentType=WindowsRuntime]
[void][Windows.Media.Editing.BackgroundAudioTrack, Windows.Media.Editing, ContentType=WindowsRuntime]
[void][Windows.Media.MediaProperties.MediaEncodingProfile, Windows.Media.MediaProperties, ContentType=WindowsRuntime]
[void][Windows.Media.MediaProperties.VideoEncodingQuality, Windows.Media.MediaProperties, ContentType=WindowsRuntime]
[void][Windows.Media.Editing.MediaTrimmingPreference, Windows.Media.Editing, ContentType=WindowsRuntime]
[void][Windows.Media.Transcoding.TranscodeFailureReason, Windows.Media.Transcoding, ContentType=WindowsRuntime]

$script:AsTaskActionMethod = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object { $_.Name -eq "AsTask" -and -not $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 } |
    Select-Object -First 1
$script:AsTaskGenericMethod = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object { $_.Name -eq "AsTask" -and $_.IsGenericMethod -and $_.GetGenericArguments().Count -eq 1 -and $_.GetParameters().Count -eq 1 } |
    Select-Object -First 1
$script:AsTaskProgressMethod = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object { $_.Name -eq "AsTask" -and $_.IsGenericMethod -and $_.GetGenericArguments().Count -eq 2 -and $_.GetParameters().Count -eq 1 } |
    Select-Object -First 1

function Invoke-WinRtAction {
    param([Parameter(Mandatory = $true)]$AsyncAction)
    $task = $script:AsTaskActionMethod.Invoke($null, @($AsyncAction))
    $task.GetAwaiter().GetResult() | Out-Null
}

function Invoke-WinRtOperation {
    param(
        [Parameter(Mandatory = $true)]$AsyncOperation,
        [Parameter(Mandatory = $true)][Type]$ResultType
    )
    $method = $script:AsTaskGenericMethod.MakeGenericMethod($ResultType)
    $task = $method.Invoke($null, @($AsyncOperation))
    return $task.GetAwaiter().GetResult()
}

function Invoke-WinRtOperationWithProgress {
    param(
        [Parameter(Mandatory = $true)]$AsyncOperation,
        [Parameter(Mandatory = $true)][Type]$ResultType,
        [Parameter(Mandatory = $true)][Type]$ProgressType
    )
    $method = $script:AsTaskProgressMethod.MakeGenericMethod($ResultType, $ProgressType)
    $task = $method.Invoke($null, @($AsyncOperation))
    return $task.GetAwaiter().GetResult()
}

function Resolve-VoiceToken {
    param($Voice)
    $tokens = $Voice.GetVoices()
    if (-not $tokens -or $tokens.Count -le 0) {
        return $null
    }
    for ($i = 0; $i -lt $tokens.Count; $i++) {
        $token = $tokens.Item($i)
        $description = [string]$token.GetDescription()
        if ($description -match "Italian" -or $description -match "Elsa") {
            return $token
        }
    }
    return $tokens.Item(0)
}

$catalog = Get-Content -LiteralPath $CatalogPath -Raw | ConvertFrom-Json
$course = $catalog.$CourseKey
if (-not $course) {
    throw "Corso tutorial non trovato: $CourseKey"
}
$slides = @($course.slides)
if (-not $slides.Count) {
    throw "Il corso tutorial non contiene slide."
}

$outputDir = Split-Path -Path $OutputPath -Parent
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
$workDir = Join-Path ([System.IO.Path]::GetTempPath()) ("tutorial-video-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $workDir | Out-Null

$voice = New-Object -ComObject SAPI.SpVoice
$fileStream = New-Object -ComObject SAPI.SpFileStream

try {
    $voiceToken = Resolve-VoiceToken -Voice $voice
    if ($voiceToken) {
        $voice.Voice = $voiceToken
    }
    $voice.Rate = -1
    $voice.Volume = 100

    $composition = New-Object Windows.Media.Editing.MediaComposition
    $clipsProp = [Windows.Media.Editing.MediaComposition].GetProperty("Clips")
    $tracksProp = [Windows.Media.Editing.MediaComposition].GetProperty("BackgroundAudioTracks")
    $clipsObj = $clipsProp.GetValue($composition)
    $tracksObj = $tracksProp.GetValue($composition)
    $clipsInsertMethod = $clipsProp.PropertyType.GetMethod("Insert")
    $tracksInsertMethod = $tracksProp.PropertyType.GetMethod("Insert")
    $offset = [TimeSpan]::Zero

    for ($index = 0; $index -lt $slides.Count; $index++) {
        $slide = $slides[$index]
        $audioText = [string]$slide.audio_text
        if ([string]::IsNullOrWhiteSpace($audioText)) {
            $audioText = [string]$slide.title
        }
        $relativeImagePath = Join-Path ("static\guide-media\images\" + $CourseKey) ("{0:D2}.png" -f ($index + 1))
        $imagePath = Join-Path $BaseDir $relativeImagePath
        if (-not (Test-Path -LiteralPath $imagePath)) {
            throw "Immagine slide non trovata: $imagePath"
        }

        $wavPath = Join-Path $workDir ("slide-{0:D2}.wav" -f ($index + 1))
        $fileStream.Open($wavPath, 3, $false)
        $voice.AudioOutputStream = $fileStream
        [void]$voice.Speak($audioText)
        $fileStream.Close()

        $audioFile = Invoke-WinRtOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($wavPath)) ([Windows.Storage.StorageFile])
        $audioProps = Invoke-WinRtOperation ($audioFile.Properties.GetMusicPropertiesAsync()) ([Windows.Storage.FileProperties.MusicProperties])
        $audioDuration = $audioProps.Duration
        if ($audioDuration.TotalMilliseconds -lt 800) {
            $audioDuration = [TimeSpan]::FromSeconds(2)
        }
        $clipDuration = $audioDuration.Add([TimeSpan]::FromMilliseconds(500))

        $imageFile = Invoke-WinRtOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($imagePath)) ([Windows.Storage.StorageFile])
        $clip = Invoke-WinRtOperation ([Windows.Media.Editing.MediaClip]::CreateFromImageFileAsync($imageFile, $clipDuration)) ([Windows.Media.Editing.MediaClip])
        [void]$clipsInsertMethod.Invoke($clipsObj, @([int]$index, $clip))

        $track = Invoke-WinRtOperation ([Windows.Media.Editing.BackgroundAudioTrack]::CreateFromFileAsync($audioFile)) ([Windows.Media.Editing.BackgroundAudioTrack])
        $track.Delay = $offset
        $track.Volume = 1.0
        [void]$tracksInsertMethod.Invoke($tracksObj, @([int]$index, $track))

        $offset = $offset.Add($clipDuration)
    }

    if (-not (Test-Path -LiteralPath $OutputPath)) {
        New-Item -ItemType File -Force -Path $OutputPath | Out-Null
    }
    $outputFile = Invoke-WinRtOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($OutputPath)) ([Windows.Storage.StorageFile])

    $profile = [Windows.Media.MediaProperties.MediaEncodingProfile]::CreateMp4(
        [Windows.Media.MediaProperties.VideoEncodingQuality]::HD1080p
    )
    $result = Invoke-WinRtOperationWithProgress (
        $composition.RenderToFileAsync(
            $outputFile,
            [Windows.Media.Editing.MediaTrimmingPreference]::Precise,
            $profile
        )
    ) ([Windows.Media.Transcoding.TranscodeFailureReason]) ([double])

    if ($result -ne [Windows.Media.Transcoding.TranscodeFailureReason]::None) {
        throw "Render tutorial non riuscito: $result"
    }

    Write-Output $OutputPath
}
finally {
    try { $fileStream.Close() } catch {}
    try { Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
}
