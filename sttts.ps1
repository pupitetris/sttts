#Requires -PSEdition Core
#Requires -Version 7

[CmdletBinding(PositionalBinding=$false, DefaultParameterSetName="Speak")]
Param (
    [Parameter(ParameterSetName="Speak", HelpMessage="whisper-stream.exe location")]
    [Parameter(ParameterSetName="Clipboard", HelpMessage="whisper-stream.exe location")]
    [PSDefaultValue(Help="Vanilla whisper", Value="whisper\bin")]
    [Alias("b")]
    [string]$BinPath = "whisper\bin",

    [Parameter(ParameterSetName="Speak", HelpMessage="Choose voice by gender")]
    [PSDefaultValue(Help="System's default voice synthethizer", Value="")]
    [ValidateSet("male", "female", "")]
    [AllowEmptyString()]
    [Alias("g")]
    [string]$Gender = "",

    [Parameter(ParameterSetName="Speak", HelpMessage="Choose voice by name (use -ListVoices for values)")]
    [PSDefaultValue(Help="System's default voice synthethizer", Value="")]
    [ValidateScript({ $_ -eq "" -or $_ -in (& $PSCommandPath -ListVoices) }, ErrorMessage="Get valid names using -ListVoices")]
    [ArgumentCompleter({
	Param($cmd, $param, $wordToComplete)
	[array]$validValues = (& $PSCommandPath -ListVoices)
	$validValues -like "$wordToComplete*"
    })]
    [AllowEmptyString()]
    [Alias("n", "Name")]
    [string]$Voice = "",

    [Parameter(ParameterSetName="Speak", HelpMessage="Set voice speed rate (-10 to 10)")]
    [PSDefaultValue(Help="Normal speed", Value=0)]
    [ValidateRange(-10, 10)]
    [Alias("r")]
    [int]$Rate = 0,

    [Parameter(ParameterSetName="Speak", HelpMessage="Set voice volume (0 to 100)")]
    [PSDefaultValue(Help="Full volume", Value=100)]
    [ValidateRange(0, 100)]
    [Alias("v")]
    [int]$Volume = 100,

    [Parameter(ParameterSetName="Speak", HelpMessage="Let whisper know what language we are speaking into the mic")]
    [Parameter(ParameterSetName="Clipboard", HelpMessage="Let whisper know what language we are speaking into the mic")]
    [Parameter(ParameterSetName="ListVoices", HelpMessage="Filter the voice list with this language")]
    [PSDefaultValue(Help="Automatic language recognition", Value="auto")]
    [ValidateSet("auto", "en", "english", "zh", "chinese", "de", "german", "es", "spanish", "ru", "russian", "ko", "korean", "fr", "french", "ja", "japanese", "pt", "portuguese", "tr", "turkish", "pl", "polish", "ca", "catalan", "nl", "dutch", "ar", "arabic", "sv", "swedish", "it", "italian", "id", "indonesian", "hi", "hindi", "fi", "finnish", "vi", "vietnamese", "he", "hebrew", "uk", "ukrainian", "el", "greek", "ms", "malay", "cs", "czech", "ro", "romanian", "da", "danish", "hu", "hungarian", "ta", "tamil", "no", "norwegian", "th", "thai", "ur", "urdu", "hr", "croatian", "bg", "bulgarian", "lt", "lithuanian", "la", "latin", "mi", "maori", "ml", "malayalam", "cy", "welsh", "sk", "slovak", "te", "telugu", "fa", "persian", "lv", "latvian", "bn", "bengali", "sr", "serbian", "az", "azerbaijani", "sl", "slovenian", "kn", "kannada", "et", "estonian", "mk", "macedonian", "br", "breton", "eu", "basque", "is", "icelandic", "hy", "armenian", "ne", "nepali", "mn", "mongolian", "bs", "bosnian", "kk", "kazakh", "sq", "albanian", "sw", "swahili", "gl", "galician", "mr", "marathi", "pa", "punjabi", "si", "sinhala", "km", "khmer", "sn", "shona", "yo", "yoruba", "so", "somali", "af", "afrikaans", "oc", "occitan", "ka", "georgian", "be", "belarusian", "tg", "tajik", "sd", "sindhi", "gu", "gujarati", "am", "amharic", "yi", "yiddish", "lo", "lao", "uz", "uzbek", "fo", "faroese", "ht", "haitian creole", "ps", "pashto", "tk", "turkmen", "nn", "nynorsk", "mt", "maltese", "sa", "sanskrit", "lb", "luxembourgish", "my", "myanmar", "bo", "tibetan", "tl", "tagalog", "mg", "malagasy", "as", "assamese", "tt", "tatar", "haw", "hawaiian", "ln", "lingala", "ha", "hausa", "ba", "bashkir", "jw", "javanese", "su", "sundanese", "yue", "cantonese")]
    [Alias("l")]
    [string]$Lang = "auto",

    [Parameter(ParameterSetName="Speak", HelpMessage="Voice Activity Detection threshold")]
    [Parameter(ParameterSetName="Clipboard", HelpMessage="Voice Activity Detection threshold")]
    [ValidateRange(0.0, 1.0)]
    [float]$Vth = 0.75,

    [Parameter(ParameterSetName="Speak", HelpMessage="Whisper model, will be downloaded if not found")]
    [Parameter(ParameterSetName="Clipboard", HelpMessage="Whisper model, will be downloaded if not found")]
    [PSDefaultValue(Help="Base size quantized to 5bit with bias", Value="base-q5_1")]
    [ValidateScript({ $_ -eq "" -or $_ -in (& $PSCommandPath -ListModels) }, ErrorMessage="Get available models using -ListModels")]
    [ArgumentCompleter({
	Param($cmd, $param, $wordToComplete)
	[array]$validValues = (& $PSCommandPath -ListModels)
	$validValues -like "$wordToComplete*"
    })]
    [Alias("m")]
    [string]$Model = "base-q5_1",

    [Parameter(Mandatory, ParameterSetName="ListVoices", HelpMessage="List voice synthesizers available to the system and exit")]
    [switch]$ListVoices,

    [Parameter(Mandatory, ParameterSetName="ListModels", HelpMessage="List available whisper models and exit")]
    [switch]$ListModels,

    [Parameter(Mandatory, ParameterSetName="Clipboard", HelpMessage="Send captured text to the clipboard")]
    [Alias("c")]
    [switch]$Clipboard,
    
    [Parameter(Mandatory, ParameterSetName="Help", HelpMessage="Show detailed help for each parameter and exit")]
    [Alias("h")]
    [switch]$Help
)


Set-StrictMode -Version 3.0
$ErrorActionPreference = "stop"


$MODEL_DIR = "$PSScriptRoot\models"
$MODEL_DOWNLOAD_SCRIPT = "$MODEL_DIR\download-ggml-model.cmd"
$WHISPER_VERSION = "v1.8.2"

Set-Variable -Option ReadOnly -Name `
MODEL_DIR,
MODEL_DOWNLOAD_SCRIPT,
WHISPER_VERSION


filter Write-Debug-Or-Verbose () {
    if ($DebugPreference) {
	$_ | Write-Debug
    }
    else
    {
	$_ | Write-Verbose
    }
}


filter Tee-Object-If-Debug () {
    if ($DebugPreference) {
	$_ | Tee-Object -Append $PSScriptRoot\debug_log.txt
    }
    else
    {
	$_
    }
}


function Write-Report ([string]$field, [string]$value) {
    Write-Host -ForegroundColor DarkCyan -NoNewline "${field}: "
    Write-Host $value
}


function Write-Prompt () {
    Write-Host -ForegroundColor Cyan -NoNewline '> '
}


function Download ($splatHash) {
    Write-Report "Downloading" $splatHash.Uri
    Invoke-WebRequest @splatHash
}


function Install-Whisper ([string]$version, [string]$path) {
    if (Test-Path $path) {
	return
    }

    $base = Split-Path -Parent $path
    # We dont install in custom locations
    if ($base -ne "$PSScriptRoot\whisper") {
	return
    }

    if (-not (Test-Path $base)) {
	$Null = New-Item -ItemType Directory -Path $base
    }

    $flavour = Split-Path -Leaf $path
    $zipBase = "$base\whisper-$version-$flavour-x64"
    $zipFile = "$zipBase.zip"
    if (-not (Test-Path "$base\$zipFile")) {
	Download @{
	    Uri = "https://github.com/ggml-org/whisper.cpp/releases/download/$version/whisper-$flavour-x64.zip"
	    OutFile = $zipFile
	}
    }

    Expand-Archive $zipFile -DestinationPath $zipBase
    if (-not (Test-Path $path)) {
	$Null = New-Item -ItemType Directory -Path $path
    }
    Copy-Item -Recurse -Force -Path "$zipBase\Release\*.dll" -Destination $path
    Copy-Item -Recurse -Force -Path "$zipBase\Release\whisper-stream.exe" -Destination $path
    Remove-Item -Recurse -Force -Path $zipBase
    Remove-Item -Path $zipFile
}


function Get-ParameterInfo-Default-Description ($paramInfo) {
    $attr = $paramInfo.Attributes | Where-Object { $_.TypeId.Name -eq "PSDefaultValueAttribute" }
    $res = ""
    if ($attr) {
	$res = "  Default:"
	if ([string]$attr.Value -ne "") {
	    $res += " $($attr.Value)"
	}
	if ($attr.Help) {
	    $res += " ($($attr.Help))"
	}
    }
    $res
}


if ($Help) {
    Get-Help $PSCommandPath
    Write-Host Parameters:    
    (Get-Command $PSCommandPath).ParameterSets.Parameters |
    Where-Object HelpMessage | Sort-Object -Property Name -Unique |
    Format-Table -HideTableHeaders @{ Expression={"-$($_.Name)"} },
    @{ Expression={"-$($_.Aliases[0])"} },
    HelpMessage,
    @{ Expression={Get-ParameterInfo-Default-Description($_)} }
    Exit
}

if ($ListVoices) {
    Add-Type -AssemblyName System.Speech
    $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
    if ($Lang -ne "" -and $Lang -ne "auto") {
	(Get-Culture -ListAvailable |
	Where-Object Name -match ^$Lang |
	ForEach-Object { $synth.GetInstalledVoices($_) }
	).VoiceInfo.Name -replace ' ', '_'
    }
    else
    {
	$synth.GetInstalledVoices().VoiceInfo.Name -replace ' ', '_'
    }
    Exit
}
    
if ($ListModels) {
    & $MODEL_DOWNLOAD_SCRIPT -l
    Exit
}


if (Split-Path -Path $BinPath -IsAbsolute) {
    $whisperPath = $BinPath
}
else
{
    $whisperPath = "$PSScriptRoot\$BinPath"
    if (-not (Test-Path $whisperPath)) {
	Install-Whisper $WHISPER_VERSION $whisperPath
    }
}
$whisper = "$whisperPath\whisper-stream.exe"


$voiceName = $Voice -replace '_', ' '
if ($voiceName -ne "") {
    $synth.SelectVoice($voiceName)
}


if ($Gender -ne "") {
    if ($Lang -ne "" -and $lang -ne "auto") {
	$synth.SelectVoiceByHints($Gender, "adult", 0, $Lang)
    }
    else
    {
	$synth.SelectVoiceByHints($Gender)
    }
}


if (-not $Clipboard) {
    Add-Type -AssemblyName System.Speech
    $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
}


if (Test-Path variable:synth) {
    $synth.Rate = $Rate
    $synth.Volume = $Volume

    Write-Report "Voice" $synth.Voice.Name
    Write-Report "Gender" $synth.Voice.Gender
    Write-Report "Rate" $synth.Rate
    Write-Report "Volume" $synth.Volume
    Format-List -InputObject $synth.Voice | Out-String -Stream | Write-Debug-Or-Verbose
}


$modelFile = "$MODEL_DIR\ggml-$Model.bin"

if (-not (Test-Path $modelFile)) {
    & $MODEL_DOWNLOAD_SCRIPT $Model $MODEL_DIR
}


$lastText = ''
$currLang = $Lang
[console]::InputEncoding = [console]::OutputEncoding = New-Object System.Text.UTF8Encoding
$args = '-ng', '-m', $modelFile, '--step', '0', '--length', '5000', '-vth', $Vth, '-l', $Lang
& $whisper @args 2>&1 | Tee-Object-If-Debug | ForEach-Object {
    Write-Debug $_
    if ($_ -eq '[Start speaking]') {
	Write-Prompt	
    }
    if ($_ -match '^whisper_full_with_state: auto-detected language: ([a-z]+)') {
	$newLang = $Matches[1]
	if ($newLang -ne $currLang) {
	    $currLang = $newLang
	    if (Test-Path variable:synth) {
		$synth.SelectVoiceByHints($synth.Voice.Gender, $synth.Voice.Age, 0, $newLang)
		Format-List -InputObject $synth.Voice | Out-String -Stream | Write-Debug-Or-Verbose
	    }
	}
    }
    if ($_ -match '^\[0') {
	if (-not $DebugPreference) {
	    Write-Verbose $_
	}
	$text = $_ -replace '^[^\]]*\] *', ''
	$text = $text -replace '[\[\(\*](laughs|risas)[\)\]\*]', 'Ja ja ja'
	$text = $text -replace '\[[^\]]*\] *', ''
	$text = $text -replace '\([^\)]\) *', ''
	$text = $text -replace '\*[^\*]\* *', ''
	$text = $text -replace '\*+', ''
	if ([string]::IsNullOrWhiteSpace($text)) {
	    return
	}
	if ($lastText -ne $text) {
	    Write-Output $text	
	    if (Test-Path variable:synth) {
		$synth.Speak($text)
	    }
	    if ($Clipboard) {
		Set-Clipboard $text
	    }
	    Write-Prompt	
	}
	$lastText = $text
    }
}
