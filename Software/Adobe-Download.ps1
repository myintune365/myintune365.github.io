$ManualDownload = $false
if ($ManualDownload) {
	$Version = "2300320201"
	"https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/$($Version)/AcroRdrDCx64$($Version)_en_US.exe" | Set-Clipboard
}

$DownloadsFolder = "C:\Windows\Temp\KB"
Remove-Item $DownloadsFolder -Recurse -ErrorAction Ignore 
mkdir $DownloadsFolder

# Download the latest Adobe Acrobat Reader DC x64
# https://armmf.adobe.com/arm-manifests/mac/AcrobatDC/reader/current_version.txt

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Get the link to the latest Adobe Acrobat Reader DC x64 installer
$TwoLetterISOLanguageName = (Get-WinSystemLocale).TwoLetterISOLanguageName
$Parameters = @{
	Uri             = "https://rdc.adobe.io/reader/products?lang=$($TwoLetterISOLanguageName)&site=enterprise&os=Windows%2011&api_key=dc-get-adobereader-cdn"
	UseBasicParsing = $true
}
$displayName = (Invoke-RestMethod @Parameters).products.reader.displayName
$Version = (Invoke-RestMethod @Parameters).products.reader.version.Replace(".", "")


$Parameters = @{
	Uri             = "https://rdc.adobe.io/reader/downloadUrl?name=$($displayName)&os=Windows%2011&site=enterprise&lang=$($TwoLetterISOLanguageName)&api_key=dc-get-adobereader-cdn"
	UseBasicParsing = $true
}
$downloadURL = (Invoke-RestMethod @Parameters).downloadURL
$saveName = (Invoke-RestMethod @Parameters).saveName

# if URl contains "reader", we need to fix the URl to download the latest version. Applicable for the Russian version
if ($downloadURL -match "reader") {
	$Parameters = @{
		Uri             = "https://rdc.adobe.io/reader/products?lang=en&site=enterprise&os=Windows 11&api_key=dc-get-adobereader-cdn"
		UseBasicParsing = $true
	}
	$Version = (Invoke-RestMethod @Parameters).products.reader.version.Replace(".", "")

	$IetfLanguageTag = (Get-WinSystemLocale).IetfLanguageTag.Replace("-", "_")
	$downloadURL = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/$($Version)/AcroRdrDCx64$($Version)_en_US.exe"
	$saveName = Split-Path -Path $downloadURL -Leaf
}

# Download the installer
Remove-Item $DownloadsFolder -Recurse -ErrorAction Ignore 
mkdir $DownloadsFolder 
$Parameters = @{
	Uri             = $downloadURL
	OutFile         = "$DownloadsFolder\$saveName"
	UseBasicParsing = $true
}


$ProgressPreference = "SilentlyContinue"
Invoke-RestMethod @Parameters

Start-Process "$DownloadsFolder"

