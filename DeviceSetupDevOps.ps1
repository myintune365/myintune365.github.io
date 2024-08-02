function Get-TenantIdFromInputBox($Label1Value) {

    try {
        $global:TenantObject = $null

        Add-Type -AssemblyName PresentationFramework

        # Calculate screen size and window size
        $screenWidth = [System.Windows.SystemParameters]::PrimaryScreenWidth
        $screenHeight = [System.Windows.SystemParameters]::PrimaryScreenHeight
        $windowWidth = [math]::Round($screenWidth / 2)
        $windowHeight = [math]::Round($screenHeight / 2)

        $XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Input Box" Height="$windowHeight" Width="$windowWidth" WindowStartupLocation="CenterScreen" Background="#FF2D2D30">
    <Grid>
        <Label Content="$($Label1Value):" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,10,0,0" FontSize="24" Foreground="White"/>
        <TextBox Name="TextBoxInput" HorizontalAlignment="Left" Height="40" Margin="10,50,0,0" VerticalAlignment="Top" Width="$($windowWidth - 40)" FontSize="24"/>
        <Button Content="OK" HorizontalAlignment="Left" Margin="10,100,0,0" VerticalAlignment="Top" Width="100" Height="40" Name="OkButton" FontSize="24" Background="#007ACC" Foreground="White"/>
        <Label Name="ResultLabel" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,150,0,0" FontSize="24" Foreground="White" Visibility="Hidden"/>
    </Grid>
</Window>
"@

        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($XAML))
        $window = [System.Windows.Markup.XamlReader]::Load($reader)

        $okButton = $window.FindName('OkButton')
        $textBox = $window.FindName('TextBoxInput')
        $resultLabel = $window.FindName('ResultLabel')

        function Update-InputBox {
            $null = [System.Windows.Forms.Application]::DoEvents()
        }

        $global:Count = 0
        $okButton.Add_Click({
                $input = $textBox.Text
                if ($input) {
                    $global:Count++

                    $DomainOrTenantId = $input -split "@" | Select-Object -Last 1

                    try {
                        $global:TenantObject = (Invoke-RestMethod -Uri ("https://accounts.accesscontrol.windows.net/" + $DomainOrTenantId + "/metadata/json/1"))
                        $TenantIdForAppAccess = $global:TenantObject.realm

                        if ($TenantIdForAppAccess.Length -eq 36) {
                            $resultLabel.Content = "[$global:Count] You entered: $input - which is associated with Tenant Id: [$TenantIdForAppAccess]"
                            $resultLabel.Visibility = [System.Windows.Visibility]::Visible

                            Update-InputBox
                            # Start-Sleep -Seconds 3

                            $null = $window.Close()  # Close the window
                        }
                        else {
                            $resultLabel.Content = "[$global:Count] You entered: $input - but we couldn't find the associated Microsoft 365 tenant"
                            $resultLabel.Visibility = [System.Windows.Visibility]::Visible
                            Update-InputBox
                        }
                    }
                    catch {
                        $resultLabel.Content = "[$global:Count] Error retrieving Tenant Id"
                        $resultLabel.Visibility = [System.Windows.Visibility]::Visible
                        Update-InputBox
                    }
                }
                else {
                    $resultLabel.Content = "Please enter a value."
                    $resultLabel.Visibility = [System.Windows.Visibility]::Visible
                    Update-InputBox
                }
            })

        $null = $window.ShowDialog()
        return $global:TenantObject

    }
    catch {

        Write-Host -f Red "Issue showing XAML box."
        Write-Host -f Magenta "$_.Exception"

        Start-Sleep -Seconds 5

        return $_.Exception
    }
}

function Get-AzureAccessControlFromTenantId($TenantId) {

    $AzureAccessControl = try { Invoke-RestMethod -Uri ("https://accounts.accesscontrol.windows.net/" + $TenantId + "/metadata/json/1") } catch { $null }
    $return_Domains = $AzureAccessControl.allowedAudiences -notmatch "^(.*@)(\w{8}-\w{4}-\w{4}-\w{4}-\w{12})$" | ForEach-Object { $_.Substring( $_.IndexOf('@') + 1 ) }
    $TenantOnMsft = $return_Domains | Where-Object { $_ -match ".onmicrosoft.com" -and $_ -notmatch "mail.onmicrosoft.com" } | Select-Object -First 1

    return [pscustomobject][ordered]@{ 
        AccessControlTenantId = $AzureAccessControl.realm
        TenantId              = "$TenantId"; 
        TenantOnMsft          = "$TenantOnMsft"
        TenantName            = $($TenantOnMsft -split "\.onmicrosoft.com" | Select-Object -First 1)
    }
}

function ModuleIsNotInstalled($Path) {
    $IsInstalled = try { Get-ChildItem $Path -EA SilentlyContinue } catch { $null }
    return $null -eq $IsInstalled
}

function ModuleIsInstalled($Path) {
    $IsInstalled = try { Get-ChildItem $Path -EA SilentlyContinue } catch { $null }
    return [bool]$IsInstalled
}


#######################################################################
###################### SCRIPT START ###################################
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force


[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$PnpPath = "C:\Program Files\WindowsPowerShell\Modules\PnP.PowerShell\1.12.0\PnP.PowerShell.psd1"
$NugetPath = "C:\Program Files\PackageManagement\ProviderAssemblies\nuget\*\Microsoft.PackageManagement.NuGetProvider.dll"

if (ModuleIsNotInstalled -Path $NugetPath) {
    Install-PackageProvider -Name nuget -MinimumVersion 2.8.5.201 -Force
}

if (ModuleIsNotInstalled -Path $PnpPath) {
    Install-Module -Name "PnP.PowerShell" -RequiredVersion 1.12.0 -Force -AllowClobber
}

$ModulesAreInstalled = (ModuleIsInstalled -Path $NugetPath) -and (ModuleIsInstalled -Path $PnpPath)

if (-not $ModulesAreInstalled) {
    
    Write-Host -f Red "Unable to continue, Microsoft PowerShell modules are not available yet."
    return
}


if ($TenantObject.realm.Length -ne 36) {
    $TenantObject = Get-TenantIdFromInputBox -Label1Value "Enter your Microsoft 365 email address"
}
if ($TenantObject) {

    $AAC = Get-AzureAccessControlFromTenantId -TenantId $TenantObject.realm
    $SpoTenantAdminUrl = "$($AAC.TenantName)-admin.sharepoint.com"
    $SpoTenantUrl = "$($AAC.TenantName).sharepoint.com"

    Connect-PnPOnline -Url $SpoTenantAdminUrl -Interactive
}