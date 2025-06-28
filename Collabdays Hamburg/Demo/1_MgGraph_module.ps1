function Install-GraphModules {
    # Define required modules
    $modules = @{
        'Microsoft Graph Authentication' = 'Microsoft.Graph.Authentication'
    }

    foreach ($module in $modules.GetEnumerator()) {
        # Check if the module is already installed
        if (Get-Module -Name $module.value -ListAvailable -ErrorAction SilentlyContinue) {
            Write-Host "Module $($module.Value) is already installed."
        }
        else {
            # Show the install confirmation popup
            $result = Show-InstallModulePopup -moduleName $module.Name

            # If the user clicks OK, proceed with the installation
            if ($result -eq $true) {
                try {
                    # Check if NuGet is installed
                    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
                        try {
                            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop | Out-Null
                            Write-Host "Installed PackageProvider NuGet"
                        }
                        catch {
                            Write-Warning "Error installing provider NuGet, exiting..."
                            return
                        }
                    }

                    # Set PSGallery as a trusted repository if not already
                    if ((Get-PSRepository -Name PSGallery).InstallationPolicy -ne 'Trusted') {
                        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
                    }

                    Write-Host ("Installing and importing PowerShell module {0}" -f $module.Value)
                    Install-Module -Name $module.Value -Force -ErrorAction Stop
                    Import-Module -Name $module.Value -ErrorAction Stop
                    Write-Host "Successfully installed and imported PowerShell module $($module.Value)"
                }
                catch {
                    Write-Warning ("Error installing or importing PowerShell module {0}, exiting..." -f $module.Value)
                    return
                }
            }
            else {
                # If the user cancels, log and close the script
                Write-Host "User canceled installation of module $($module.Value)."
                Exit
            }
        }
    }
}