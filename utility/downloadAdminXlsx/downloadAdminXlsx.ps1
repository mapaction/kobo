<#  
.SYNOPSIS  
    Downloads all the Admin Boundaries Excel files from HDX   

.DESCRIPTION  
    Queries the HDX website and downloads any Excel files tagged with Admin Boundaries  

.NOTES  
    File Name   : downloadAdminXlsx.ps1  
    Author      : Steve Hurst - shurst@mapaction.org
      
.PARAMETERS
    $folder     : Root folder to download files to.

.EXAMPLE
     .\downloadAdminXlsx.ps1 -folder C:\temp\adm
#>

param(
        [Parameter(Mandatory)]
        [string]$folder
     )
# Test Folder exists
if (Test-Path -Path $folder) {
    $WShell = New-Object -com “Wscript.Shell”

    # Get all the package names
    $url = "https://data.humdata.org/api/3/action/package_list"
    $resultJson = Invoke-WebRequest -Uri $url -Method POST
    $myJson = $resultJson | ConvertFrom-Json 

    # For every package
    foreach($result in $myJson.result)
    {
        # Stop screen from timing out
        $WShell.sendkeys(“{SCROLLLOCK}”)

        # Get all the details for this package
        $url = "https://data.humdata.org/api/3/action/package_show?id=" + $result
        $packageJson = Invoke-WebRequest -Uri $url -Method POST

        $valid = $true

        try {
            $packageJsonC = ConvertFrom-Json $packageJson -ErrorAction SilentlyContinue;
        } catch 
        {
            $valid = $false
        }

        if ($valid -eq $true)
        {
            # Loop around tags
            $adm = $false
            foreach($tag in $packageJsonC.result.tags)
            {
                # We will want to download this package if it has anything to do with "Admin"
                if ($tag.name.ToUpper().Contains( "ADM"))
                {
                    $adm = $true
                }
            }
            if ($adm -eq $true)
            {
                # For each resource in this Admin package
                foreach($resource in $packageJsonC.result.resources)
                {
                    # If it's active and a XLSX format, download it
                    if ($resource.state -eq "active" -and $resource.format -eq "XLSX")
                    {
                        $resourceFolder = $folder + "\" + $resource.id
                        $new = New-Item -Path $resourceFolder -ItemType "directory"
                        $resourceFilePath = $resourceFolder + "\" + $resource.name
                        Write-Host ("Downloading " + $resource.name + " to " + $resourceFolder)
                        Invoke-WebRequest -URI $resource.url -OutFile $resourceFilePath
                    }
                }
            }
        }
    }
} else {
    "Path doesn't exist."
}
