# Install-Module -Name WinSCP
# Import-Module WinSCP

# .---.           .                  .    .---.
# |               |                  |       / 
# |--- .-. .-. .-.|.  .    ._.-. .--.|.-.   /  
# |   (.-'(.-'(   | \  \  / (   )|   |-.'  /   
# '    `--'`--'`-'`- `' `'   `-' '   '  `-'---'

## TABLE OF CONTENTS
# 1 Misc Functions
# 2 WinSCP Functions
# 3 Azure Storage

#   .  
# .'|  
#   |  
#   |  
# '---'
## Misc Functions

# Funcion is used if the file needs to append the file date to the filename.
function generate_date {
    param (
        [string]$days_to_subtract
    )
    $desired_date = (Get-Date).AddDays($days_to_subtract)
    return $desired_date
}

# Function is used to Unzip files
function unzip_files {
    # If you pass only the zip path, it will extract all files into their own folder in the original directory
    # If you pass in a extract path, it will extract files into their own folder in that directory
    param (
        [string] $zip_path,
        [Parameter(Mandatory=$false)][string] $extract_to_path
    )
    $zip_files = Get-ChildItem -Path $zip_path -Filter *.zip

    foreach($zip_file in $zip_files) {
        if($extract_to_path){
            # $extract_path = Join-Path -Path $extract_to_path -ChildPath $zip_file.BaseName
            $extract_path = $extract_to_path
        }else{
            $extract_path = Join-Path -Path $zip_path -ChildPath $zip_file.BaseName
        }

        Expand-Archive -Path $zip_file.FullName -DestinationPath $extract_path

        Write-Output "Extracted $($zip_file.Name) to $extract_path"
    }
}

#  .-. 
# (   )
#   .' 
#  /   
# '---'
## WinSCP

# Creates WinSCP session to be used by functions.
function creat_winscp_session {
    param (
        [string] $hostname,
        [PSCredential] $credentials,
        [string] $ssh_hostkey
    )
    try{
        $session = New-WinSCPSession -SessionOption (New-WinSCPSessionOption -HostName $hostname -Protocol Sftp -Credential $credentials -SshHostKeyFingerprint $ssh_hostkey)
    }catch{
        "Unable to connect to FTP Site" 
    }
    

    return $session
}

# Checks the FTP site to see if files are available. 
function find_if_files_are_present {
    param (
        [string] $dir_path,
        [String] $filename,
        [WinSCP.Session] $session
    )
    try {
    
        # Navigate to remote folder
        $directoryInfo = $session.ListDirectory($dir_path)
    
        # Check for specific file
        $fileExists = $directoryInfo.Files | Where-Object { $_.Name.Trim() -eq $filename.Trim() }
    
        if ($fileExists) {
            Write-Host "File exists: $dir_path/$filename"
            return $true
        } else {
            Write-Host "File does not exist: $dir_path/$filename"
            return $false
        }
    
    } catch {
        Write-Host "Error: $($_.Exception.Message)"
   
    }   
}

function download_files {
    param (
        [WinSCP.Session] $session,
        [string] $dir_path,
        [string] $local_path,
        [Parameter(Mandatory=$false)][string] $filename
    )

    $transferOptions = New-Object WinSCP.TransferOptions
    $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary

    # Set ErrorActionPreference to Stop to catch all exceptions
    $ErrorActionPreference = "Stop"

    if ($filename){
        
        # Download specific file if passed as a parameter
        $remoteFilePath = Join-Path $dir_path $filename

        try{
            $transferResult = Receive-WinSCPItem -WinSCPSession $session -RemotePath $remoteFilePath -LocalPath $local_path
            
            if ($transferResult.IsSuccess){
                Write-Output "$filename downloaded successfully"
            }else{
                Write-Output "$filename did not download"
            }
        }
        catch{
            Write-Output "Error downloading file: $($_.Exception.Message)"
        }
    }else{
        # Download everything in directory (If no file parameter)
        $remoteFilesPattern = Join-Path $dir_path "*"
        
        try{
            $transferResult = Receive-WinSCPItem -WinSCPSession $session -RemotePath $remoteFilesPattern -LocalPath $local_path
      
            if ($transferResult.IsSuccess){
                Write-Output $transferResult.Transfers
            }else{
                $failures = $transferResult.Failures
                Write-Output $failures
            }
        }
        
        catch{
            Write-Output "Error downloading file: $($_.Exception.Message)"
        }
    }
}

function retry_missing_files {
    param (
        [WinSCP.Session] $session,
        [array] $missing_files
    )
    $max_retries = 10
    $retries = 0

    while($missing_files.Length -gt 0 -or $retries -lt $max_retries){
        foreach($full_path in $missing_files){
            $file = Split-Path $full_path -Leaf
            $dir = Split-Path $full_path -Parent
            $dir = $dir -replace "\\", "/"
            # Write-Output "Directory is: "  $dir
            # Write-output "File is: " $file
            $file_present = find_if_files_are_present -dir_path $dir -filename $file -session $win_scp_session
            if($file_present -eq $true){
                $missing_files = $missing_files | Where-Object{$_ -ne $full_path}
            }
        }
        Start-Sleep -Seconds 30
    }
    
}

function dispose_WinSCP_session {
    param (
        [WinSCP.Session] $session
    )
    $session.Dispose()
}

# .--. 
#     )
#  --: 
#     )
# `--' 
## Azure Storage

# Moves files from Temp directory to Azure Blob Storage. Needs a 'New-AzStorageContext' passed in as the $connection variable. 
function move_files_to_azure_blob {
    param (
        [String] $containerName,
        [String] $blobName,
        [String] $source,
        # [String] $file,
        [Microsoft.WindowsAzure.Commands.Storage.Common.AzureStorageContext] $connection
    )
    $file_path = $source + "\" + $blobName
    $container = Get-AzStorageContainer -Name $containerName -Context $connection
    if (-not $container) {
        New-AzStorageContainer -Name $containerName -Context $connection
    }

    $blob = Set-AzStorageBlobContent -Container $containerName -Blob $blobName -Context $connection -File $file_path

    Write-Output "Blob '$blobName' uploaded to container '$containerName'"
}

# Moves files from Temp location to Azure File Share
function move_files_to_azure_fileshare {
    param (
        [string] $fileShareName,
        [string] $sourcePath,
        [string] $fileName,
        [string] $destinationPath,
        [Microsoft.WindowsAzure.Commands.Storage.Common.AzureStorageContext] $connection

    )
    
    $destination = $destinationPath + '\' + $fileName
    $source = $sourcePath + '\' + $fileName
    $existingFile = Get-AzStorageFile -ShareName $fileShareName -Path $destination -Context $context -ErrorAction SilentlyContinue

    if ($existingFile) {
        Write-Host "File '$destination' already exists in the file share."
    } else {
        # File does not exist, proceed with upload
        Set-AzStorageFileContent -ShareName $fileShareName -Source $source -Path $destination -Context $connection
        Write-Host "File '$source' uploaded as '$destination'."
    }
}


# .  . 
# |  | 
# '--|-
#    | 
#    ' 
## ?

# .---.
# |    
# '--. 
# .   )
#  `-' 
## ?

# https://www.asciiart.eu/text-to-ascii-art Font: Swan
