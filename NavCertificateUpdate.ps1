<# This is intended to be run as CertifyTheWeb Post-Renewal script #>

param($result)

. (Join-Path $PSScriptRoot 'NavSettings.ps1')
Import-Module (Join-Path $NavServiceRoot 'Microsoft.Dynamics.Nav.Management.dll')

$newCertThumbprint = $result.ManagedItem.CertificateThumbprintHash

$cert = Get-Item "Cert:\LocalMachine\My\$newCertThumbprint"

if ($cert -ne $null -and $cert.PrivateKey.CspKeyContainerInfo.UniqueKeyContainerName -ne $null) 
{
    #Grant permission on certificate to NAV service account 
    {
        # Get Location of the machine related keys
        $keyPath = $env:ProgramData + '\Microsoft\Crypto\RSA\MachineKeys\'; 
        $keyName = $cert.PrivateKey.CspKeyContainerInfo.UniqueKeyContainerName;
        $keyFullPath = $keyPath + $keyName;
        Write-Host "Found Certificate..." -ForegroundColor Green
        Write-Host "Granting access to $NavAccountName..." -ForegroundColor Green
        #Grant FullControl/Read/Write to account 
        $permissionType = 'Read'
        $acl = (Get-Item $keyFullPath).GetAccessControl('Access') #Get Current Access
        $buildAcl = New-Object  System.Security.AccessControl.FileSystemAccessRule($NavAccountName, $permissionType, 'Allow')
        $acl.SetAccessRule($buildAcl) #Add Access Rule
        Set-Acl $keyFullPath $acl #Save Access Rules
        Write-Host "Access granted to $NavAccountName..." -ForegroundColor Green   
    }

    Write-Host "Updating certificate Thumbprint for NAV instance $NavInstanceName..." -ForegroundColor Green   
    Set-NAVServerConfiguration -ServerInstance $NAVInstanceName `
        -KeyName 'ServicesCertificateThumbprint' -KeyValue $cert.Thumbprint `
        -Force

    Write-Host "Restarting NAV instance $NavInstanceName..." -ForegroundColor Green   
    Set-NAVServerInstance -ServerInstance $NAVInstanceName -Restart -Force

    Write-Host 'Done.' -ForegroundColor Green   
}
else
{
    Write-Host "Unable to find Certificate that matches thumbprint $newCertThumbprint or the private key is missing..." -ForegroundColor Red    
}

