#Get Environment Home Path
$path=$env:Home

# Load SharePoint Management Shell assemblies
$env:PSModulePath += ";$path\Modules\SharePoint Online Management Shell\"

# GET method: each querystring parameter is its own variable
if ($req_query_url) 
{
    $url = $req_query_url 
}
if ($req_query_useremail) 
{
    $userEmail = $req_query_useremail
}
if ($req_query_isadmin) 
{
    $isAdmin = $req_query_isadmin
}
$siteUrl = $env:TenantURL
$username = $env:GlobalAdminUserName
$connectionError= "False"
try
{
    #Getting Password from Vault from using Module GetSecretValueFromKeyVault
    $env:PSModulePath += ";$path\Modules\"
    Import-Module "LEGO\GetSecretValueFromKeyVault"
    $clientSecretPasswordValue = Get-SecretsFromKeyVault -vaultURI $env:CLIENTIDVAULTURI -endPoint $env:MSI_ENDPOINT -secret $env:MSI_SECRET
    $encpassword = convertto-securestring -String $clientSecretPasswordValue -AsPlainText -Force
    #Create Credential Object to use in Connect-SPOnline
    $cred = New-Object -typename System.Management.Automation.PSCredential($username, $encpassword)
    
    #Connect with Tenant 
    Connect-SPOService -Url $siteUrl -Credential $cred    
    Write-Output 'Connected' 
}
catch [System.Exception]
{
        $connectionError="True"
        Write-Output -f red $_.Exception.ToString() 
        Out-File -Encoding Ascii -FilePath $res -inputObject "Failure: Network Error while Connecting to Server. Please wait for 10-15 minutes and resubmit request."
}

if($connectionError -eq 'False')
{
try
{    
    Get-SPOSite -Identity $url  
    [Boolean]$boolValue = [System.Convert]::ToBoolean($isAdmin)
        try
        {
            #Set site collection Administrator
            Set-SPOUser -IsSiteCollectionAdmin $boolValue -LoginName $userEmail -Site $url
            if($isAdmin -eq 'True')
            {
            Out-File -Encoding Ascii -FilePath $res -inputObject "Success: $userEmail added as SCA to site $url"    
            }
            else
            {
                Out-File -Encoding Ascii -FilePath $res -inputObject "Success: $userEmail removed from SCA to site $url"    
            }   
        }
        catch [System.Exception]
        {
            Write-Output -f red $_.Exception.ToString() 
            Out-File -Encoding Ascii -FilePath $res -inputObject "Failure: "$_.Exception.ToString()
        }      
        
}
catch [System.Exception]
{
      Write-Output -f red "Failure: Invalid URL or Site Not Available"
      Out-File -Encoding Ascii -FilePath $res -inputObject "Failure - Invalid URL or Site Not Available"  
      Write-Output -f red  $_.Exception.ToString()    
  
}        
}