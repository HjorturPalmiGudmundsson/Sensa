<#
.SYNOPSIS
Skrifar út ítarlegar upplýsingar um öll skilríki í notkun í IIS Web Sites.

.DESCRIPTION
Skrifar út í skrá upplýsingar um öll skilríki í IIS Web Sites s.s. nafn, útgáfuaðila, gildistíma, fingrafar og bindingu.

.NOTES
MSM 2020.10.30: Created

.PARAMETER
Enginn

.INPUTS
Ekkert

.OUTPUTS
Skrá með upplýsingum um skilríki í notkun í IIS, %USERPROFILE%\Desktop\o-get-iis-cert-bindings.%COMPUTERNAME%.yyyyMMdd_HHmmss.txt

.EXAMPLE
Bara keyra stefju, gæti þurft að gera með "Run as administrator" réttindum
#>

$webSites = Get-Website
$myLocalCerts = Get-ChildItem -Path CERT:\LocalMachine -Recurse

$certItems = @()
$certItemsExpiring = @()

#--- Loop through all IIS Web and FTP sites
foreach( $website in $webSites )
{
    #--- Loop through all IIS Web and FTP site bindings
    foreach( $webBinding in $webSite.Bindings )
    {
        #--- Check certificate for each IIS Web and FTP site binding, if applicable
        foreach( $bindingCollection in $webBinding.Collection )
        {
            $myLocalCert = $null
            $certThumbprint = $null 
            
            #--- Check certificate for each HTTP Web 
            if( $bindingCollection.protocol -like "http*" )
            {
                $protocolSite = $bindingCollection.protocol
                $stateSite = $website.state
                $certThumbprint = $bindingCollection.certificateHash
            }
            #--- Check certificate for each FTP binding 
            elseif( $bindingCollection.protocol -like "ftp*" )
            {
                $protocolSite = "ftps"
                $stateSite = $webSite.ftpServer.state
                $certThumbprint = $webSite.ftpServer.security.ssl.serverCertHash 
            }

            #-- Check if cert binding is specified
            if( $null -ne $certThumbprint -and "" -ne $certThumbprint )
            {
                #--- If cert exists, check trust and add it to the report 
                $myLocalCert = $myLocalCerts | Where-Object { $_.Thumbprint -eq $certThumbprint }
                if( $? -and $null -ne $myLocalCert )
                {
                    $certOK = $myLocalCert | Test-Certificate 

                    $sanExt = $myLocalCert.Extensions | Where-Object {$_.Oid.FriendlyName -eq "Subject Alternative Name"}
                    if( $null -eq $sanExt )
                    {
                        $sanExtExclNewLine = ""
                    }
                    else
                    {
                        $sanExtString = $sanExt.Format(1)
                        $sanExtStringExclNewLine = $sanExtString.Replace("`r`n",";")
                    }

                    $certItem = New-Object PSObject
                    $certItem | Add-Member NoteProperty ServerName    $env:COMPUTERNAME
                    $certItem | Add-Member NoteProperty FriendlyName  $myLocalCert.FriendlyName
                    $certItem | Add-Member NoteProperty OK            $certOK
                    $certItem | Add-Member NoteProperty ValidTo       $myLocalCert.NotAfter
                    $certItem | Add-Member NoteProperty WebSite       $website.name
                    $certItem | Add-Member NoteProperty Protocol      $protocolSite
                    $certItem | Add-Member NoteProperty State         $stateSite
                    $certItem | Add-Member NoteProperty Binding       $bindingCollection.bindingInformation
                    $certItem | Add-Member NoteProperty SAN           $sanExtStringExclNewLine
                    $certItem | Add-Member NoteProperty Thumbprint    $myLocalCert.Thumbprint
                    $certItem | Add-Member NoteProperty ValidFrom     $myLocalCert.NotBefore
                    $certItem | Add-Member NoteProperty Issuer        $myLocalCert.Issuer
                    $certItem | Add-Member NoteProperty Subject       $myLocalCert.Subject
                    $certItems += $certItem

					if( ($myLocalCert.NotAfter).AddDays(-30) -lt (Get-Date) )
					{
						$certItemsExpiring += $certItem
					}
                }
            }
        }
    }
}

$certLogFile = "{0}\Desktop\o-get-iis-cert-bindings.{1}.{2:yyyyMMdd_HHmmss}.txt" -f $env:USERPROFILE,$env:COMPUTERNAME,(Get-Date)
$certItems | Sort FriendlyName, WebSite | Format-Table -AutoSize | Out-File -Width 500 $certLogFile

if( $certItemsExpiring -ne $null )
{ 
    $certLogFileExpiring = "{0}\Desktop\o-get-iis-cert-bindings.{1}.expiring.txt" -f $env:USERPROFILE,$env:COMPUTERNAME
    $certItemsExpiring | Sort ValidTo, FriendlyName, WebSite | Format-Table -AutoSize | Out-File -Width 500 $certLogFileExpiring 
}
