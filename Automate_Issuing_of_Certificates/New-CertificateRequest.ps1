    param (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$True,HelpMessage = "Please enter the subject beginning with CN=")]
        [ValidatePattern("CN=")]
        [string]$Subject,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$True,HelpMessage = "Please enter the SAN domains as a comma separated list")]
        [string]$SAN,
        [Parameter(Mandatory=$false, HelpMessage = "Please enter the Online Certificate Authority")]
        [string]$OnlineCA,
        [Parameter(Mandatory=$false, HelpMessage = "Please enter the Certificate Template")]
        [string]$CATemplate = "WebServer",
		[Parameter(Mandatory=$false, HelpMessage = "Please specify if you want to export the certificate")]
        [switch]$Export,
		[Parameter(Mandatory=$false, HelpMessage = "Please specify password to protect PKCS#12 (PFX) bundle")]
		[string]$Password = "password"
    )
	
	process {
	
	$SANs = $SAN.split(',')
 
     ### Preparation
    $subjectDomain = $Subject.split(',')[0].split('=')[1]
    if ($subjectDomain -match "\*.") {
        $subjectDomain = $subjectDomain -replace "\*", "star"
    }
    $CertificateINI = "$subjectDomain.ini"
    $CertificateREQ = "$subjectDomain.req"
    $CertificateRSP = "$subjectDomain.rsp"
    $CertificateCER = "$subjectDomain.cer"
	$CertificatePFX = "$subjectDomain.pfx"
 
    ### INI file generation
    new-item -type file $CertificateINI -force
    add-content $CertificateINI '[Version]'
    add-content $CertificateINI 'Signature="$Windows NT$"'
    add-content $CertificateINI ''
    add-content $CertificateINI '[NewRequest]'
    $temp = 'Subject="' + $Subject + '"'
    add-content $CertificateINI $temp
    add-content $CertificateINI 'Exportable=TRUE'
    add-content $CertificateINI 'KeyLength=2048'
    add-content $CertificateINI 'KeySpec=1'
    add-content $CertificateINI 'KeyUsage=0xA0'
    add-content $CertificateINI 'MachineKeySet=True'
    add-content $CertificateINI 'ProviderName="Microsoft RSA SChannel Cryptographic Provider"'
    add-content $CertificateINI 'ProviderType=12'
    add-content $CertificateINI 'SMIME=FALSE'
    add-content $CertificateINI 'RequestType=PKCS10'
    add-content $CertificateINI '[Strings]'
    add-content $CertificateINI 'szOID_ENHANCED_KEY_USAGE = "2.5.29.37"'
    add-content $CertificateINI 'szOID_PKIX_KP_SERVER_AUTH = "1.3.6.1.5.5.7.3.1"'
    add-content $CertificateINI 'szOID_PKIX_KP_CLIENT_AUTH = "1.3.6.1.5.5.7.3.2"'
    if ($SANs) {
        add-content $CertificateINI 'szOID_SUBJECT_ALT_NAME2 = "2.5.29.17"'
        add-content $CertificateINI '[Extensions]'
        add-content $CertificateINI '2.5.29.17 = "{text}"'
 
        foreach ($item in $SANs) {
            $temp = '_continue_ = "dns=' + $item + '&"'
            add-content $CertificateINI $temp
        }
    }
 
    ### Certificate request generation
    if (test-path $CertificateREQ) {del $CertificateREQ}
    certreq -new $CertificateINI $CertificateREQ
 
    ### Online certificate request and import
    if ($OnlineCA) {
        if (test-path $CertificateCER) {del $CertificateCER}
        if (test-path $CertificateRSP) {del $CertificateRSP}
        certreq -submit -attrib "CertificateTemplate:$CATemplate" -config $OnlineCA $CertificateREQ $CertificateCER
        certreq -accept $CertificateCER
    }
	
	if($Export)
		{
		    Write-Debug "export parameter is set. => export certificate"
		    Write-Verbose "exporting certificate and private key"
		    $mypwd = ConvertTo-SecureString -String $Password -Force -AsPlainText
			$cert = Get-Childitem "cert:\LocalMachine\My" | where-object {$_.Thumbprint -eq (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2((Get-Item "$CertificateCER").FullName,"")).Thumbprint}
		    Write-Debug "Certificate found in computerstore: $cert"	
			
		    # write pfx file
			$cert | Export-PfxCertificate -FilePath .\$CertificatePFX -ChainOption "BuildChain" -Password $mypwd
		    Write-Host "Certificate successfully exported to $CertificatePFX !" -ForegroundColor Green
		    
		    Write-Verbose "Deleting exported certificate from computer store"
		    # delete certificate from computer store
		    $certstore = new-object system.security.cryptography.x509certificates.x509Store('My', 'LocalMachine')
		    $certstore.Open('ReadWrite')
		    $certstore.Remove($cert)
		    $certstore.close() 	
}

function Remove-ReqTempfiles()
	{
		param(
			[String[]]$tempfiles
		)
		Write-Verbose "Cleanup temp files and pending requests"
	    
		# delete pending request (if a request exists for the CN)
		$certstore = new-object system.security.cryptography.x509certificates.x509Store('REQUEST', 'LocalMachine')
		$certstore.Open('ReadWrite')
		foreach($certreq in $($certstore.Certificates))
		{
			if($certreq.Subject -eq "CN=$CN")
			{
				$certstore.Remove($certreq)
			}
		}
		$certstore.close()
		
		foreach($file in $tempfiles){remove-item ".\$file" -ErrorAction silentlycontinue}
	}
	
	Remove-ReqTempfiles -tempfiles $CertificateINI,$CertificateREQ,$CertificateRSP

}
