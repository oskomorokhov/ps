# Purpose
Automate generation of SSL/TLS certificates in a Microsoft-PKI-enabled infrastructure.

Only works on WIN platform at this time.

# How it works
- Script automatically makes a request to corporate PKI which generates the certificate bundle based on the arguments provided.
- Script either gathers the required arguments directly from CLI or pipes the CSV file with arguments as input (CA/PKI to utilize, certificate properties etc)
- Script outputs the result PFX bundle file (PCKS#12 format) to **working directory** (if **-Export** parameter has been passed).
- Script protects the result PFX bundle with **default passphrase "password"**. Use **-Password** parameter if you want to modify the passphrase string.

Refer to usage section for more details. 

# Requirements
- Powershell installed
- User account executing the powershell script has permissions to access corporate PKI server in order to utilize StandardWebServer (or similar) certificate template
- Powershell started as admin user
- Navigate to the folder where the script file(s) reside or add that folder to system's PATH variable

# Usage
- Pass arguments via powershell command line, generate 1 cert at a time
```
.\New-CertificateRequest.ps1 -subject "CN=host1.domain.com,C=CC,s=StateName,l=CityName,o=CompanyName,ou=DepartmentName" -SAN host1.domain.com -OnlineCA "subca01.domain.com\Enterprise Issuing CA" -CATemplate "StandardWebServer" -Export -Password "12345"
```
- Pass main arguments via input CSV file
```
Import-Csv -UseCulture .\cert_params.csv | .\New-CertificateRequest.ps1 -OnlineCA "subca01.domain.com\Enterprise Issuing CA" -CATemplate "StandardWebServer" -Export -Password "12345"
```
csv file structure
```
subject,SAN
"CN=host1.domain.com,C=CC,s=StateName,l=CityName,o=CompanyName,ou=DepartmentName","host1.domain.com"
"CN=host2.domain.com,C=CC,s=StateName,l=CityName,o=CompanyName,ou=DepartmentName","host2.domain.com"
"CN=*.subdomain.domain.com,C=CC,s=StateName,l=CityName,o=CompanyName,ou=DepartmentName","*.subdomain.domain.com"
```
# Links

- [Powershell basics](https://docs.microsoft.com/en-us/sccm/develop/core/understand/windows-powershell-basics)
- [PKI Basics](https://docs.microsoft.com/en-us/windows/desktop/seccertenroll/public-key-infrastructure)
