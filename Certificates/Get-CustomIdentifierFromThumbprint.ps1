#  Disclaimer:    This code is not supported under any Microsoft standard support program or service.
#                 This code and information are provided "AS IS" without warranty of any kind, either
#                 expressed or implied. The entire risk arising out of the use or performance of the
#                 script and documentation remains with you. Furthermore, Microsoft or the author
#                 shall not be liable for any damages you may sustain by using this information,
#                 whether direct, indirect, special, incidental or consequential, including, without
#                 limitation, damages for loss of business profits, business interruption, loss of business
#                 information or other pecuniary loss even if it has been advised of the possibility of
#                 such damages. Read all the implementation and usage notes thoroughly.

#Use PowerShell to convert Thumbprint to CustomIdentifier - Useful for matching which certificate is currently being used by SAML etc

#Get the Service Principal information
$SPNs =  Get-AzureADServicePrincipal -ObjectId "<Service Principal Object ID>" | select PreferredTokenSigningKeyThumbprint,KeyCredentials
#Convert the certificate thumbprint (PreferredTokenSigningKeyThumbprint) to base 64 string (This is the value that appears in the KeyCredential object)
$hasher = [System.Security.Cryptography.HashAlgorithm]::Create('sha256')
$hash = $hasher.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($SPNs.PreferredTokenSigningKeyThumbprint))
$PreferredTokenSigningKeyThumbprintCustomKeyIdentifier = [System.Convert]::ToBase64String($hash)
#Incase we have multiple certificates, find the active one (PreferredTokenSigningKeyThumbprint) - Matching the CustomKeyIdentifier and the Base 64 Thumbprint
#This will match for both the signing and verifying certs
foreach($key in $SPNs.KeyCredentials){
    if([System.Convert]::ToBase64String($key.CustomKeyIdentifier) -like $PreferredTokenSigningKeyThumbprintCustomKeyIdentifier){
        Write-Output "Match Thumbprint: $($SPNs.PreferredTokenSigningKeyThumbprint) ThumbprintBase64: $($PreferredTokenSigningKeyThumbprintCustomKeyIdentifier)"
        Write-Output "Certificate Expiry: $($key.EndDate)"
    }
}