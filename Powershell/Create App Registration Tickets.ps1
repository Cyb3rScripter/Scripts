# Run each section separately, in order
# Export pfx and cer
# Import cer file to "Certificates & secrets" of Send-OAuth-MailMessage app registration

# Create certificate
$newCertSplat = @{
    DnsName           = 'Org-it.com'
    CertStoreLocation = 'Cert:\CurrentUser\My'
    NotAfter          = (Get-Date).AddMonths(13)
    KeySpec           = 'KeyExchange'
}
$mycert = New-SelfSignedCertificate @newCertSplat

# Export certificate to .pfx file
$exportCertSplat = @{
    FilePath = 'ExoCert.pfx'
    Password = $(ConvertTo-SecureString -String "JjBXchhKaeKaTSnaP2aywvwzlNmhDrWBAr56tfzvMFNzYmjGGTLjDJabd8vkFhq" -AsPlainText -Force)
}
$mycert | Export-PfxCertificate @exportCertSplat

# Export certificate to .cer file
$mycert | Export-Certificate -FilePath ExoCert.cer