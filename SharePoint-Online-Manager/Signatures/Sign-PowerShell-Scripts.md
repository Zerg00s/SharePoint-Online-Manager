# Signing PowerShell Scripts with Azure Trusted Signing

Sign `.ps1` scripts so they show **CleverPoint Solutions Incorporated** as the publisher and work under `AllSigned` execution policy.

**Prerequisite:** Azure Trusted Signing account already set up (see `CERTIFICATE-MANUAL.md`).

---

## One-Time Setup

```powershell
# Install the PowerShell module
Install-Module -Name TrustedSigning -Scope CurrentUser

# Login to Azure
az login
```

---

## Sign a Single Script

```powershell
Invoke-TrustedSigning `
    -Endpoint "https://eus.codesigning.azure.net" `
    -CodeSigningAccountName "Clever-Artifact-Signing" `
    -CertificateProfileName "cleverpoint-public-trust" `
    -FilesFolder "C:\path\to\folder" `
    -FilesFolderFilter "MyScript.ps1" `
    -FileDigest SHA256 `
    -TimestampRfc3161 "http://timestamp.acs.microsoft.com" `
    -TimestampDigest SHA256
```

## Sign All Scripts in a Folder

```powershell
Invoke-TrustedSigning `
    -Endpoint "https://eus.codesigning.azure.net" `
    -CodeSigningAccountName "Clever-Artifact-Signing" `
    -CertificateProfileName "cleverpoint-public-trust" `
    -FilesFolder "C:\path\to\folder" `
    -FilesFolderFilter "*.ps1" `
    -FileDigest SHA256 `
    -TimestampRfc3161 "http://timestamp.acs.microsoft.com" `
    -TimestampDigest SHA256
```

---

## Verify a Signature

```powershell
Get-AuthenticodeSignature "C:\path\to\MyScript.ps1"
```

Expected output:
```
Status      : Valid
SignerCertificate : [Subject] CN=CleverPoint Solutions Incorporated...
```

---

## Important Notes

- **Sign last.** Any edit to the script after signing invalidates the signature.
- **Workflow:** Edit -> Test -> Sign -> Distribute.
- **Timestamped signatures remain valid** even after the certificate expires or you cancel the Azure subscription.
- The signature is appended as a comment block at the bottom of the `.ps1` file — do not edit or remove it.
