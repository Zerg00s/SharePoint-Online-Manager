# Code Signing with Azure Artifact Signing

Sign `SharePoint-Online-Manager.exe` so Windows shows **CleverPoint Solutions Incorporated** as the verified publisher.

**Cost:** $9.99/month (Azure Artifact Signing Basic SKU, 5,000 signatures included)

---

## What's Already Done

- [x] Created Azure Artifact Signing account: **Clever-Artifact-Signing**
- [x] Submitted Organization identity validation for **CleverPoint Solutions Incorporated**
- [x] Installed Azure CLI (`az`)
- [x] Installed Artifact Signing Client Tools
- [x] Assembly metadata added to `.csproj` (Company, Authors, etc.)
- [x] `Sign-App.ps1` signing script created
- [x] `signing-metadata.json` template created (in project root, gitignored)

---

## Remaining Steps

### Step 1: Complete Identity Verification

Microsoft will email you when your identity validation status changes to **"Action Required"** (1-7 business days).

1. Go to Azure Portal > your Artifact Signing account > **Identity validations**
2. Click on your validation > follow the prompts
3. You'll need to verify your personal identity using:
   - Government-issued photo ID
   - Microsoft Authenticator app on your phone
4. You may also need to upload business documents (articles of incorporation, etc.)
5. Wait for status to change to **"Completed"**

### Step 2: Create Certificate Profile

Once identity validation is complete:

1. Azure Portal > your Artifact Signing account > **Certificate profiles** > **Create**
2. **Profile type:** Public Trust
3. **Name:** choose a name (e.g. `cleverpoint-public-trust`)
4. **Identity validation:** select your completed "CleverPoint Solutions Incorporated" validation
5. Click **Create**

The certificate will show: `CN=CleverPoint Solutions Incorporated, O=CleverPoint Solutions Incorporated`

### Step 3: Assign Signer Role

1. Azure Portal > your Artifact Signing account > **Access Control (IAM)**
2. Click **+ Add** > **Add role assignment**
3. Select **Artifact Signing Certificate Profile Signer** > Next
4. **Assign access to:** User, group, or service principal
5. Click **+ Select members** > find your account > Select
6. **Review + assign**

(You should already have the **Artifact Signing Identity Verifier** role from earlier.)

### Step 4: Update signing-metadata.json

Open `signing-metadata.json` in the project root and fill in your certificate profile name:

```json
{
  "Endpoint": "https://eus.codesigning.azure.net",
  "CodeSigningAccountName": "Clever-Artifact-Signing",
  "CertificateProfileName": "cleverpoint-public-trust"
}
```

Replace `cleverpoint-public-trust` with whatever name you chose in Step 2.

### Step 5: Sign the EXE

Open PowerShell in the project root and run:

```powershell
az login
.\Sign-App.ps1
```

This will:
1. Build the project in Release mode
2. Sign `SharePoint-Online-Manager.exe` using Azure Artifact Signing
3. Verify the signature

The signed exe will be at: `SharePoint-Online-Manager\bin\Release\net8.0-windows\SharePoint-Online-Manager.exe`

To skip the build and just re-sign an existing exe:

```powershell
.\Sign-App.ps1 -SkipBuild
```

---

## Verifying the Signature

After signing:

1. Right-click the `.exe` > **Properties** > **Digital Signatures** tab
2. Should show **CleverPoint Solutions Incorporated**
3. Click **Details** > **View Certificate** to see the full chain:
   - Microsoft ID Verified Code Signing PCA 2021
   - Microsoft Identity Verification Root CA 2020

---

## Troubleshooting

**"SignTool.exe not found"**
```powershell
winget install -e --id Microsoft.Azure.ArtifactSigningClientTools
```

**"Azure.CodeSigning.Dlib.dll not found"**
Same install as above. The dlib is included with the Artifact Signing Client Tools.

**"Signing failed" / authentication errors**
```powershell
az login   # re-authenticate
az account show   # verify correct subscription
```

**"Access denied" or role errors**
Verify you have the **Artifact Signing Certificate Profile Signer** role assigned (Step 3).

**SmartScreen still shows warning**
This is normal initially. SmartScreen reputation builds over time with a new certificate. The publisher name will still show correctly.
