# SharePoint Online Manager

## Commands

**Build a release and digitally sign the executable:**
```bash
.\Sign-App.ps1
```

**Build a release with signing:**
```bash
winget install -e --id GitHub.cli
gh auth login
.\Sign-App.ps1 -Version 1.0.1 -Release
```

**Run:**
```bash
dotnet run
```

**Publish (standalone EXE, no runtime required):**
```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

Output: `bin\Release\net8.0-windows\win-x64\publish\`


## Debugging
To view the trace output:                                                                                                                                                                                                                    
  1. Download DebugView from Microsoft: https://learn.microsoft.com/en-us/sysinternals/downloads/debugview                                                                                                                                  
  2. Run Dbgview.exe as Administrator
  3. In DebugView menu: Capture → Capture Global Win32 (enable it)
  4. Optionally add a filter: Edit → Filter/Highlight and enter SPOManager in Include field
  5. Run your app and click Re-authenticate


  You'll see output like:
  SPOManager
  [LoginForm] Navigation completed: https://ottawaheart-admin.sharepoint.com/...
  [LoginForm] Detected SharePoint site, capturing cookies for domain: ...
  [LoginForm] FedAuth: found (1234 chars)
  [LoginForm] Fetching current user email...
  [LoginForm] JavaScript result (raw): "sg_migration1@..."
