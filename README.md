# SharePoint Online Manager

## Commands

**Build:**
```bash
dotnet build
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
