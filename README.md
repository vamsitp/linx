# linx
**Extract and save links from PowerPoint/Word as Markdown/Html/Csv!**

---

> **USAGE**: 

**`linx`** `DeckWithLinks.pptx [/DocWithLinks.docx]` `md [/html/csv]`

---

> **PRE-REQ**: [.NET Core SDK](https://dotnet.microsoft.com/download/dotnet-core/3.0)

```batch
# Install from nuget.org
dotnet tool install -g linx

# Upgrade to latest version from nuget.org
dotnet tool update -g linx --no-cache

# Install a specific version from nuget.org
dotnet tool install -g linx --version 1.0.x

# Uninstall
dotnet tool uninstall -g linx
```

> **NOTE**: If the Tool is not accessible post installation, add `%USERPROFILE%\.dotnet\tools` to the PATH env-var.

---

> ##### CONTRIBUTION
> 
```batch
# Install from local project path
dotnet tool install -g --add-source ./bin linx

# Publish package to nuget.org
nuget push ./bin/Linx.1.0.0.nupkg -ApiKey <key> -Source https://api.nuget.org/v3/index.json
```