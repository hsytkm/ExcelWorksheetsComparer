# .NET8 artifacts
$dirPath = ".\.artifacts\publish\ExcelWorksheetsComparer\release_win-x64"

# フォルダを再帰的に削除
Remove-Item -Path $dirPath -Recurse

# リリース
dotnet publish -c Release -r win-x64 -p:PublishSingleFile=true --self-contained true -p:IncludeNativeLibrariesForSelfExtract=true

# 表示
explorer $dirPath
