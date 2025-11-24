# Sharepoint 365 Indexer

Indexes Sharepoint 365 sites and documents, outputting text to the console and JSON lines to a file, all only using stolen browser cookies. Also includes support for downloading files en mass. Genuinely works better than the terrible search function provided by Microsoft

## Building

Build the indexer using GoLang `1.24.4` or later:

```shell
go build sharepoint.go
```

## Usage

The SharePoint indexer expects the `fedAuth` and `rtFa` cookies from a browser session. You can obtain these cookies using browser developer tools, or by stealing cookies from a compromised machine using a tool like [ChromeKatz](https://github.com/Meckazin/ChromeKatz). Get the help page with:

```shell
./sharepoint -help
```

To list available sharepoint sites:

```shell
./sharepoint -cookies "fedAuth=...; rtFa=...;" -url https://example.sharepoint.com -list
```

Index a specific site

```shell
./sharepoint -cookies "fedAuth=...; rtFa=...;" -url https://example.sharepoint.com -site https://example.sharepoint.com/sites/blargh
```

Index all sites and write text to console, and JSON lines to a file called `output.jsonl` so it can be processed using a tool like `jq`:

```shell
./sharepoint -cookies "fedAuth=...; rtFa=...;" -url https://example.sharepoint.com -output output.jsonl -format text -fformat jsonl
```

To download a file or directory of files, use the `-download` flag:

```shell
./sharepoint -cookies "fedAuth=...; rtFa=...;" -download "https://example.sharepoint.com/sites/blargh/Shared Documents/example.txt"
./sharepoint -cookies "fedAuth=...; rtFa=...;" -download "https://example.sharepoint.com/sites/blargh/Shared Documents/foldername"
```
