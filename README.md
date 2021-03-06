# Excel2Resource #

This tool will convert sheets in an Excel file or a path full of Excel files to .NET resource files.
The reason this tool was developed was to give other non-technical people the possibility to translate resources from a
mobile app in Xamarin without needing to edit resource files. Everybody knows how to use Excel, right?

[![GitHub go.mod Go version of a Go module](https://img.shields.io/github/go-mod/go-version/jtorvald/excel2resource)](https://github.com/jtorvald/excel2resource)
[![GoReportCard](https://goreportcard.com/badge/github.com/jtorvald/excel2resource)]((https://goreportcard.com/report/github.com/jtorvald/excel2resource))


## Download releases ##

The best is to [download the binary for your platform and run the executable](https://github.com/jtorvald/excel2resource/releases).

To build from source, you'll need to clone the repository and run `go build -o ./bin/Excel2Resource main.go`
from within the cmd directory.

## How to run once ## 

To run the command only once to generate the resources, you can use:

```
./Excel2Resource --output=~/Projects/Mobile.App/Resx/ --input=~/Projects/Mobile.App/Translations.xlsx
```

## Watch for changes ##

To watch a file or directory and generate on each save.

Command
```
./Excel2Resource --output=~/Projects/Mobile.App/Resx/ --input=~/Projects/Mobile.App/Translations.xlsx --watch=true
```

## Replacement rules ##
- It will remove spaces in the sheet name to use as a file name.
- It will replace spaces in identifiers (first column) with underscores `_`.
- It will skip rows that are empty or start with `'-`.

## Template ##
The [example template](template.xlsx) from this repo will result in:
- AppResources.resx
- AppResources.en-UK.resx
- AppResources.nl.resx
- AppResources.de.resx
- AppResources.es.resx

If you want to ignore a row it needs to be empty or start with a dash. You can achieve this in Excel with entering `'-` in a cell. 

## Convert resource file into XLSX file ##
The following command reads the resource file Translations.resx
(and the specific language translations .nl.resx, .en-UK.resx, .de.resx, .es.resx)
and saves it as Translations.xlsx. 

**This functionality is not fully tested yet.**
```
./Excel2Resource -input=./Resx/Translations.resx -output=. -invert=true
```

## Disclaimer ##

I don't think it will do any harm but just in case: **use this tool at your own risk**. 

# License #
[MIT](LICENSE)

[![made-with-Go](https://img.shields.io/badge/Made%20with-Go-1f425f.svg)](https://go.dev/)
