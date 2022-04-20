package main

import (
	"errors"
	"github.com/fsnotify/fsnotify"
	"log"
	"os/signal"
	"os/user"
	"path/filepath"
	"syscall"

	"encoding/xml"
	"fmt"
	"os"

	"strings"

	"flag"

	"io/ioutil"

	"github.com/tealeg/xlsx"
)

func main() {
	outputPath := flag.String("output", "", "Path to output the resource files")
	excelFile := flag.String("input", "", "Excel file or directory to watch. All .xlsx files will be parsed")
	watch := flag.Bool("watch", false, "Watch for file changes. If false, just execute once")

	flag.Parse()

	if *outputPath == "" {
		fmt.Println("Specify the output path for resource files: --output=PATH")
		os.Exit(1)
	}
	if *excelFile == "" {
		fmt.Println("Specify the Excel file or path: --input=PATH")
		os.Exit(2)
	}

	*outputPath = expandUserDirectory(*outputPath)
	*excelFile = expandUserDirectory(*excelFile)

	if _, err := os.Stat(*excelFile); errors.Is(err, os.ErrNotExist) {
		fmt.Println("Given Excel file/path does not exist")
		os.Exit(3)
	}

	if _, err := os.Stat(*outputPath); errors.Is(err, os.ErrNotExist) {
		fmt.Println("Given output path does not exist")
		os.Exit(3)
	}

	if !*watch {
		processXlsx(*excelFile, *outputPath)
	}

	if *watch {
		watchForFileChanges(*excelFile, *outputPath)
	}

}

func watchForFileChanges(excelFile, outputPath string) {

	watcher, err := fsnotify.NewWatcher()
	if err != nil {
		log.Fatal(err)
	}
	defer watcher.Close()

	go func() {
		for {
			select {
			case event := <-watcher.Events:
				log.Println("event:", event)
				if event.Op&fsnotify.Write == fsnotify.Write || event.Op&fsnotify.Create == fsnotify.Create || event.Op&fsnotify.Rename == fsnotify.Rename {

					if strings.HasSuffix(event.Name, ".xlsx") {
						log.Println("modified or created file:", event.Name)
						processXlsx(event.Name, outputPath)
					}
				}
			case err := <-watcher.Errors:
				log.Println("error:", err)
			}
		}
	}()

	// TODO: watch a specific file from config or command line
	err = watcher.Add(excelFile)
	if err != nil {
		log.Fatal(err)
	}

	// Handle OS Signals
	ch := make(chan os.Signal, 5)

	signal.Notify(ch, syscall.SIGINT, syscall.SIGTERM, syscall.SIGHUP, syscall.SIGQUIT)

loop:
	for {
		select {
		case sig := <-ch:
			switch sig {
			case os.Interrupt:
				fallthrough
			case syscall.SIGQUIT:

				fmt.Println("")
				fmt.Println("Got interrupt....")

				break loop
			case syscall.SIGHUP:
				fmt.Println("SIGHUP, need to re-read some configuration...")

			default:
				fmt.Println("Uknown signal: ", sig)
			}
		}
	}

	fmt.Println("Shutting down...")
}

func expandUserDirectory(path string) string {
	usr, _ := user.Current()
	dir := usr.HomeDir

	if path == "~" {
		// In case of "~", which won't be caught by the "else if"
		path = strings.Replace(path, "~", dir, 1)
	} else if strings.HasPrefix(path, "~/") {
		// Use strings.HasPrefix so we don't match paths like
		// "/something/~/something/"
		path = filepath.Join(dir, path[2:])
	}
	return path
}

func processXlsx(excelFileName, outputPath string) {
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Fprint(os.Stderr, err)
		return
	}

	descriptions := map[string]string{}             // identifier : description
	languages := make(map[string]map[string]string) // language => identifier : value
	languageIndex := make(map[int]string)           // index => language

	for _, sheet := range xlFile.Sheets {
		rowIndex := 0
		for _, row := range sheet.Rows {
			rowIndex++

			cellIndex := 0
			// first row, extract langauges
			if rowIndex == 1 {

				for _, cell := range row.Cells {

					// skip first and second cell in first row (identifier / description)
					if cellIndex < 2 {
						cellIndex++
						continue
					}
					cellIndex++

					text := cell.String()
					if _, ok := languages[text]; !ok {
						languages[text] = make(map[string]string)
						languageIndex[len(languageIndex)] = text
					}
				}
				continue
			}

			cellIndex = 0
			identifier := ""
			for _, cell := range row.Cells {
				text := cell.String()

				if cellIndex == 0 {

					if len(text) == 0 || text == "-" {
						goto NEXT_ROW
					}
					identifier = text
					cellIndex++
					continue
				}

				if cellIndex == 1 {
					// identifier
					if _, ok := descriptions[text]; !ok {
						//if cell.GetStyle().Font.Color == "FFFF0000" {}
						descriptions[identifier] = text
					}
				}

				if lang, ok := languages[languageIndex[cellIndex-2]]; ok {
					lang[identifier] = text
				}

				cellIndex++
			}
		NEXT_ROW:
			//fmt.Println("")
		}

		headers := []ResHeader{
			{Key: "resmimetype", Value: ResHeaderValue{Value: "text/microsoft-resx"}},
			{Key: "version", Value: ResHeaderValue{Value: "2.0"}},
			{Key: "reader", Value: ResHeaderValue{Value: "System.Resources.ResXResourceReader, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"}},
			{Key: "writer", Value: ResHeaderValue{Value: "System.Resources.ResXResourceWriter, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"}},
		}

		for lang, data := range languages {

			p := ResRoot{Headers: headers}

			for identifier, word := range data {
				if len(word) == 0 {
					continue
				}
				p.Data = append(p.Data, ResData{Name: strings.Replace(identifier, " ", "_", -1), Space: "preserve", Value: word, Comment: descriptions[identifier]})
			}

			if xmlstring, err := xml.MarshalIndent(p, "", "    "); err == nil {
				xmlstring = []byte(xml.Header + string(xmlstring))
				filename := fmt.Sprintf("%s/%s.%s.resx", strings.TrimRight(outputPath, "/"), strings.Trim(sheet.Name, " "), strings.Trim(lang, ""))
				if lang == "neutral" {
					filename = fmt.Sprintf("%s/%s.resx", strings.TrimRight(outputPath, "/"), strings.Trim(sheet.Name, " "))
				}

				d1 := []byte(xmlstring)
				err := ioutil.WriteFile(filename, d1, 0644)
				if err != nil {
					log.Printf("error writing Filename: %s\n%s\n", filename, xmlstring)
				} else {
					log.Printf("written filename: %s\n", filename)
				}

			} else {
				log.Printf("error %v", err)
			}
		}
	}

}

// ResHeader ...
type ResHeader struct {
	Key   string         `xml:"name,attr"`
	Value ResHeaderValue `xml:",innerxml"`
}

// ResHeaderValue ...
type ResHeaderValue struct {
	XMLName xml.Name `xml:"value"`
	Value   string   `xml:",chardata"`
}

// ResData ...
type ResData struct {
	XMLName xml.Name `xml:"data"`
	Name    string   `xml:"name,attr"`
	Space   string   `xml:"xml:space,attr"`
	Value   string   `xml:"value"`
	Comment string   `xml:"comment,omitempty"`
}

// ResRoot ...
type ResRoot struct {
	XMLName xml.Name    `xml:"root"`
	Headers []ResHeader `xml:"resheader,omitempty"`
	Data    []ResData   `xml:"data,omitempty"`
}
