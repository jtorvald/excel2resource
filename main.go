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
	outputPath := flag.String("output", "", "Path to output the resource files or Excel file when using -invert=true")
	inputFile := flag.String("input", "", "Excel file or directory to watch. All .xlsx files will be parsed. With -invert=true this will read the .resx file(s)")
	watch := flag.Bool("watch", false, "Watch for file changes. If false, just execute once")
	invert := flag.Bool("invert", false, "Indicate that input file is .RESX and generate an Excel file as output")

	flag.Parse()

	if *outputPath == "" {
		fmt.Println("Specify the output path for resource files: --output=PATH")
		os.Exit(1)
	}
	if *inputFile == "" {
		fmt.Println("Specify the Excel file or path: --input=PATH")
		os.Exit(2)
	}

	*outputPath = expandUserDirectory(*outputPath)
	*inputFile = expandUserDirectory(*inputFile)

	if _, err := os.Stat(*inputFile); errors.Is(err, os.ErrNotExist) {
		fmt.Println("Given Excel file/path does not exist")
		os.Exit(3)
	}

	if _, err := os.Stat(*outputPath); errors.Is(err, os.ErrNotExist) {
		fmt.Println("Given output path does not exist")
		os.Exit(3)
	}

	if *invert {
		success, err := importResx(*inputFile, *outputPath)
		if err != nil {
			panic(err)
		}

		if success {
			fmt.Println("successfully converted the RESX to Excel")

		} else {
			fmt.Println("could not convert RESX to Excel for some unknown reason")
		}

	} else {

		if !*watch {
			processXlsx(*inputFile, *outputPath)
		}

		if *watch {
			watchForFileChanges(*inputFile, *outputPath)
		}

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

func readFile(file string) ([]byte, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}

	defer f.Close()
	byteValue, err := ioutil.ReadAll(f)
	if err != nil {
		return nil, err
	}
	return byteValue, nil
}

func importResx(resxFileName, outputPath string) (bool, error) {
	byteValue, err := readFile(resxFileName)
	if err != nil {
		return false, err
	}

	root := &ResRoot{}
	err = xml.Unmarshal(byteValue, root)
	if err != nil {
		return false, err
	}

	// find all locale files for this resource file
	// split file name and extension, search matches, parse matches.
	ext := filepath.Ext(resxFileName)
	name := strings.TrimSuffix(resxFileName, ext)
	pattern := name + ".*" + ext

	slashedPath := filepath.ToSlash(resxFileName)
	baseDir := slashedPath[:strings.LastIndexByte(slashedPath, '/')] + "/"
	fileName := strings.TrimPrefix(name, baseDir)

	matches, err := filepath.Glob(pattern)
	if err != nil {
		return false, err
	}

	newData := make(map[string]translation, 0)
	for _, v := range root.Data {
		newData[v.Name] = translation{
			key:          v.Name,
			neutral:      v.Value,
			comment:      v.Comment,
			translations: map[string]string{},
		}

		//fmt.Println(k, v)
		_ = v
	}

	for _, localeFile := range matches {
		// ./Resx/Resources.se.resx
		code := strings.TrimPrefix(localeFile, filepath.Clean(baseDir)+string(os.PathSeparator))
		// Resources.se.resx
		code = strings.TrimPrefix(code, fileName+".")
		// se.resx
		code = strings.TrimSuffix(code, ext)
		// se

		byteValue, err := readFile(localeFile)
		if err != nil {
			return false, err
		}

		translationRoot := &ResRoot{}
		err = xml.Unmarshal(byteValue, translationRoot)
		if err != nil {
			return false, err
		}
		for _, v := range translationRoot.Data {
			newData[v.Name].translations[code] = v.Value
		}
	}

	//for k, v := range newData {
	//	fmt.Println(k, v.key, v.neutral, v.translations, v.comment)
	//}

	err = writeExcelFile(filepath.Join(outputPath, fileName+".xlsx"), fileName, newData)
	if err != nil {
		return false, err
	}
	return true, nil
}

func writeExcelFile(outputFile, name string, data map[string]translation) error {
	var wb *xlsx.File
	if _, err := os.Stat(outputFile); errors.Is(err, os.ErrNotExist) {
		wb = xlsx.NewFile()

	} else {
		// open an existing file
		wb, err = xlsx.OpenFile(outputFile)
		if err != nil {
			return err
		}
	}

	if len(wb.Sheets) == 0 {
		wb.AddSheet(name)
	} else {
		for i := 0; i < wb.Sheets[0].MaxRow; i++ {
			wb.Sheets[0].RemoveRowAtIndex(0)
		}
	}

	row := wb.Sheets[0].AddRow()
	row.AddCell().Value = "identifier"
	row.AddCell().Value = "description"
	row.AddCell().Value = "neutral"
	for _, v := range data {
		for locale, _ := range v.translations {
			row.AddCell().Value = locale
		}
		break
	}

	for _, v := range data {
		row := wb.Sheets[0].AddRow()

		row.AddCell().Value = v.key
		row.AddCell().Value = v.comment
		row.AddCell().Value = v.neutral
		for _, value := range v.translations {
			row.AddCell().Value = value
		}
	}
	return wb.Save(outputFile)
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

type translation struct {
	name         string
	key          string
	neutral      string
	comment      string
	translations map[string]string
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
