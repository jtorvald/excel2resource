package main

import (
	"errors"
	"github.com/fsnotify/fsnotify"
	"log"
	"os/signal"
	"os/user"
	"path/filepath"
	"sort"
	"syscall"

	"encoding/xml"
	"fmt"
	"os"

	"strings"

	"flag"

	"io/ioutil"

	"github.com/sirupsen/logrus"
	"github.com/tealeg/xlsx"
)

func main() {
	outputPath := flag.String("output", "", "Path to output the resource files or Excel file when using -invert=true")
	inputFile := flag.String("input", "", "Excel file or directory to watch. All .xlsx files will be parsed. With -invert=true this will read the .resx file(s)")
	watch := flag.Bool("watch", false, "Watch for file changes. If false, just execute once")
	invert := flag.Bool("invert", false, "Indicate that input file is .RESX and generate an Excel file as output")
	verbose := flag.Bool("v", false, "Verbose")
	trace := flag.Bool("vv", false, "Very verbose")

	flag.Parse()

	if *outputPath == "" {
		logrus.Error("Specify the output path for resource files: --output=PATH")
		os.Exit(1)
	}
	if *inputFile == "" {
		logrus.Error("Specify the Excel file or path: --input=PATH")
		os.Exit(2)
	}

	if *trace {
		logrus.StandardLogger().SetLevel(logrus.TraceLevel)
		logrus.Println("VERY VERBOSE MODE")
	} else if *verbose {
		logrus.StandardLogger().SetLevel(logrus.DebugLevel)
		logrus.Println("VERBOSE MODE")
	} else {
		logrus.StandardLogger().SetLevel(logrus.InfoLevel)
	}

	*outputPath = expandUserDirectory(*outputPath)
	*inputFile = expandUserDirectory(*inputFile)

	if _, err := os.Stat(*inputFile); errors.Is(err, os.ErrNotExist) {
		logrus.Error("Given Excel file/path does not exist")
		os.Exit(3)
	}

	if _, err := os.Stat(*outputPath); errors.Is(err, os.ErrNotExist) {
		logrus.Error("Given output path does not exist")
		os.Exit(3)
	}

	if *invert {
		success, err := importResx(*inputFile, *outputPath)
		if err != nil {
			panic(err)
		}

		if success {
			logrus.Info("Successfully converted the RESX to Excel")

		} else {
			logrus.Info("Could not convert RESX to Excel for some unknown reason")
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
				logrus.Debug("event:", event)
				if event.Op&fsnotify.Write == fsnotify.Write || event.Op&fsnotify.Create == fsnotify.Create || event.Op&fsnotify.Rename == fsnotify.Rename {

					if strings.HasSuffix(event.Name, ".xlsx") {
						logrus.Info("modified or created file:", event.Name)
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

				logrus.Debug("")
				logrus.Debug("Got interrupt....")

				break loop
			case syscall.SIGHUP:
				logrus.Debug("SIGHUP, need to re-read some configuration...")

			default:
				logrus.Debug("Unknown signal: ", sig)
			}
		}
	}

	logrus.Info("Shutting down...")
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

	baseDir, baseName, err := getPathInfo(resxFileName, os.PathSeparator)
	if err != nil {
		panic(err)
	}
	logrus.Trace("Base Dir", baseDir, "Base Name", baseName)
	//slashedPath := filepath.ToSlash(resxFileName)
	//baseDir := slashedPath[:strings.LastIndexByte(slashedPath, '/')] + "/"
	//fileName := strings.TrimPrefix(name, baseDir)

	matches, err := filepath.Glob(pattern)
	if err != nil {
		return false, err
	}
	orderedKeys := make([]string, 0)
	newData := data{
		sheetName:    baseName,
		cultureCodes: make([]string, 0),
		identifiers:  map[string]translation{},
	}
	for _, v := range root.Data {
		orderedKeys = append(orderedKeys, v.Name)
		newData.identifiers[v.Name] = translation{
			key:          v.Name,
			neutral:      v.Value,
			comment:      v.Comment,
			translations: map[string]string{},
		}
	}

	for _, localeFile := range matches {

		// ./Resx/Resources.se.resx
		code := strings.TrimPrefix(localeFile, filepath.Clean(baseDir)+string(os.PathSeparator))
		// Resources.se.resx
		code = strings.TrimPrefix(code, baseName+".")
		// se.resx
		code = strings.TrimSuffix(code, ext)
		// se

		newData.cultureCodes = append(newData.cultureCodes, code)
		logrus.Debug("Found locale file: ", localeFile, " Culture code: ", code)
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
			logrus.Trace("\\ Found translation for: ", v.Name, " Value: ", strings.Replace(v.Value, "\n", " ", 0))
			orderedKeys = append(orderedKeys, v.Name)

			if newData.identifiers[v.Name].translations == nil {

				logrus.Warn(v.Name, " does not exist in neutral language")
				newData.identifiers[v.Name] = translation{
					name:         baseName,
					key:          v.Name,
					neutral:      "MISSING",
					comment:      "WARNING",
					translations: map[string]string{},
				}
			}
			logrus.Trace("Adding to key ", v.Name, " translation code ", code, " value ", v.Value)
			newData.identifiers[v.Name].translations[code] = v.Value

		}
	}

	orderedKeys = unique(orderedKeys)

	sort.Strings(orderedKeys)

	newData.orderedKeys = orderedKeys

	err = writeExcelFile(filepath.Join(outputPath, baseName+".xlsx"), newData)
	if err != nil {
		return false, err
	}
	return true, nil
}

func unique(intSlice []string) []string {
	keys := make(map[string]bool)
	list := []string{}
	for _, entry := range intSlice {
		if _, value := keys[entry]; !value {
			keys[entry] = true
			list = append(list, entry)
		}
	}
	return list
}

func getPathInfo(path string, pathSeparator rune) (baseDir string, baseName string, err error) {

	ext := filepath.Ext(path)
	name := strings.TrimSuffix(path, ext)

	baseDir = path[:strings.LastIndex(path, string(pathSeparator))] + string(pathSeparator)
	baseName = strings.TrimPrefix(name, baseDir)

	return
}

func writeExcelFile(outputFile string, data data) error {

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

	var sht *xlsx.Sheet
	var err error
	// new file
	if len(wb.Sheets) == 0 {
		sht, err = wb.AddSheet(data.sheetName)
	} else {
		if _, exists := wb.Sheet[data.sheetName]; exists {
			sht = wb.Sheet[data.sheetName]
			for i := 0; i < sht.MaxRow; i++ {
				_ = sht.RemoveRowAtIndex(0)
			}
		} else {
			sht, err = wb.AddSheet(data.sheetName)
			if err != nil {
				panic(err)
			}
		}
	}
	row := sht.AddRow()
	row.AddCell().Value = "identifier"
	row.AddCell().Value = "description"
	row.AddCell().Value = "neutral"
	for _, v := range data.cultureCodes {
		row.AddCell().Value = v
	}

	for _, key := range data.orderedKeys {
		row := sht.AddRow()
		obj := data.identifiers[key]

		row.AddCell().Value = obj.key
		row.AddCell().Value = obj.comment
		row.AddCell().Value = obj.neutral
		for _, v := range data.cultureCodes {
			row.AddCell().Value = obj.translations[v]
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

type data struct {
	sheetName    string
	cultureCodes []string
	identifiers  map[string]translation
	orderedKeys  []string
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
