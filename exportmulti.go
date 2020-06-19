package hoboexcel

import (
	"archive/zip"
	"bufio"
	"io"
	"log"
	"math/rand"
	"os"
	"strconv"
	"time"
)

var letterRunes = []rune("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")

func RandStringRunes(n int) string {
	b := make([]rune, n)
	for i := range b {
		b[i] = letterRunes[rand.Intn(len(letterRunes))]
	}
	return string(b)
}
func ExportMultisheet(filename string, fetcher SheetFetcher) error {
	now := time.Now()
	sheetName := now.Format("20060102150405") //filename should be (pseudo)random
	shaStr, e := os.Create(sheetName + ".ss")
	if e != nil {
		log.Println(e)
		return e
	}
	//defer shaStr.Close()
	sharedStrWriter := bufio.NewWriter(shaStr)
	sharedStrWriter.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
	sharedStrWriter.WriteString("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\">")
	cellCount := 0
	outputFile := filename
	sheetNames := fetcher.GetSheetNames()
	file := make(map[string]io.Reader)
	file["_rels/.rels"] = DummyRelsDotRels()
	file["docProps/app.xml"] = AppXmlGenerator(sheetNames)
	file["docProps/core.xml"] = DummyCoreXml()
	file["xl/_rels/workbook.xml.rels"] = WorkbookRelGenerator(sheetNames)
	file["xl/theme/theme1.xml"] = DummyThemeXml()
	//change this crap
	//file["xl/worksheets/sheet1.xml"], _ = os.Open(sheetName)
	file["xl/styles.xml"] = DummyStyleXml()
	file["xl/workbook.xml"] = WorkbookXMLGenerator(sheetNames)

	file["[Content_Types].xml"] = ContentTypeGenerator(sheetNames)

	toCloseList := []string{}
	toDeleteList := []string{}
	sheetCounter := 1
	for true {
		curSheet := fetcher.NextSheet()
		if curSheet == nil {
			break
		}
		sheetName := now.Format("20060102150405") + RandStringRunes(5)
		e = ExportWorksheet(sheetName, curSheet, sharedStrWriter, &cellCount)
		if e != nil {
			log.Println(e)
			return e
		}
		file["xl/worksheets/sheet"+strconv.Itoa(sheetCounter)+".xml"], _ = os.Open(sheetName)
		toCloseList = append(toCloseList, "sheet"+strconv.Itoa(sheetCounter))
		toDeleteList = append(toDeleteList, sheetName)
		sheetCounter++
	}
	sharedStrWriter.WriteString("</sst>")
	sharedStrWriter.Flush()
	shaStr.Close()
	file["xl/sharedStrings.xml"], e = os.Open(sheetName + ".ss")
	if e != nil {
		log.Println(e)
		return e
	}
	of, e := os.Create(outputFile)
	if e != nil {
		log.Println(e)
		return e
	}
	defer of.Close()
	zipWriter := zip.NewWriter(of)
	for k, v := range file {
		fWriter, e := zipWriter.Create(k)
		if e != nil {
			log.Println(e)
			return e
		}
		_, e = io.Copy(fWriter, v)
		if e != nil {
			log.Println(e)
			return e
		}
	}
	for _, toClose := range toCloseList {
		(file["xl/worksheets/"+toClose+".xml"]).(*os.File).Close()
	}
	for _, toDelete := range toDeleteList {
		os.Remove("./" + toDelete)
	}
	e = zipWriter.Close()
	if e != nil {
		log.Println(e)
		return e
	}
	(file["xl/sharedStrings.xml"]).(*os.File).Close()
	e = os.Remove(sheetName + ".ss")
	if e != nil {
		log.Println(e)
		return e
	}
	return nil
}
