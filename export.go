package hoboexcel

import (
	"archive/zip"
	"bufio"
	"encoding/xml"
	"fmt"
	"html"
	"io"
	"log"
	"os"
	"strconv"
	"strings"
	"time"
)

func CleanNonUtfAndControlChar(s string) string {
	s = strings.Map(func(r rune) rune {
		if r <= 31 { //if r is control character
			if r == 10 || r == 13 || r == 9 { //because newline
				return r
			}
			return -1
		}
		return r
	}, s)
	return s
}

func ExportWorksheet(filename string, rows RowFetcher, sharedStrWriter *bufio.Writer, cellsCount *int) error {
	file, e := os.Create(filename)
	if e != nil {
		log.Println(e)
		return e
	}
	defer file.Close()

	Writer := bufio.NewWriter(file)

	Writer.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">")
	Writer.WriteString("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><selection activeCell=\"A1\" sqref=\"A1\"/></sheetView></sheetViews>")
	Writer.WriteString("<sheetFormatPr defaultRowHeight=\"15\" x14ac:dyDescent=\"0.25\"/>")
	Writer.WriteString("<sheetData>")

	rowCount := 1
	//uniqueString := map[string]int{}
	//sortedUsedStr := []string{}
	//cellsCount := 0
	for {
		rawRow := rows.NextRow()
		if rawRow == nil {
			break
		}
		rr := row{}
		rr.R = rowCount
		for idx, val := range rawRow {
			colName := colCountToAlphaabet(idx)
			newCol := XlsxC{}
			newCol.T = "s"
			newCol.R = fmt.Sprintf("%s%d", colName, rowCount)

			newCol.V = strconv.Itoa(*cellsCount)
			*cellsCount++
			rr.C = append(rr.C, newCol)
			fmt.Println(val, html.EscapeString(CleanNonUtfAndControlChar(val)))
			sharedStrWriter.WriteString(fmt.Sprintf("<si><t>%s</t></si>", html.EscapeString(CleanNonUtfAndControlChar(val))))
		}
		rr.Spans = "1:10"
		rr.Descent = "0.25"
		bb, e := xml.Marshal(rr)
		if e != nil {
			log.Println(e)
			return e
		}
		pp, e := Writer.Write(bb)
		if e != nil {
			log.Println(e)
			return e
		}
		if pp != len(bb) {
			return fmt.Errorf("wrote %d, expect %d", pp, len(bb))
		}
		if rowCount%1000 == 0 {
			e = sharedStrWriter.Flush()
			if e != nil {
				log.Println(e)
				return e
			}
			e = Writer.Flush()
			if e != nil {
				log.Println(e)
				return e
			}
		}
		rowCount++

	}
	_, e = Writer.WriteString("</sheetData>")
	if e != nil {
		log.Println(e)
		return e
	}
	_, e = Writer.WriteString("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>")
	if e != nil {
		log.Println(e)
		return e
	}
	_, e = Writer.WriteString("</worksheet>")
	if e != nil {
		log.Println(e)
		return e
	}
	e = Writer.Flush()
	if e != nil {
		log.Println(e)
		return e
	}

	//write shared strings
	//sharedString := xlsxSST{}
	//sharedString.Count = len(sortedUsedStr)
	//sharedString.UniqueCount = len(sortedUsedStr)
	// for _, val := range sortedUsedStr {
	// 	ss := xlsxSI{}
	// 	ss.T = val
	// 	sharedString.SI = append(sharedString.SI, ss)
	// }

	// encoder := xml.NewEncoder(shaStr)
	// e := encoder.Encode(sharedString)
	// if e != nil {
	// 	fmt.Println(e.Error())
	// }
	return nil

}
func colCountToAlphaabet(idx int) string {
	var colName string
	if idx >= 26 {
		firstLetter := (idx / 26) - 1
		secondLetter := (idx % 26)
		colName = string(65+firstLetter) + string(65+secondLetter)
	} else {
		colName = string(65 + idx)
	}
	return strings.ToUpper(colName)
}

func Export(filename string, fetcher RowFetcher) error {
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
	e = ExportWorksheet(sheetName, fetcher, sharedStrWriter, &cellCount)
	if e != nil {
		log.Println(e)
		return e
	}
	_, e = sharedStrWriter.WriteString("</sst>")
	if e != nil {
		log.Println(e)
		return e
	}
	e = sharedStrWriter.Flush()
	if e != nil {
		log.Println(e)
		return e
	}
	outputFile := filename
	file := make(map[string]io.Reader)
	file["_rels/.rels"] = DummyRelsDotRels()
	file["docProps/app.xml"] = DummyAppXml()
	file["docProps/core.xml"] = DummyCoreXml()
	file["xl/_rels/workbook.xml.rels"] = DummyWorkbookRels()
	file["xl/theme/theme1.xml"] = DummyThemeXml()
	file["xl/worksheets/sheet1.xml"], _ = os.Open(sheetName)
	file["xl/styles.xml"] = DummyStyleXml()
	file["xl/workbook.xml"] = DummyWorkbookXml()
	file["xl/sharedStrings.xml"], _ = os.Open(sheetName + ".ss")
	file["[Content_Types].xml"] = DummyContentTypes()
	of, e := os.Create(outputFile)
	if e != nil {
		log.Println(e)
		return e
	}
	defer of.Close()
	zipWriter := zip.NewWriter(of)
	for k, v := range file {
		fWriter, _ := zipWriter.Create(k)
		_, e = io.Copy(fWriter, v)
		if e != nil {
			log.Println(e)
		}
	}
	e = zipWriter.Close()
	if e != nil {
		log.Println(e)
	}
	(file["xl/sharedStrings.xml"].(*os.File)).Close()
	(file["xl/worksheets/sheet1.xml"].(*os.File)).Close()
	e = os.Remove("./" + sheetName)
	if e != nil {
		log.Println(e)
	}
	e = shaStr.Close()
	if e != nil {
		log.Println(e)
		return e
	}
	e = os.Remove("./" + sheetName + ".ss")
	if e != nil {
		log.Println(e)
		return e
	}
	return nil
}
