// Copyright 2011-2015, The xlsx2csv Authors.
// All rights reserved.
// For details, see the LICENSE file.

package main

import (
	"encoding/csv"
	"errors"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"strings"
	"time"

	"github.com/tealeg/xlsx/v3"
)

func generateCSVFromXLSXFile(w io.Writer, excelFileName string, sheetIndex int, csvOpts csvOptSetter) (string, error) {
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		return "",err
	}
	sheetLen := len(xlFile.Sheets)
	switch {
	case sheetLen == 0:
		return "",errors.New("This XLSX file contains no sheets.")
	case sheetIndex >= sheetLen:
		return "",fmt.Errorf("No sheet %d available, please select a sheet between 0 and %d\n", sheetIndex, sheetLen-1)
	}
	cw := csv.NewWriter(w)
	if csvOpts != nil {
		csvOpts(cw)
	}
	sheet := xlFile.Sheets[sheetIndex]
	var vals []string
	err = sheet.ForEachRow(func(row *xlsx.Row) error {
		if row != nil {
			vals = vals[:0]
			err := row.ForEachCell(func(cell *xlsx.Cell) error {
				str, err := cell.FormattedValue()
				if err != nil {
					return err
				}
				chk:= cell.GetNumberFormat()
				
				if chk == "mm-dd-yy" {
					println("Get Date: ",chk,": ",str)
					t,_ := time.Parse("01-02-06",str)
					str = t.Format("2/1/2006");
					println("Convert to: ","d/m/YYYY",": ",str)
				}
				//str = `"`+str+`"`
				vals = append(vals, str)
				return nil
			})
			if err != nil {
				return err
			}
		}
		cw.Write(vals)
		return nil
	})
	if err != nil {
		return "",err
	}
	cw.Flush()
	return "",cw.Error()
}

type csvOptSetter func(*csv.Writer)

func main() {
	var (
		//outFile    = flag.String("o", "case_mst.csv", "filename to output to. -=stdout")
		sheetIndex = flag.Int("i", 0, "Index of sheet to convert, zero based")
		delimiter  = flag.String("d", "|", "Delimiter to use between fields")
	)
	flag.Usage = func() {
		fmt.Fprintf(os.Stderr, `%s
	dumps the given xlsx file's chosen sheet as a CSV,
	with the specified delimiter, into the specified output.

Usage:
	%s [flags] <xlsx-to-be-read>
`, os.Args[0], os.Args[0])
		flag.PrintDefaults()
	}

	flag.Parse()
	if flag.NArg() != 1 {
		flag.Usage()
		os.Exit(1)
	}
	out := os.Stdout
	csv2:= strings.TrimSuffix(flag.Arg(0),".xlsx")+".csv"
	//if !(*outFile == "" || *outFile == "-") {7
		var err error
		if out, err = os.Create(csv2); err != nil {
			log.Fatal(err)
	}
	defer func() {
		if closeErr := out.Close(); closeErr != nil {
			log.Fatal(closeErr)
		}
	}()

	if _,err := generateCSVFromXLSXFile(out, flag.Arg(0), *sheetIndex,
		func(cw *csv.Writer) { cw.Comma = ([]rune(*delimiter))[0]; cw.UseCRLF = true;},
	); err != nil {
		log.Fatal(err)
	}
	
}
