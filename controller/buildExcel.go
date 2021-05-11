package controller

import (
	"encoding/json"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"io/ioutil"
	"net/http"
	"strconv"
)

type Table struct {
	ColWidth  float64
	RowHeight float64
	Head      []interface{}
	HeadStyle string
	Data      [][]interface{}
	DataStyle string
}

func BuildExcel(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("content-type", "text/plain")
	var table Table

	err := r.ParseForm()
	if err != nil {
		_, _ = w.Write([]byte(err.Error()))
		return
	}
	body, err := ioutil.ReadAll(r.Body)

	if err != nil {
		_, _ = w.Write([]byte(err.Error()))
		return
	}
	if err = json.Unmarshal(body, &table); err != nil {
		_, _ = w.Write([]byte(err.Error()))
		return
	}
	excel, err := excel(&table)
	if err != nil {
		_, _ = w.Write([]byte(err.Error()))
		return
	}

	w.Header().Set("content-type", "application/vnd.ms-excel")

	if _, err := excel.WriteTo(w); err != nil {
		fmt.Fprintf(w, err.Error())
	}

}

func excel(t *Table) (*excelize.File, error) {

	sheet := "Sheet1"

	if t.RowHeight == 0 {
		t.RowHeight = 20
	}
	if t.ColWidth == 0 {
		t.ColWidth = 20
	}
	file := excelize.NewFile()

	col, _ := excelize.ColumnNumberToName(len(t.Head))

	if err := file.SetColWidth(sheet, "A", col, t.ColWidth); err != nil {
		return file, err
	}

	if err := file.SetRowHeight(sheet, 1, t.RowHeight); err != nil {
		return file, err
	}

	if err := file.SetSheetRow(sheet, "A1", &t.Head); err != nil {
		return file, err
	}

	headStyleId, err := file.NewStyle(t.HeadStyle)
	if err != nil {
		return file, err
	}

	if err := file.SetCellStyle(sheet, "A1", col+"1", headStyleId); err != nil {
		return file, err
	}
	dataStyleId, err := file.NewStyle(t.DataStyle)
	if err != nil {
		return file, err
	}
	if err := file.SetCellStyle(sheet, "A2", col+strconv.Itoa(len(t.Data)+1), dataStyleId); err != nil {
		return file, err
	}
	for index, row := range t.Data {
		cell, _ := excelize.CoordinatesToCellName(1, index+2)
		if err := file.SetSheetRow(sheet, cell, &row); err != nil {
			return file, err
		}
		if err := file.SetRowHeight(sheet, index+2, t.RowHeight); err != nil {
			return file, err
		}
	}
	return file, nil
}
