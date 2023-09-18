package table

import (
	"bufio"
	"bytes"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"reflect"
)

type Options struct {
	SheetName    string // 表名
	Title        string // 标题
	ShowRemind   bool   //  显示提示
	DefaultStyle bool
	SwNum        int64 // 流式写入
}

type ExcellingEntity struct {
	Model  interface{}       // 模型
	Rows   map[string]*Field // 字段表格显示名/ 字段
	Fields map[string]*Field // 字段名称 / 字段
	Sw     *excelize.StreamWriter
	Option Options
}

func (e *ExcellingEntity) New(sheetName, title string, tips bool, model interface{}) {
	if sheetName != "" {
		e.Option.SheetName = sheetName
	}
	if title != "" {
		e.Option.Title = title
	}
	e.Model = model
	rows, fields, err := getField(model)
	if err != nil {
		log.Fatal(err)
	}
	e.Rows = rows
	e.Fields = fields
	e.Option.ShowRemind = tips
}

func (e *ExcellingEntity) Import(bytes []byte) (interface{}, error) {
	//TODO implement me
	panic("implement me")
}

func (e *ExcellingEntity) Export(i interface{}) ([]byte, error) {
	if !isEqual(e.Model, i) {
		return nil, UnKnowType
	}

	f := excelize.NewFile()
	_, err := f.NewSheet(e.Option.SheetName)
	if err != nil {
		return nil, CreateErr
	}
	e.Sw, err = f.NewStreamWriter(e.Option.SheetName)
	if err != nil {
		return nil, err
	}

	styleID, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#DFEBF6"}, Pattern: 1},
	})

	err = e.Sw.SetRow("A1",
		[]interface{}{excelize.Cell{Value: e.Option.Title, StyleID: styleID}},
		excelize.RowOpts{Height: 30, Hidden: false})
	if err != nil {
		return nil, err
	}
	vCell, _ := numberToLetters(len(e.Fields) - 1)
	if err := e.Sw.MergeCell("A1", vCell+"1"); err != nil {
		return nil, err
	}
	header := []interface{}{1, 2, 3, 4, 1, 23, 1}
	if err := e.Sw.SetRow("A2", header); err != nil {
		return nil, err
	}

	err = e.SetValue(i)
	if err != nil {
		return nil, err
	}
	var b bytes.Buffer
	writer := bufio.NewWriter(&b)
	f.Write(writer)
	return b.Bytes(), nil

}

func (e *ExcellingEntity) SetValue(data interface{}) error {
	rv := reflect.ValueOf(data)

	switch rv.Kind() {
	case reflect.Ptr:
		return e.SetValue(rv.Elem().Interface())
	case reflect.Slice:
		colList := make([]string, len(e.Fields))
		for k, v := range e.Fields {
			colList[v.Index] = k
		}
		for i := 3; i < rv.Len()+1; i++ {

			row := e.Sw.SetRow()
			for _, col := range colList {

				index := i - 1
				cell, _ := excelize.CoordinatesToCellName(1, i)
				cellValue := reflect.Value{}
				var cellVal interface{}
				if rv.Index(index).Kind() == reflect.Ptr {
					cellValue = reflect.ValueOf(rv.Index(index).Elem().Interface())
				} else {
					cellValue = reflect.ValueOf(rv.Index(index).Interface())
				}
				if !cellValue.FieldByName(e.Fields[col].FieldName).IsValid() {
					cellVal = reflect.Zero(e.Fields[col].Typ)
				} else {
					if e.Fields[col].Typ.Kind() == reflect.Float64 {
						cellVal = cellValue.FieldByName(e.Fields[col].FieldName).Float()
					} else {
						cellVal = cellValue.FieldByName(e.Fields[col].FieldName).Interface()
					}
				}
				e.Sw.SetRow(cell)
				cell.SetValue(cellVal)

			}
		}

	default:
		return UnKnowType
	}
	return nil
}

func (e *ExcellingEntity) StreamWriter() {
	//TODO implement me
	panic("implement me")
}

func (e *ExcellingEntity) Flush() {
	//TODO implement me
	panic("implement me")
}

func numberToLetters(num int) (string, error) {
	if num <= 0 {
		return "", fmt.Errorf("数字必须大于0")
	}

	// 定义字母表
	alphabet := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	base := len(alphabet)

	var result string

	for num > 0 {
		// 计算余数和商
		remainder := (num - 1) % base
		num = (num - 1) / base

		// 将余数对应的字母添加到结果字符串的前面
		result = string(alphabet[remainder]) + result
	}

	return result, nil
}
