package table

import (
	"bufio"
	"bytes"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"reflect"
	"strconv"
)

type Options struct {
	SheetName    string // 表名
	Title        string // 标题
	ShowRemind   bool   //  显示提示
	DefaultStyle bool
	SwNum        int64 // 流式写入
}

type ExcellingEntity struct {
	Model    interface{}       // 模型
	Rows     map[string]*Field // 字段表格显示名/ 字段
	Fields   map[string]*Field // 字段名称 / 字段
	Sw       *excelize.StreamWriter
	Option   Options
	RowStyle excelize.Style
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

func (e *ExcellingEntity) Import(data []byte) (interface{}, error) {
	r := bytes.NewReader(data)
	f, err := excelize.OpenReader(r)
	if err != nil {
		return nil, err
	}
	index, err := f.GetSheetIndex(e.Option.SheetName)
	if err != nil {
		return nil, err
	}
	f.SetActiveSheet(index)

	f.GetRows()
}

func (e *ExcellingEntity) Export(i interface{}) ([]byte, error) {
	if !isEqual(e.Model, i) {
		return nil, UnKnowType
	}

	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			return
		}
	}()

	index, err := f.NewSheet(e.Option.SheetName)
	if err != nil {
		return nil, CreateErr
	}

	e.Sw, err = f.NewStreamWriter(e.Option.SheetName)
	if err != nil {
		return nil, err
	}
	err = e.Sw.SetColWidth(1, 4, 20)
	if err != nil {
		return nil, err
	}
	titleStyle, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#DFEBF6"}, Pattern: 1},
		Font: &excelize.Font{
			Bold: true,
			Size: 25,
		},
	})

	err = e.Sw.SetRow("A1",
		[]interface{}{excelize.Cell{Value: e.Option.Title, StyleID: titleStyle}},
		excelize.RowOpts{Height: 30, Hidden: false})
	if err != nil {
		return nil, err
	}
	vCell, _ := numberToLetters(len(e.Fields))
	if err := e.Sw.MergeCell("A1", vCell+"1"); err != nil {
		return nil, err
	}

	err = e.SetValue(i)
	if err != nil {
		return nil, err
	}
	e.Sw.Flush()
	f.SetActiveSheet(index)
	f.DeleteSheet("Sheet1")
	var b bytes.Buffer
	writer := bufio.NewWriter(&b)
	f.Write(writer)
	return b.Bytes(), nil
}

func (e *ExcellingEntity) SetValue(data interface{}) error {
	rv := reflect.ValueOf(data)
	rangeBottoms := "B"
	switch rv.Kind() {
	case reflect.Ptr:
		return e.SetValue(rv.Elem().Interface())
	case reflect.Slice:
		colList := make([]string, len(e.Fields))
		rangeBottoms, _ = numberToLetters(len(colList))
		rangeBottoms = rangeBottoms + strconv.Itoa(rv.Len()+2)
		for k, v := range e.Fields {
			colList[v.Index] = k
		}
		for i := 2; i < rv.Len()+3; i++ {
			vals := make([]interface{}, 0)
			cell, _ := excelize.CoordinatesToCellName(1, i)
			for _, col := range colList {
				var cellValue interface{}
				if i == 2 {
					cellValue = e.Fields[col].Name
				} else {
					index := i - 3
					x := reflect.ValueOf(rv.Index(index).Interface()).FieldByName(e.Fields[col].FieldName).Interface()
					cellValue = x
				}
				vals = append(vals, cellValue)
			}

			err := e.Sw.SetRow(cell, vals)
			if err != nil {
				return err
			}
		}

	default:
		return UnKnowType
	}

	err := e.Sw.AddTable(&excelize.Table{
		Range:             "A2:" + rangeBottoms,
		Name:              "table",
		StyleName:         "TableStyleMedium2",
		ShowFirstColumn:   true,
		ShowLastColumn:    true,
		ShowColumnStripes: true,
	})
	if err != nil {
		return err
	}
	return nil
}

func (e *ExcellingEntity) StreamWriter() {
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

// Newlize 获取excel 实体对象
func Newlize(sheetName, title string, tips bool, model interface{}) *ExcellingEntity {
	e := new(ExcellingEntity)
	e.New(sheetName, title, tips, model)
	return e
}

func (e *ExcellingEntity) ExportFile(i interface{}) error {
	if !isEqual(e.Model, i) {
		return UnKnowType
	}

	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			return
		}
	}()

	index, err := f.NewSheet(e.Option.SheetName)
	if err != nil {
		return CreateErr
	}

	e.Sw, err = f.NewStreamWriter(e.Option.SheetName)
	if err != nil {
		return err
	}
	err = e.Sw.SetColWidth(1, 4, 20)
	if err != nil {
		return err
	}
	titleStyle, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#DFEBF6"}, Pattern: 1},
		Font: &excelize.Font{
			Bold: true,
			Size: 25,
		},
	})

	err = e.Sw.SetRow("A1",
		[]interface{}{excelize.Cell{Value: e.Option.Title, StyleID: titleStyle}},
		excelize.RowOpts{Height: 30, Hidden: false})
	if err != nil {
		return err
	}
	vCell, _ := numberToLetters(len(e.Fields))
	if err := e.Sw.MergeCell("A1", vCell+"1"); err != nil {
		return err
	}

	err = e.SetValue(i)
	if err != nil {
		return err
	}
	e.Sw.Flush()
	f.SetActiveSheet(index)
	f.DeleteSheet("Sheet1")
	if err := f.SaveAs(e.Option.SheetName + ".xlsx"); err != nil {
		return err
	}
	return nil
}
