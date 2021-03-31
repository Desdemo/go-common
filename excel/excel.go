package excel

import (
	"bufio"
	"bytes"
	"errors"
	"fmt"
	"log"
	"reflect"
	"strconv"
	"strings"

	"github.com/desdemo/go-common/orm"
	"github.com/gogf/gf/os/gtime"
	"github.com/tealeg/xlsx"
)

// 导出
type Excel interface {
	New(sheetName, title string, tips bool, model interface{})
	// 导入
	Import([]byte) (interface{}, error)
	// 导出
	Export(interface{}) ([]byte, error)
}

type Entity struct {
	Model      interface{}       // 模型
	SheetName  string            // 表名
	Title      string            // 标题
	Rows       map[string]*Field // 字段表格显示名/ 字段
	Fields     map[string]*Field // 字段名称 / 字段
	ShowRemind bool              //  显示提示
}

type Field struct {
	Name      string              // 显示名称
	FieldName string              // 字段名称
	Value     []interface{}       // 值
	UqiMap    map[string]struct{} // 唯一值
	Remind    string              // 提示
	Uqi       bool                // 唯一
	Required  bool                // 必填
	Typ       reflect.Type        // 类型
	Index     int                 // 索引
}

type Site struct {
	FieldName string // 字段名称
	ShowName  string // 显示名称
	Index     int    // 索引
}

var (
	TypeParseErr = errors.New("数据转换出错")
	UqiErr       = errors.New("数据存在重复")
	CreateErr    = errors.New("创建表信息失败")
)

func (e *Entity) New(sheetName, title string, tips bool, model interface{}) {
	if sheetName != "" {
		e.SheetName = sheetName
	}
	if title != "" {
		e.Title = title
	}
	e.Model = model
	rows, fields, err := getField(model)
	if err != nil {
		log.Fatal(err)
	}
	e.Rows = rows
	e.Fields = fields
	e.ShowRemind = tips
}

func (e *Entity) Import(bytes []byte) (interface{}, error) {
	// Unmarshal
	file, err := xlsx.OpenBinary(bytes)
	if err != nil {
		return nil, err
	}
	var sheet *xlsx.Sheet
	if _, ok := file.Sheet[e.SheetName]; ok {
		sheet = file.Sheet[e.SheetName]
	} else {
		sheet = file.Sheets[0]
	}
	data, err := e.ReadValue(sheet)
	if err != nil {
		return nil, err
	}
	return data, nil
}

func (e *Entity) Export(i interface{}) ([]byte, error) {
	if isEqual(e.Model, i) {
		file := xlsx.NewFile()
		sheet, err := e.addSheet(file, i)
		if err != nil {
			return nil, err
		}
		if sheet == nil {
			return nil, CreateErr
		}
		err = e.SetValue(sheet, i)
		if err != nil {
			return nil, err
		}
		var b bytes.Buffer
		writer := bufio.NewWriter(&b)
		file.Write(writer)
		return b.Bytes(), nil
	}
	return nil, nil
}

func (e *Entity) addSheet(wb *xlsx.File, data interface{}) (*xlsx.Sheet, error) {

	sh, err := wb.AddSheet(e.SheetName)
	if err != nil {
		return nil, err
	}
	return sh, nil
}

func isEqual(model, data interface{}) bool {
	pv := reflect.ValueOf(model)
	rv := reflect.ValueOf(data)
	switch rv.Kind() {
	case reflect.Slice:
		if rv.Len() > 0 {
			return isEqual(model, rv.Index(0).Interface())
		}
		return false
	case reflect.Ptr:
		return isEqual(model, rv.Elem().Interface())
	case reflect.Struct:
		if rv.Type().String() == pv.Elem().Type().String() {
			return true
		}
		return false
	}
	return false
}

func (e *Entity) SetValue(sheet *xlsx.Sheet, data interface{}) error {
	rv := reflect.ValueOf(data)

	switch rv.Kind() {
	case reflect.Slice:
		titleRow := sheet.AddRow()
		titleCell := titleRow.AddCell()
		titleCell.SetString(e.Title)
		titleCell.Merge(len(e.Fields), 0)
		for i := 0; i < rv.Len()+1; i++ {
			row := sheet.AddRow()
			for _, v := range e.Fields {
				if i == 0 {
					colCell := row.AddCell()
					colCell.SetString(v.Name)
				} else {
					index := i - 1
					cell := row.AddCell()
					if !rv.Index(index).FieldByName(v.FieldName).IsValid() {
						cell.SetValue(reflect.Zero(v.Typ))
					} else {
						cell.SetValue(rv.Index(index).FieldByName(v.FieldName))
					}
				}

			}
		}
		log.Println(sheet)
	}
	return nil
}

func getField(model interface{}) (rowsMap, fieldsMap map[string]*Field, err error) {
	val, is := orm.RefType(model)
	if !is {
		return nil, nil, errors.New("must be struct")
	}
	rt := reflect.TypeOf(val)
	fieldsMap = make(map[string]*Field)
	rowsMap = make(map[string]*Field)
	for i := 0; i < rt.NumField(); i++ {
		// 可以获取到标签/有值
		tagName := rt.Field(i).Tag.Get("excel")
		if tagName != "" {
			// 字段名称
			filed := new(Field)
			tags := strings.Split(tagName, " ")
			if len(tags) >= 1 {
				filedMap := getFiledMap(tags)
				// 列名
				filed.Name = tags[0]
				// 提示
				if len(tags) > 1 && strings.HasPrefix(tags[1], "tips") {
					filed.Remind = strings.TrimPrefix(tags[1], "tips:'")
					filed.Remind = strings.TrimSuffix(filed.Remind, "'")
				}
				// 设置字段名
				filed.FieldName = rt.Field(i).Name
				// 判断 uqi required
				_, filed.Uqi = filedMap["uqi"]
				filed.UqiMap = make(map[string]struct{})
				_, filed.Required = filedMap["required"]
				// 类型赋值
				filed.Typ = rt.Field(i).Type
				//  字段map
				rowsMap[tags[0]] = filed
				fieldsMap[rt.Field(i).Name] = filed
			}
		}
	}
	return rowsMap, fieldsMap, nil
}

// 获取excel 实体对象
func New(sheetName, title string, tips bool, model interface{}) *Entity {
	e := new(Entity)
	e.New(sheetName, title, tips, model)
	return e
}

func getFiledMap(tags []string) map[string]struct{} {
	filedMap := make(map[string]struct{})
	for _, k := range tags {
		if _, ok := filedMap[k]; !ok {
			filedMap[k] = struct{}{}
		}
	}
	return filedMap
}

func (e *Entity) ReadValue(sheet *xlsx.Sheet) (interface{}, error) {
	/*
			sheet 第一行为标题
				  第二行为列名
		          第三行为提示
		 		  第四行往后为数据
	*/
	// 记录索引位置
	for k, v := range sheet.Rows[1].Cells {
		if _, ok := e.Rows[v.Value]; ok {
			e.Rows[v.Value].Index = k
		}
	}
	// 创建切片值
	sliceType := reflect.SliceOf(reflect.TypeOf(e.Model))
	lens := len(sheet.Rows) - 3
	sliceData := reflect.MakeSlice(sliceType, lens, lens)
	rt := reflect.TypeOf(e.Model).Elem()
	for i := 3; i < len(sheet.Rows); i++ {
		if sheet.Rows[i].Cells == nil {
			continue
		}
		//rv := reflect.ValueOf(e.Model)
		rv := reflect.New(rt)
		for _, fie := range e.Rows {
			cellValue := ""
			if fie.Index < len(sheet.Rows[i].Cells) {
				cellValue = sheet.Rows[i].Cells[fie.Index].Value
			}
			value, isNil := isNil(cellValue)
			location := fmt.Sprintf("当前第%v行的%v", i+1, fie.Name)
			if fie.Required && isNil {
				return nil, errors.New(location + "为必填项")
			}
			if fie.Uqi && !isNil {
				if _, ok := fie.UqiMap[value]; ok {
					return nil, errors.New(location + UqiErr.Error())
				}
				fie.UqiMap[value] = struct{}{}
			}
			switch fie.Typ.Kind() {
			case reflect.String:
				rv.Elem().FieldByName(fie.FieldName).SetString(value)
			case reflect.Int64:
				n, err := parseInt64(value)
				if err != nil {
					return nil, errors.New(location + TypeParseErr.Error() + err.Error())
				}
				rv.Elem().FieldByName(fie.FieldName).SetInt(n)
			case reflect.Int:
				n, err := parseInt(value)
				if err != nil {
					return nil, errors.New(location + err.Error())
				}
				rv.Elem().FieldByName(fie.FieldName).Set(reflect.ValueOf(n))
			case reflect.Bool:
				// todo:布尔值的处理情况
				rv.Elem().FieldByName(fie.FieldName).SetBool(true)
			case reflect.Ptr:
				if !isNil {
					// 判断类型是否为*gtime.Time
					if fie.Typ == reflect.TypeOf(&gtime.Time{}) {
						gt := gtime.NewFromStr(value)
						rv.Elem().FieldByName(fie.FieldName).Set(reflect.ValueOf(gt))
					}
				}
			}
		}
		sliceData.Index(i - 3).Set(rv)
	}

	return sliceData.Interface(), nil
}

// 判断是否为空值
func isNil(value string) (string, bool) {
	if value == "" {
		return value, true
	}
	return value, false
}

func parseInt64(value string) (int64, error) {
	v, no := isNil(value)
	if no {
		return 0, nil
	} else {
		n, err := strconv.ParseInt(v, 10, 64)
		if err != nil {
			return 0, err
		}
		return n, nil
	}
}

func parseInt(value string) (int, error) {
	n, err := parseInt64(value)
	return int(n), err
}
