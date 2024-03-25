package excel

//
//import (
//	"bufio"
//	"bytes"
//	"errors"
//	"fmt"
//	"github.com/gogf/gf/v2/os/gtime"
//	xlsx "github.com/tealeg/xlsx/v3"
//	"log"
//	"reflect"
//	"strconv"
//)
//
//type Entity struct {
//	Model      interface{}       // 模型
//	SheetName  string            // 表名
//	Title      string            // 标题
//	Rows       map[string]*Field // 字段表格显示名/ 字段
//	Fields     map[string]*Field // 字段名称 / 字段
//	ShowRemind bool              //  显示提示
//}
//
//type Site struct {
//	FieldName string // 字段名称
//	ShowName  string // 显示名称
//	Index     int    // 索引
//}
//
//var (
//	TypeParseErr = errors.New("数据转换出错")
//	UqiErr       = errors.New("数据存在重复")
//	CreateErr    = errors.New("创建表信息失败")
//	UnKnowType   = errors.New("未知类型,解析失败")
//)
//
//func (e *Entity) New(sheetName, title string, tips bool, model interface{}) {
//	if sheetName != "" {
//		e.SheetName = sheetName
//	}
//	if title != "" {
//		e.Title = title
//	}
//	e.Model = model
//	rows, fields, err := getField(model)
//	if err != nil {
//		log.Fatal(err)
//	}
//	e.Rows = rows
//	e.Fields = fields
//	e.ShowRemind = tips
//}
//
//func (e *Entity) Import(bytes []byte) (interface{}, error) {
//	// Unmarshal
//	file, err := xlsx.OpenBinary(bytes)
//	if err != nil {
//		return nil, err
//	}
//	var sheet *xlsx.Sheet
//	if _, ok := file.Sheet[e.SheetName]; ok {
//		sheet = file.Sheet[e.SheetName]
//	} else {
//		sheet = file.Sheets[0]
//	}
//	data, err := e.ReadValue(sheet)
//	if err != nil {
//		return nil, err
//	}
//	return data, nil
//}
//
//func (e *Entity) Export(i interface{}) ([]byte, error) {
//	if isEqual(e.Model, i) {
//		file := xlsx.NewFile()
//		sheet, err := e.addSheet(file)
//		if err != nil {
//			return nil, err
//		}
//		if sheet == nil {
//			return nil, CreateErr
//		}
//		err = e.SetValue(sheet, i)
//		if err != nil {
//			return nil, err
//		}
//		var b bytes.Buffer
//		writer := bufio.NewWriter(&b)
//		file.Write(writer)
//		return b.Bytes(), nil
//	}
//	return nil, UnKnowType
//}
//
//func (e *Entity) StreamWriter() {}
//
//func (e *Entity) Flush() {}
//
//func (e *Entity) ExportFile(i interface{}, fullPath string) error {
//	if isEqual(e.Model, i) {
//		file := xlsx.NewFile()
//		sheet, err := e.addSheet(file)
//		if err != nil {
//			return err
//		}
//		if sheet == nil {
//			return CreateErr
//		}
//		err = e.SetValue(sheet, i)
//		if err != nil {
//			return err
//		}
//		err = file.Save(fullPath)
//		if err != nil {
//			return err
//		}
//	}
//	return nil
//}
//
//func (e *Entity) addSheet(wb *xlsx.File) (*xlsx.Sheet, error) {
//	sh, err := wb.AddSheet(e.SheetName)
//	if err != nil {
//		return nil, err
//	}
//	return sh, nil
//}
//
//func isEqual(model, data interface{}) bool {
//	pv := reflect.ValueOf(model)
//	rv := reflect.ValueOf(data)
//	switch rv.Kind() {
//	case reflect.Slice:
//		if rv.Len() > 0 {
//			return isEqual(model, rv.Index(0).Interface())
//		}
//		return false
//	case reflect.Ptr:
//		return isEqual(model, rv.Elem().Interface())
//	case reflect.Struct:
//		if rv.Type().String() == pv.Elem().Type().String() {
//			return true
//		}
//		return false
//	}
//	return false
//}
//
//func (e *Entity) SetValue(sheet *xlsx.Sheet, data interface{}) error {
//	rv := reflect.ValueOf(data)
//
//	switch rv.Kind() {
//	case reflect.Ptr:
//		return e.SetValue(sheet, rv.Elem().Interface())
//	case reflect.Slice:
//		titleRow := sheet.AddRow()
//		titleCell := titleRow.AddCell()
//		titleCell.SetString(e.Title)
//		titleCell.Merge(len(e.Fields)-1, 0)
//		colList := make([]string, len(e.Fields))
//		for k, v := range e.Fields {
//			colList[v.Index] = k
//		}
//		for i := 0; i < rv.Len()+1; i++ {
//			row := sheet.AddRow()
//			for _, col := range colList {
//				if i == 0 {
//					colCell := row.AddCell()
//					colCell.SetString(e.Fields[col].Name)
//				} else {
//					index := i - 1
//					cell := row.AddCell()
//					cellValue := reflect.Value{}
//					var cellVal interface{}
//					if rv.Index(index).Kind() == reflect.Ptr {
//						cellValue = reflect.ValueOf(rv.Index(index).Elem().Interface())
//					} else {
//						cellValue = reflect.ValueOf(rv.Index(index).Interface())
//					}
//					if !cellValue.FieldByName(e.Fields[col].FieldName).IsValid() {
//						cellVal = reflect.Zero(e.Fields[col].Typ)
//					} else {
//						if e.Fields[col].Typ.Kind() == reflect.Float64 {
//							cellVal = cellValue.FieldByName(e.Fields[col].FieldName).Float()
//						} else {
//							cellVal = cellValue.FieldByName(e.Fields[col].FieldName).Interface()
//						}
//					}
//					cell.SetValue(cellVal)
//				}
//			}
//		}
//
//	default:
//		return UnKnowType
//	}
//	return nil
//}
//
//// New 获取excel 实体对象
//func New(sheetName, title string, tips bool, model interface{}) *Entity {
//	e := new(Entity)
//	e.New(sheetName, title, tips, model)
//	return e
//}
//

//
//func (e *Entity) ReadValue(sheet *xlsx.Sheet) (interface{}, error) {
//	/*
//			sheet 第一行为标题
//				  第二行为列名
//		          第三行为提示
//		 		  第四行往后为数据
//	*/
//	// 记录索引位置
//	row, _ := sheet.Row(1)
//	err := row.ForEachCell(func(c *xlsx.Cell) error {
//		if _, ok := e.Rows[c.Value]; ok {
//			x, _ := c.GetCoordinates()
//			e.Rows[c.Value].Index = x
//		}
//		return nil
//	})
//	if err != nil {
//		return nil, err
//	}
//	// 创建切片值
//	sliceType := reflect.SliceOf(reflect.TypeOf(e.Model))
//	lens := sheet.MaxRow - 3
//	sliceData := reflect.MakeSlice(sliceType, lens, lens)
//	rt := reflect.TypeOf(e.Model).Elem()
//	for i := 3; i < sheet.MaxRow; i++ {
//		cells, _ := sheet.Row(i)
//		if cells == nil {
//			continue
//		}
//		rv := reflect.New(rt)
//		for _, fie := range e.Rows {
//			cellValue := ""
//			cells, _ := sheet.Row(i)
//			cellValue = cells.GetCell(fie.Index).Value
//			value, isNil := isNil(cellValue)
//			location := fmt.Sprintf("当前第%v行的%v", i+1, fie.Name)
//			if fie.Required && isNil {
//				return nil, errors.New(location + "为必填项")
//			}
//			if fie.Uqi && !isNil {
//				if _, ok := fie.UqiMap[value]; ok {
//					return nil, errors.New(location + UqiErr.Error())
//				}
//				fie.UqiMap[value] = struct{}{}
//			}
//			switch fie.Typ.Kind() {
//			case reflect.String:
//				rv.Elem().FieldByName(fie.FieldName).SetString(value)
//			case reflect.Int64:
//				n, err := parseInt64(value)
//				if err != nil {
//					return nil, errors.New(location + TypeParseErr.Error() + err.Error())
//				}
//				rv.Elem().FieldByName(fie.FieldName).SetInt(n)
//			case reflect.Int:
//				n, err := parseInt(value)
//				if err != nil {
//					return nil, errors.New(location + err.Error())
//				}
//				rv.Elem().FieldByName(fie.FieldName).Set(reflect.ValueOf(n))
//			case reflect.Bool:
//				// todo:布尔值的处理情况
//				rv.Elem().FieldByName(fie.FieldName).SetBool(true)
//			case reflect.Ptr:
//				if !isNil {
//					// 判断类型是否为*gtime.Time
//					if fie.Typ == reflect.TypeOf(&gtime.Time{}) {
//						gt := gtime.NewFromStr(value)
//						rv.Elem().FieldByName(fie.FieldName).Set(reflect.ValueOf(gt))
//					}
//				}
//			}
//		}
//		sliceData.Index(i - 3).Set(rv)
//	}
//
//	return sliceData.Interface(), nil
//}
//
//// 判断是否为空值
//func isNil(value string) (string, bool) {
//	if value == "" {
//		return value, true
//	}
//	return value, false
//}
//
//func parseInt64(value string) (int64, error) {
//	v, no := isNil(value)
//	if no {
//		return 0, nil
//	} else {
//		n, err := strconv.ParseInt(v, 10, 64)
//		if err != nil {
//			return 0, err
//		}
//		return n, nil
//	}
//}
//
//func parseInt(value string) (int, error) {
//	n, err := parseInt64(value)
//	return int(n), err
//}
//
//// Flush 缓存写入
//func Flush() {
//
//}
