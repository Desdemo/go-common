package excel

import (
	"errors"
	"fmt"
	"github.com/desdemo/go-common/orm"
	"github.com/gogf/gf/os/gtime"
	"github.com/tealeg/xlsx"
	"log"
	"reflect"
	"strconv"
	"strings"
)

// 导出
type Excel interface {
	New(sheetName, title string, model interface{})
	// 导入
	Import([]byte) (interface{}, error)
	// 导出
	Export(interface{}) ([]byte, error)
}

type Entity struct {
	Model      interface{}       // 模型
	SheetName  string            // 表名
	Title      string            // 标题
	Rows       map[string]*Field // 字段名/ 字段
	ShowRemind bool              //  显示提示
}

type Field struct {
	Name      string // 显示名称
	FieldName string // 字段名称
	// FieldMap  map[string]*Site // Map[显示名称]字段名称
	Value    []interface{}       // 值
	UqiMap   map[string]struct{} // 唯一值
	Remind   string              // 提示
	Uqi      bool                // 唯一
	Required bool                // 必填
	Typ      reflect.Type        // 类型
	Index    int                 // 索引
}

type Site struct {
	FieldName string // 字段名称
	ShowName  string // 显示名称
	Index     int    // 索引
}

var (
	TypeParseErr = errors.New("数据转换出错")
	UqiErr       = errors.New("数据存在重复")
)

func (e *Entity) New(sheetName, title string, tips bool, model interface{}) {
	if sheetName != "" {
		e.SheetName = sheetName
	}
	if title != "" {
		e.Title = title
	}
	e.Model = model
	fields, err := getField(model)
	if err != nil {
		log.Fatal(err)
	}
	e.Rows = fields
	e.ShowRemind = tips
}

func (e *Entity) Import(bytes []byte) (interface{}, error) {
	// Unmarshal

	return nil, nil
}

func (e *Entity) Unmarshal(row *xlsx.Row) error {

	return nil
}

func (e *Entity) Export(i interface{}) ([]byte, error) {
	panic("implement me")
}

func getField(model interface{}) (map[string]*Field, error) {
	val, is := orm.RefType(model)
	if !is {
		return nil, errors.New("must be struct")
	}
	rt := reflect.TypeOf(val)
	filedsMap := make(map[string]*Field)
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
				filedsMap[tags[0]] = filed
			}
		}
	}
	return filedsMap, nil
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
		//rv := reflect.ValueOf(e.Model)
		rv := reflect.New(rt)
		for _, fie := range e.Rows {
			value, isNil := isNil(sheet.Rows[i].Cells[fie.Index].Value)
			location := fmt.Sprintf("当前第%v行的%v", i+1, fie.Name)
			if fie.Required && isNil {
				return nil, errors.New(location + "为必填项")
			}
			if fie.Uqi {
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
