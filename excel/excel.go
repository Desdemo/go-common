package excel

import (
	"errors"
	"github.com/desdemo/go-common/orm"
	"github.com/tealeg/xlsx"
	"log"
	"reflect"
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
	Value    []interface{} // 值
	Remind   string        // 提示
	Uqi      bool          // 唯一
	Required bool          // 必填
	Typ      reflect.Type  // 类型
	Index    int           // 索引
}

type Site struct {
	FieldName string // 字段名称
	ShowName  string // 显示名称
	Index     int    // 索引
}

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
	for i := 3; i < lens; i++ {
		//rv := reflect.ValueOf(e.Model)
		rv := reflect.New(rt)
		for _, fie := range e.Rows {
			switch fie.Typ.Kind() {
			case reflect.String:
				rv.Elem().FieldByName(fie.FieldName).SetString(sheet.Rows[i].Cells[fie.Index].Value)
			case reflect.Int64:
				value, err := sheet.Rows[i].Cells[fie.Index].Int64()
				if err != nil {
					return nil, err
				}
				rv.Elem().FieldByName(fie.FieldName).SetInt(value)
			case reflect.Int:
				value, err := sheet.Rows[i].Cells[fie.Index].Int()
				if err != nil {
					return nil, err
				}
				rv.Elem().FieldByName(fie.FieldName).Set(reflect.ValueOf(value))
			case reflect.Bool:
				// todo:布尔值的处理情况, 唯一性校验，必填校验
			}
		}
		log.Println(rv)
		sliceData.Index(i - 3).Set(rv)
	}

	return sliceData.Interface(), nil
}
