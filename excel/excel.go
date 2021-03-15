package excel

import (
	"errors"
	"git.dashoo.cn/guodj/common/orm"
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

type A struct {
	Id   int    `excel:"样本Id"`
	Code string `excel:"样本编码 uqi"`
	Name string `excel:"样本名称 required"`
}

type Entity struct {
	Model     interface{}
	SheetName string
	Title     string  // 标题
	Rows      []Field // 字段
	ShowRemind bool  // 显示提示
}

type Field struct {
	Name     string  // 显示名称
	Value    []interface{} // 值
	Remind   string // 提示
	Uqi      bool // 唯一
	Required bool // 必填
	Typ      reflect.Type // 类型
}

func (e *Entity) New(sheetName, title string, model interface{}) {
	if sheetName != "" {
		e.SheetName = sheetName
	}
	if title != "" {
		e.Title = title
	}
	fields, err := getField(model)
	if err != nil {
		log.Fatal(err)
	}
	e.Rows = fields

}

func (e *Entity) Import(bytes []byte) (interface{}, error) {
	panic("implement me")
}

func (e *Entity) Export(i interface{}) ([]byte, error) {
	panic("implement me")
}

func getField(model interface{}) ([]Field, error) {
	val, is := orm.RefType(model)
	if !is {
		return nil, errors.New("must be struct")
	}
	fileds := make([]Field, 0)
	rt := reflect.TypeOf(val)
	for i := 0; i < rt.NumField(); i++ {
		// 可以获取到标签/有值
		tagName := rt.Field(i).Tag.Get("filter")
		if tagName != "" {
			// 字段名称
			filed := new(Field)
			tags := strings.Split(rt.Field(i).Name, " ")
			if len(tags) >= 1 {
				filedMap := getFiledMap(tags)
				// 列名
				filed.Name = tags[0]
				// 提示
				filed.Remind = tags[1]
				// 判断 uqi required
				_, filed.Required = filedMap["uqi"]
				_, filed.Required = filedMap["required"]
				filed.Typ = rt
			}
			fileds[i] = *filed
		}
	}
	return fileds, nil
}

// 获取excel 实体对象
func New(entity interface{}) {

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
