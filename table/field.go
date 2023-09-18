package table

import (
	"errors"
	"github.com/desdemo/go-common/orm"
	"reflect"
	"strings"
)

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

func getField(model interface{}) (rowsMap, fieldsMap map[string]*Field, err error) {
	val, is := orm.RefType(model)
	if !is {
		return nil, nil, errors.New("must be struct")
	}
	rt := reflect.TypeOf(val)
	fieldsMap = make(map[string]*Field)
	rowsMap = make(map[string]*Field)
	index := 0
	for i := 0; i < rt.NumField(); i++ {
		// 可以获取到标签/有值
		tagName := rt.Field(i).Tag.Get("table")
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
				// 索引位置
				filed.Index = index
				//  字段map
				rowsMap[tags[0]] = filed
				fieldsMap[rt.Field(i).Name] = filed

				index++
			}
		}
	}
	return rowsMap, fieldsMap, nil
}
