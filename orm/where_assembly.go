package orm

import (
	"errors"
	"github.com/gogf/gf/v2/os/gtime"
	"reflect"

	"github.com/gogf/gf/v2/database/gdb"
)

func OrmPipe(model *gdb.Model, value interface{}) (*gdb.Model, error) {
	val, is := RefType(value)
	if !is {
		return nil, errors.New("must be struct")
	}
	rv := reflect.ValueOf(val)
	rt := reflect.TypeOf(val)
	for i := 0; i < rt.NumField(); i++ {
		//  获取字段类型, 如果是结构体 进去调用
		if rt.Field(i).Type.Kind() == reflect.Struct {
			smodel, err := OrmPipe(model, rv.Field(i).Interface())
			if err != nil {
				return nil, err
			}
			model = smodel
		}
		// 可以获取到标签/有值
		tagName := rt.Field(i).Tag.Get("filter")
		if tagName != "" {
			// 字段名称
			typeName := rt.Field(i).Name
			// like sql
			if tagName == "like" {
				if refValue := rv.Field(i).Interface().(string); refValue != "" {
					model = WhereLike(model, typeName, refValue)
				}
			}
			// equal sql
			if tagName == "equal" {
				// string 非空
				if rv.Field(i).Type().Kind() == reflect.String {
					if refValue := rv.Field(i).Interface().(string); refValue != "" {
						model = WhereEqualForString(model, typeName, refValue)
					}
				}
				// int64 非0
				if rv.Field(i).Type().Kind() == reflect.Int64 || rv.Field(i).Type().Kind() == reflect.Int {
					if refValue, ok := rv.Field(i).Interface().(int64); ok && refValue != 0 {
						model = WhereEqualForInt(model, typeName, refValue)
					} else {
						if refValue, ok := rv.Field(i).Interface().(int); ok && refValue != 0 {
							model = WhereEqualForInt(model, typeName, int64(refValue))
						}
					}
				}
			}
			// between ... and ...
			if tagName == "between" {
				// 切片长度为2,类型为 子符串或者时间
				if rv.Field(i).Type().Kind() == reflect.Slice && rv.Field(i).Type().String() == "[]string" {
					if refValue := rv.Field(i).Interface().([]string); len(refValue) == 2 {
						model = WhereBetweenForString(model, typeName, refValue)
					}
				}
				if rv.Field(i).Type().Kind() == reflect.Slice && rv.Field(i).Type().String() == "[]gtime.Time" {
					if refValue := rv.Field(i).Interface().([]gtime.Time); len(refValue) == 2 {
						model = WhereBetweenForTime(model, typeName, refValue)
					}
				}
			}
		}
	}
	return model, nil
}

// like string
func WhereLike(model *gdb.Model, typeName, refValue string) *gdb.Model {
	return model.Where(typeName+" like ?", "%"+refValue+"%")
}

// equal string
func WhereEqualForString(model *gdb.Model, typeName, refValue string) *gdb.Model {
	return model.Where(typeName+" = ?", "'"+refValue+"'")
}

// equal int64
func WhereEqualForInt(model *gdb.Model, typeName string, refValue int64) *gdb.Model {
	return model.Where((typeName)+" = ?", refValue)
}

// between ... and ... 时间类型
func WhereBetweenForString(model *gdb.Model, typeName string, refValue []string) *gdb.Model {
	if len(refValue) == 2 {
		return model.Where(typeName+" between ? and ?", refValue[0], refValue[1])
	}
	return model
}

func WhereBetweenForTime(model *gdb.Model, typeName string, refValue []gtime.Time) *gdb.Model {
	if len(refValue) == 2 {
		return model.Where(typeName+" between ? and ?", refValue[0].String(), refValue[1].String())
	}
	return model
}
