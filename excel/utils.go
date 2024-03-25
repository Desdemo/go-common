package excel

import (
	"reflect"
)

func isEqual(model, data interface{}) bool {
	pv := reflect.ValueOf(model)
	rv := reflect.ValueOf(data)
	switch rv.Kind() {
	case reflect.Slice:
		if pv.Type() == reflect.TypeOf(data).Elem() {
			return true
		}
		if pv.Type().Kind() == reflect.Ptr {
			return pv.Type().Elem() == reflect.TypeOf(data).Elem()
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

func getFiledMap(tags []string) map[string]struct{} {
	filedMap := make(map[string]struct{})
	for _, k := range tags {
		if _, ok := filedMap[k]; !ok {
			filedMap[k] = struct{}{}
		}
	}
	return filedMap
}
