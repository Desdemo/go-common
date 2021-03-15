package orm

import "reflect"

// 类型判断
func RefType(value interface{}) (val interface{}, is bool) {
	if value == nil {
		return nil, false
	}
	rv := reflect.ValueOf(value)
	switch rv.Kind() {
	case reflect.Struct:
		return value, true
	case reflect.Ptr:
		return rv.Elem().Interface(), true
	case reflect.Slice:
		return
	default:
		return nil, false
	}
}
