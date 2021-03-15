package ref

import (
	"errors"
	"reflect"
)

// 判断类型是否是切片
func IsSlice(list interface{}) (interface{}, error) {
	if list == nil {
		return nil, errors.New("input value is nil")
	}
	tv := reflect.ValueOf(list)
	switch tv.Kind() {
	case reflect.Slice:
		return list, nil
	case reflect.Ptr:
		return IsSlice(tv.Elem().Interface())
	default:
		return nil, errors.New("type is slice or *slice")
	}
}
