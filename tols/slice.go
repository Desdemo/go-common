package tols

//// 单元素slice 转 map 去重后返回
//func SliceDed(list interface{}, model interface{}) (interface{}, error) {
//	list, err := ref.IsSlice(list)
//	if err != nil {
//		return nil, err
//	}
//	listMap := make(map[interface{}]struct{})
//	for _, val := range list.()
//}

// 判断切片是否 包含 然后返回
// 字符串列表
func contains(list []string) (bool, error) {
	listMap := make(map[string]struct{})
	for _, val := range list {
		if _, ok := listMap[val]; ok {
			return true, nil
		} else {
			listMap[val] = struct{}{}
		}
	}
	return false, nil
}
