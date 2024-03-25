package excel

import "errors"

var (
	TypeParseErr = errors.New("数据转换出错")
	UqiErr       = errors.New("数据存在重复")
	CreateErr    = errors.New("创建表信息失败")
	UnKnowType   = errors.New("未知类型,解析失败")
)
