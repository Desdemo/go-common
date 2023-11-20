package table

// Excel 导出
type Excel interface {
	New(sheetName, title string, tips bool, model interface{})
	//Import 导入
	Import([]byte) (interface{}, error)
	//Export 导出
	Export(interface{}) ([]byte, error)
}
