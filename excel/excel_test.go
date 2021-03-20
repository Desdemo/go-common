package excel

import (
	"github.com/gogf/gf/os/gtime"
	"github.com/tealeg/xlsx"
	"reflect"
	"testing"
)

type A struct {
	Id        int         `excel:"样本Id"`
	Code      string      `excel:"样本编码 tips:'小提示' uqi required"`
	Name      string      `excel:"样本名称 required"`
	StartTime *gtime.Time `excel:"样本时间"`
}

func Test_getField(t *testing.T) {
	type args struct {
		model interface{}
	}
	filesMap := make(map[string]*Field)
	filesMap["样本Id"] = &Field{
		Name: "样本Id", Value: nil, Remind: "", Uqi: false, FieldName: "Id",
		Required: false, Typ: reflect.ValueOf(int(0)).Type()}
	filesMap["样本编码"] = &Field{
		Name: "样本编码", Value: nil, Remind: "小提示", Uqi: true, FieldName: "Code",
		Required: false, Typ: reflect.ValueOf("").Type()}
	filesMap["样本名称"] = &Field{
		Name: "样本名称", Value: nil, Remind: "", Uqi: false, FieldName: "Name",
		Required: true, Typ: reflect.ValueOf("").Type()}
	tests := []struct {
		name    string
		args    args
		want    map[string]*Field
		wantErr bool
	}{
		// TODO: Add test cases.
		{"excel标签测试", struct{ model interface{} }{model: new(A)}, filesMap, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := getField(tt.args.model)
			if (err != nil) != tt.wantErr {
				t.Errorf("getField() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if !reflect.DeepEqual(len(got), len(tt.want)) {
				t.Errorf("get length got = %v, want %v", len(got), len(tt.want))
			}
			for _, i := range got {
				if !reflect.DeepEqual(i, tt.want[i.Name]) {
					t.Errorf("getField detail got = %v, want %v", i, tt.want[i.Name])
				}
			}

		})
	}
}

func TestEntity_ReadValue(t *testing.T) {
	type fields struct {
		Model      interface{}
		SheetName  string
		Title      string
		Rows       map[string]*Field
		ShowRemind bool
	}
	type args struct {
		sheet *xlsx.Sheet
	}
	// 新建测试文件
	rSlice := make([]A, 0)
	testCols := []string{"样本Id", "样本编码", "样本名称", "样本时间"}
	testA := A{Id: 1, Code: "2021031001", Name: "box031002", StartTime: gtime.NewFromStr("2021-03-19 11:56:56")}
	testB := A{Id: 12, Code: "002", Name: "东东", StartTime: gtime.NewFromStr("2021-03-19 11:56:56")}
	testUqi := A{Id: 2, Code: "2021031001", Name: "东东", StartTime: gtime.NewFromStr("2021-03-19 11:56:56")}
	testRequired1 := A{Id: 13, Code: "003", StartTime: gtime.NewFromStr("2021-03-19 11:56:56")}
	testRequired2 := A{Id: 1, Code: "001", Name: "东东"}
	rSlice = append(rSlice, testA, testB)
	xlsxMap := make(map[string]*xlsx.File)
	fileName := []string{"xFile", "xFileUqi", "xFileRequired1", "xFileRequired2"}
	for _, f := range fileName {
		xlsxMap[f] = xlsx.NewFile()
		sheetObj, _ := xlsxMap[f].AddSheet("Test1")
		sheetObj.AddRow()
		cols := sheetObj.AddRow()
		cols.WriteSlice(&testCols, -1)
		sheetObj.AddRow()
		for _, k := range rSlice {
			row := sheetObj.AddRow()
			row.WriteStruct(&k, -1)
		}
		if f == "xFileUqi" {
			row := sheetObj.AddRow()
			row.WriteStruct(&testUqi, -1)
		} else if f == "xFileRequired1" {
			row := sheetObj.AddRow()
			row.WriteStruct(&testRequired1, -1)
		} else if f == "xFileRequired2" {
			row := sheetObj.AddRow()
			row.WriteStruct(&testRequired2, -1)
		}
	}
	file := New("Sheet2", "", false, new(A))
	fileData := New("Test1", "", false, new(A))
	wb, err := xlsx.OpenFile("./1.xlsx")
	if err != nil {
		panic(err)
	}
	sheet := wb.Sheets[0]
	wantData := make([]*A, 0)
	a := &A{
		Id:        1,
		Code:      "2021031001",
		Name:      "box031002",
		StartTime: gtime.NewFromStr("2021-03-19 11:56:56"),
	}
	wantData = append(wantData, a)
	tests := []struct {
		name    string
		fields  fields
		args    args
		want    interface{}
		wantErr bool
	}{
		{"表格文件读取数据测试", fields{Model: file.Model, SheetName: file.SheetName,
			Title: file.Title, Rows: file.Rows, ShowRemind: file.ShowRemind},
			args{sheet: sheet}, wantData, false},
		{"表格数据读取测试", fields{Model: fileData.Model, SheetName: fileData.SheetName,
			Title: fileData.Title, Rows: fileData.Rows, ShowRemind: fileData.ShowRemind},
			args{sheet: xlsxMap["xFile"].Sheet["Test1"]}, wantData, false},
		{"唯一性标签数据测试", fields{Model: file.Model, SheetName: file.SheetName,
			Title: file.Title, Rows: file.Rows, ShowRemind: file.ShowRemind},
			args{sheet: xlsxMap["xFileUqi"].Sheet["Test1"]}, nil, true},
		{"必填标签数据测试1", fields{Model: file.Model, SheetName: file.SheetName,
			Title: file.Title, Rows: file.Rows, ShowRemind: file.ShowRemind},
			args{sheet: xlsxMap["xFileRequired1"].Sheet["Test1"]}, nil, true},
		{"必填标签数据测试2", fields{Model: file.Model, SheetName: file.SheetName,
			Title: file.Title, Rows: file.Rows, ShowRemind: file.ShowRemind},
			args{sheet: xlsxMap["xFileRequired2"].Sheet["Test1"]}, nil, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			e := &Entity{
				Model:      tt.fields.Model,
				SheetName:  tt.fields.SheetName,
				Title:      tt.fields.Title,
				Rows:       tt.fields.Rows,
				ShowRemind: tt.fields.ShowRemind,
			}
			got, err := e.ReadValue(tt.args.sheet)
			if (err != nil) != tt.wantErr {
				t.Errorf("ReadValue() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if got != nil && !reflect.DeepEqual(got.([]*A)[0], tt.want.([]*A)[0]) {
				t.Errorf("ReadValue() got = %v, want %v", got, tt.want)
			}
		})
	}
}
