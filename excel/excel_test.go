package excel

import (
	"github.com/tealeg/xlsx"
	"reflect"
	"testing"
)

type A struct {
	Id   int    `excel:"样本Id"`
	Code string `excel:"样本编码 tips:'小提示' uqi"`
	Name string `excel:"样本名称 required"`
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
	tests := []struct {
		name    string
		fields  fields
		args    args
		want    interface{}
		wantErr bool
	}{
		// TODO: Add test cases.
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
			if !reflect.DeepEqual(got, tt.want) {
				t.Errorf("ReadValue() got = %v, want %v", got, tt.want)
			}
		})
	}
}
