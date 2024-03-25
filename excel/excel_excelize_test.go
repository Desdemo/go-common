package excel

import (
	"github.com/gogf/gf/v2/os/gtime"
	"math/rand"
	"os"
	"reflect"
	"strconv"
	"testing"
	"time"
)

func TestExcellingEntity_SetValue(t *testing.T) {
	e := Newlize("xx", "test", false, new(A))
	data01 := []A{
		{Id: 12345678, Code: "2021031001", Name: "box031002",
			StartTime: gtime.NewFromStr("2021-03-19 11:56:56")},
		{Id: 12345678, Code: "2021031003", Name: "box031004",
			StartTime: gtime.NewFromStr("2021-03-19 11:56:56")}}
	type args struct {
		data interface{}
	}
	tests := []struct {
		name    string
		args    args
		wantErr bool
	}{
		{"test", args{data: data01}, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if err := e.SetValue(tt.args.data); (err != nil) != tt.wantErr {
				t.Errorf("SetValue() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

func TestExcellingEntity_Export(t *testing.T) {
	e := Newlize("xx", "test", false, new(A))
	data01 := []A{
		{Id: 12345678, Code: "2021031001", Name: "box031002",
			StartTime: gtime.NewFromStr("2021-03-19 11:56:56")},
		{Id: 12345678, Code: "2021031003", Name: "box031004",
			StartTime: gtime.NewFromStr("2021-03-19 11:56:56")}}

	type args struct {
		i interface{}
	}
	tests := []struct {
		name    string
		args    args
		wantErr bool
	}{
		{"test", args{i: data01}, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := e.Export(tt.args.i)
			if (err != nil) != tt.wantErr {
				t.Errorf("Export() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			t.Logf("%+v", got)
		})
	}
}

func TestExcellingEntity_ExportFile(t *testing.T) {
	e := Newlize("Sheet1", "列表数据", false, new(A))
	data01 := []A{
		{Id: 12345678, Code: "2021031001", Name: "box031002",
			StartTime: gtime.NewFromStr("2021-03-19 11:56:56")},
		{Id: 12345678, Code: "2021031001", Name: "box031002"},
		{Id: 12345678, Code: "2021031003", Name: "box031004",
			StartTime: gtime.NewFromStr("2021-03-19 11:56:56")}}

	for i := 0; i < 1000000; i++ {
		data01 = append(data01, A{Id: int(rand.Int63n(100000000)),
			Name: strconv.Itoa(int(rand.Int63n(100000000)))})
	}
	type args struct {
		i interface{}
	}
	tests := []struct {
		name    string
		args    args
		wantErr bool
	}{
		{"test", args{i: data01}, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if err := e.ExportFile(tt.args.i); (err != nil) != tt.wantErr {
				t.Errorf("ExportFile() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

func TestExcellingEntity_Import(t *testing.T) {
	e := Newlize("Sheet1", "列表数据", true, new(As))
	//testDir := "./" // 当前目录
	//filePath := filepath.Join(testDir,"data.txt")

	dataByte, err := os.ReadFile("1.xlsx")
	if err != nil {
		t.Error(err)
	}

	type args struct {
		data []byte
	}
	tests := []struct {
		name    string
		args    args
		wantErr bool
	}{
		{"test", args{data: dataByte}, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := e.Import(tt.args.data)
			if (err != nil) != tt.wantErr {
				t.Errorf("Import() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			x := got.([]*As)
			for _, v := range x {
				t.Logf("%+v\n", v)
			}

		})
	}
}

func Test_convertStringToType(t *testing.T) {

	timeWant, _ := time.Parse(time.DateTime, "2016-12-23 12:12:00")

	type args struct {
		val string
		typ reflect.Type
	}
	tests := []struct {
		name string
		args args
		want interface{}
	}{
		{"str", args{val: "A", typ: reflect.TypeOf("A")}, "A"},
		{"time", args{
			val: "2016-12-23 12:12:00", typ: reflect.TypeOf(timeWant)},
			timeWant},
		{"default_time", args{
			val: "", typ: reflect.TypeOf(timeWant)},
			time.Time{}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := convertStringToType(tt.args.val, tt.args.typ)
			t.Log(got)
			if !reflect.DeepEqual(got, tt.want) {
				t.Errorf("convertStringToType() = %v, want %v", got, tt.want)
			}
		})
	}
}
