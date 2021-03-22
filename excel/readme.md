# excel 导入导出

## 使用

```go
go get -u github.com/desdemo/go-common
```



## 标签

```go
type A struct {
	Id        int         `excel:"样本Id"`
	Code      string      `excel:"样本编码 tips:'小提示' uqi required"`
	Name      string      `excel:"样本名称 required"`
	StartTime *gtime.Time `excel:"样本时间"`
}
// 标签中第一个为表格名称 
// 标签中第二个为表格提示
// uqi  为唯一,会进行表格内唯一性校验
// required 为必填,读取或写入时，如果为空，则报错 
```

## 表格说明

![image-20210322165959614](https://gitee.com/desdemo/blog-image/raw/master/img/image-20210322165959614.png)

## 说明：

读取或者写入时,默认第一行为标题,第二行为字段名,第三行为提示信息，第4行开始为数据

目前支持的样式为： string, int, int64, *gtime.Time

### 导入读取

```go
package main

import (
	"github.com/desdemo/go-common/excel"
	"github.com/gin-gonic/gin"
	"github.com/gogf/gf/os/gtime"
	"io/ioutil"
	"log"
	"net/http"
)

type A struct {
	Id        int         `excel:"样本Id"`
	Code      string      `excel:"样本编码 tips:'小提示' uqi required"`
	Name      string      `excel:"样本名称 required"`
	StartTime *gtime.Time `excel:"样本时间"`
}

func main() {
	r := gin.Default()
	r.POST("/", HandleUploadFile)
	r.Run()
}

// HandleUploadFile 上传单个文件
func HandleUploadFile(c *gin.Context) {
	file, _, err := c.Request.FormFile("file")
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"msg": "文件上传失败"})
		return
	}
	content, err := ioutil.ReadAll(file)
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"msg": "文件读取失败"})
		return
	}
    // 初始化对象并读取
	entity, err := excel.New("Test1", "", false, new(A)).Import(content)
	if err != nil {
		log.Println(err)
	}
	for _, k := range entity.([]*A) {
		log.Println(*k)
	}
	c.JSON(http.StatusOK, gin.H{"msg": "上传成功"})
}

```

