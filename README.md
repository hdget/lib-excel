# lib-excel

# 1. How to read excel from file

```go
reader, err := NewFileReader("example.xlsx")
if err != nil {
    log.Fatal(err)
}

sheet, err := reader.ReadSheet("Sheet1")
if err != nil {
    return nil, errors.Wrap(err, "读取表格失败")
}

for i, row := range sheet.Rows {
    fmt.Println(row.Get("column1"))
    fmt.Println(row.GetInt64("column2"))
}
```

# 2. How to read excel from network
```go
reader, err := NewHttpReader("http://example.com/download/example.xlsx")
if err != nil {
    log.Fatal(err)
}

sheet, err := reader.ReadSheet("Sheet1")
if err != nil {
    return nil, errors.Wrap(err, "读取表格失败")
}

for i, row := range sheet.Rows {
    fmt.Println(row.Get("column1"))
    fmt.Println(row.GetInt64("column2"))
}
```

# 3. How to write excel file
```go
// define row
type row struct {
    Id int64 `col_name:"序号"`
    Name string `col_name:"名字"`
}

rows := make([]any, 5)
for i:=0; i<5; i++ {
    rows[i] = &row{
		Id: i+1,
        Name: "name" + i,
    }
}

writer, err := NewFileWriter()
if err != nil {
    log.Fatal(err)
}

err = writer.CreateDefaultSheet(rows)
if err != nil {
    log.Fatal(err)
}

err = writer.Save("example.xlsx")
if err != nil {
    log.Fatal(err)
}
```

# 4. How to write excel file
```go
// define row
type Row struct {
    Id int64 `col_name:"序号"`
    Name string `col_name:"名字"`
}

type RowWithAxis struct {
    Id int64 `col_name:"序号",col_axis:"B"`
    Name string `col_name:"名字",col_axis:"A"`
}

rows1 := make([]any, 5)
for i:=0; i<5; i++ {
    rows1[i] = &Row{
		Id: i+1,
        Name: "name" + i,
    }
}

rows2 := make([]any, 5)
for i:=0; i<5; i++ {
    rows2[i] = &RowWithAxis{
        Id: i+1,
        Name: "name" + i,
    }
}

writer, err := NewFileWriter()
if err != nil {
    log.Fatal(err)
}

err = writer.CreateDefaultSheet(rows1)
if err != nil {
    log.Fatal(err)
}

err = writer.Save("example1.xlsx")
if err != nil {
    log.Fatal(err)
}

err = writer.CreateDefaultSheet(rows2)
if err != nil {
    log.Fatal(err)
}

err = writer.Save("example2.xlsx")
if err != nil {
    log.Fatal(err)
}
```
