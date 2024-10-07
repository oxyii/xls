# Very simple XLS file data reader

Partial port of [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) xls reader.

Cannot read styles, margins and formula cells, only values. Only for XLS files, not for XLSX.

## Usage

```go
package main

import (
    "fmt"
	
    "github.com/oxyii/xls"
)

func main() {
    filename := "file.xls"
    xlFile, err := xls.Open(filename)
    if err != nil {
        panic(err)
    }

    sheets := xlFile.Sheets()
	
    for _, sheet := range sheets {
        fmt.Printf("Sheet name: %s\n", sheet.Name())
        fmt.Printf("have %d rows with max %d columns\n", sheet.Rows(), sheet.Cols())
        for i := 0; i < sheet.Rows(); i++ {
            row := sheet.Row(i)
            // get Cols exactly from row! Otherwise you can get nil pointer error in merged cells
            for j := 0; j < row.Cols(); j++ {
                cell := row.Cell(j)
                fmt.Printf("cell[%d][%d] = %v\n", i, j, cell.Value())
            }
        }
    }
}
```