package xls

type CellDataType string

const (
	CellDataTypeString  CellDataType = "s"
	CellDataTypeNumeric CellDataType = "n"
	CellDataTypeBool    CellDataType = "b"
	CellDataTypeError   CellDataType = "e"
	//CellDataTypeString2 CellDataType = "str"
	//CellDataTypeFormula CellDataType = "f"
	//CellDataTypeNull    CellDataType = "null"
	//CellDataTypeInline  CellDataType = "inlineStr"
)

type Cell struct {
	value    interface{}
	dataType CellDataType
}

func (c *Cell) Value() interface{} {
	return c.value
}

func (c *Cell) DataType() CellDataType {
	return c.dataType
}

func (c *Cell) setValue(value interface{}, dataType CellDataType) {
	c.value = value
	c.dataType = dataType
}
