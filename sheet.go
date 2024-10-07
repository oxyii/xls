package xls

const (
	XLS_SHEET_STATE_VISIBLE    = 0x00
	XLS_SHEET_STATE_HIDDEN     = 0x01
	XLS_SHEET_STATE_VERYHIDDEN = 0x02
)

type Sheet struct {
	name       string
	offset     int
	sheetState int
	sheetType  byte

	rows map[int]*Row

	maxRow int
	maxCol int
}

func (s *Sheet) Name() string {
	return s.name
}

/*
	func (s *Sheet) Offset() int {
		return s.offset
	}
*/

func (s *Sheet) SheetState() int {
	return s.sheetState
}

func (s *Sheet) SheetType() byte {
	return s.sheetType
}

func (s *Sheet) Row(index int) *Row {
	if index < 0 || index > s.maxRow {
		return nil
	}
	if r, ok := s.rows[index]; ok {
		return r
	}
	return s.getRow(index)
}

func (s *Sheet) Rows() int {
	return s.maxRow + 1
}

func (s *Sheet) Cols() int {
	return s.maxCol + 1
}

func (s *Sheet) setValue(row, col int, value interface{}, dataType CellDataType) {
	s.maxRow = max(s.maxRow, row)
	s.maxCol = max(s.maxCol, col)

	r := s.getRow(row)
	c := r.getCell(col)

	c.setValue(value, dataType)
}

func (s *Sheet) getRow(row int) *Row {
	if r, ok := s.rows[row]; ok {
		return r
	}
	r := new(Row)
	r.cells = make(map[int]*Cell)
	for i := 0; i <= s.maxCol; i++ {
		r.cells[i] = new(Cell)
	}
	s.rows[row] = r
	return r
}
