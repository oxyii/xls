package xls

type Row struct {
	cells map[int]*Cell
}

func (r *Row) Cell(index int) *Cell {
	if c, ok := r.cells[index]; ok {
		return c
	}
	return nil
}

func (r *Row) Cols() int {
	return len(r.cells)
}

func (r *Row) getCell(col int) *Cell {
	if c, ok := r.cells[col]; ok {
		return c
	}
	c := new(Cell)
	r.cells[col] = c
	return c
}
