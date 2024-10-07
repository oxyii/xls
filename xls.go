package xls

import (
	"golang.org/x/text/encoding"
	"unicode/utf16"
)

const (
	XLS_BIFF8 = 0x0600
	XLS_BIFF7 = 0x0500

	XLS_WORKBOOKGLOBALS = 0x0005
	XLS_WORKSHEET       = 0x0010
)

type stringConvertion struct {
	value string
	size  int
}

type sstConvertion struct {
	recordData    []byte
	spliceOffsets []int
}

type XLS struct {
	ole  *OLE
	data []byte

	CodePage encoding.Encoding

	dataSize int
	pos      int

	version int

	sheets []*Sheet

	sst []string
}

func Open(filename string) (*XLS, error) {
	ole, err := readOLE(filename)
	if err != nil {
		return nil, err
	}

	xls := &XLS{ole: ole, data: ole.getStream(ole.wrkbook), CodePage: DefaultCodePage}

	xls.setDocumentSummaryInformation(ole.getStream(ole.documentSummaryInformation))

	xls.dataSize = len(xls.data)
	xls.pos = 0

	xls.sst = []string{}

external1:
	for xls.pos < xls.dataSize {
		code := getUInt2d(xls.data, xls.pos)
		switch code {
		case XLS_TYPE_BOF:
			xls.readBof() // <- implemented
			break
		case XLS_TYPE_FILEPASS:
			xls.readDefault()
			break
		case XLS_TYPE_CODEPAGE:
			//xls.CodePage = parseCodePage(getUInt2d(xls.data, xls.pos+4))
			xls.readDefault()
			break
		case XLS_TYPE_DATEMODE:
			xls.readDefault()
			break
		case XLS_TYPE_FONT:
			xls.readDefault()
			break
		case XLS_TYPE_FORMAT:
			xls.readDefault()
			break
		case XLS_TYPE_XF:
			xls.readDefault()
			break
		case XLS_TYPE_XFEXT:
			xls.readDefault()
			break
		case XLS_TYPE_STYLE:
			xls.readDefault()
			break
		case XLS_TYPE_PALETTE:
			xls.readDefault()
			break
		case XLS_TYPE_SHEET:
			xls.readSheet() // <- implemented
			break
		case XLS_TYPE_EXTERNALBOOK:
			xls.readDefault()
			break
		case XLS_TYPE_EXTERNNAME:
			xls.readDefault()
			break
		case XLS_TYPE_EXTERNSHEET:
			xls.readDefault()
			break
		case XLS_TYPE_DEFINEDNAME:
			xls.readDefault()
			break
		case XLS_TYPE_MSODRAWINGGROUP:
			xls.readDefault()
			break
		case XLS_TYPE_SST:
			xls.readSst() // <- implemented
			break
		case XLS_TYPE_EOF:
			xls.readDefault()
			break external1
		default:
			xls.readDefault()
		}
	}

	for _, sheet := range xls.sheets {
		if sheet.sheetType != 0x00 {
			// 0x00: Worksheet, 0x02: Chart, 0x06: Visual Basic module
			continue
		}

		xls.pos = sheet.offset

	external2:
		for xls.pos < xls.dataSize-4 {
			code := getUInt2d(xls.data, xls.pos)
			switch code {
			case XLS_TYPE_BOF:
				xls.readDefault()
				break
			case XLS_TYPE_PRINTGRIDLINES:
				xls.readDefault()
				break
			case XLS_TYPE_DEFAULTROWHEIGHT:
				xls.readDefault()
				break
			case XLS_TYPE_SHEETPR:
				xls.readDefault()
				break
			case XLS_TYPE_HORIZONTALPAGEBREAKS:
				xls.readDefault()
				break
			case XLS_TYPE_VERTICALPAGEBREAKS:
				xls.readDefault()
				break
			case XLS_TYPE_HEADER:
				xls.readDefault()
				break
			case XLS_TYPE_FOOTER:
				xls.readDefault()
				break
			case XLS_TYPE_HCENTER:
				xls.readDefault()
				break
			case XLS_TYPE_VCENTER:
				xls.readDefault()
				break
			case XLS_TYPE_LEFTMARGIN:
				xls.readDefault()
				break
			case XLS_TYPE_RIGHTMARGIN:
				xls.readDefault()
				break
			case XLS_TYPE_TOPMARGIN:
				xls.readDefault()
				break
			case XLS_TYPE_BOTTOMMARGIN:
				xls.readDefault()
				break
			case XLS_TYPE_PAGESETUP:
				xls.readDefault()
				break
			case XLS_TYPE_PROTECT:
				xls.readDefault()
				break
			case XLS_TYPE_SCENPROTECT:
				xls.readDefault()
				break
			case XLS_TYPE_OBJECTPROTECT:
				xls.readDefault()
				break
			case XLS_TYPE_PASSWORD:
				xls.readDefault()
				break
			case XLS_TYPE_DEFCOLWIDTH:
				xls.readDefault()
				break
			case XLS_TYPE_COLINFO:
				xls.readDefault()
				break
			case XLS_TYPE_DIMENSION:
				xls.readDefault()
				break
			case XLS_TYPE_ROW:
				xls.readDefault()
				break
			case XLS_TYPE_DBCELL:
				xls.readDefault()
				break
			case XLS_TYPE_RK:
				xls.readRK(sheet) // <- implemented
				break
			case XLS_TYPE_LABELSST:
				xls.readLabelSst(sheet) // <- implemented
				break
			case XLS_TYPE_MULRK:
				xls.readDefault()
				break
			case XLS_TYPE_NUMBER:
				xls.readNumber(sheet) // <- implemented
				break
			case XLS_TYPE_FORMULA:
				xls.readDefault()
				break
			case XLS_TYPE_SHAREDFMLA:
				xls.readDefault()
				break
			case XLS_TYPE_BOOLERR:
				xls.readBoolErr(sheet) // <- implemented
				break
			case XLS_TYPE_MULBLANK:
				xls.readDefault()
				break
			case XLS_TYPE_LABEL:
				xls.readLabel(sheet) // <- implemented
				break
			case XLS_TYPE_BLANK:
				xls.readDefault()
				break
			case XLS_TYPE_MSODRAWING:
				xls.readDefault()
				break
			case XLS_TYPE_OBJ:
				xls.readDefault()
				break
			case XLS_TYPE_WINDOW2:
				xls.readDefault()
				break
			case XLS_TYPE_PAGELAYOUTVIEW:
				xls.readDefault()
				break
			case XLS_TYPE_SCL:
				xls.readDefault()
				break
			case XLS_TYPE_PANE:
				xls.readDefault()
				break
			case XLS_TYPE_SELECTION:
				xls.readDefault()
				break
			case XLS_TYPE_MERGEDCELLS:
				xls.readDefault()
				break
			case XLS_TYPE_HYPERLINK:
				xls.readDefault()
				break
			case XLS_TYPE_DATAVALIDATIONS:
				xls.readDefault()
				break
			case XLS_TYPE_DATAVALIDATION:
				xls.readDefault()
				break
			case XLS_TYPE_SHEETLAYOUT:
				xls.readDefault()
				break
			case XLS_TYPE_SHEETPROTECTION:
				xls.readDefault()
				break
			case XLS_TYPE_RANGEPROTECTION:
				xls.readDefault()
				break
			case XLS_TYPE_NOTE:
				xls.readDefault()
				break
			case XLS_TYPE_TXO:
				xls.readDefault()
				break
			case XLS_TYPE_CONTINUE:
				xls.readDefault()
				break
			case XLS_TYPE_EOF:
				xls.readDefault()
				break external2
			default:
				xls.readDefault()
			}
		}
	}

	return xls, nil
}

func (xls *XLS) Sheets() []*Sheet {
	return xls.sheets
}

func (xls *XLS) setDocumentSummaryInformation(data []byte) {
	secOffset := getInt4d(data, 44)
	countProperties := getInt4d(data, secOffset+4)

	for i := 0; i < countProperties; i++ {
		id := getInt4d(data, (secOffset+8)+(8*i))
		offset := getInt4d(data, (secOffset+12)+(8*i))
		typeID := getInt4d(data, secOffset+offset)

		var value interface{}

		switch typeID {
		case 0x02: // 2 byte signed integer
			value = getUInt2d(data, secOffset+4+offset)
			break
		case 0x03: // 4 byte signed integer
			value = getInt4d(data, secOffset+4+offset)
			break
		case 0x0B: // Boolean
			value = getUInt2d(data, secOffset+4+offset) != 0
			break
		case 0x13: // 4 byte unsigned integer
			break
		case 0x1E: // null-terminated string prepended by dword string length
			//byteLength := getInt4d(data, secOffset+4+offset)
			//value = string(data[secOffset+8+offset : secOffset+8+offset+byteLength])
			//value = strings.TrimRight(value.(string), "\x00")
			break
		case 0x40: // Filetime (64-bit value representing the number of 100-nanosecond intervals since January 1, 1601)
			//value = ole2LocalDate(data[secOffset+4+offset : secOffset+4+offset+8])
			break
		}

		switch id {
		case 0x01:
			xls.CodePage = parseCodePage(value.(uint16))
		case 0x02:
			//ole.spreadsheet.Properties.Category = value.(string)
			break
		case 0x0E:
			//ole.spreadsheet.Properties.Manager = value.(string)
			break
		case 0x0F:
			//ole.spreadsheet.Properties.Company = value.(string)
			break
		}
	}
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

func (xls *XLS) readDefault() {
	length := getUInt2d(xls.data, xls.pos+2)
	xls.pos += 4 + int(length)
}

func (xls *XLS) readBof() {
	length := getUInt2d(xls.data, xls.pos+2)
	recordData := xls.data[xls.pos+4 : xls.pos+4+int(length)]

	// move stream pointer to next record
	xls.pos += 4 + int(length)

	// offset: 2; size: 2; type of the following data
	substreamType := getUInt2d(recordData, 2)

	switch substreamType {
	case XLS_WORKBOOKGLOBALS:
		version := getUInt2d(recordData, 0)
		if version != XLS_BIFF8 && version != XLS_BIFF7 {
			//return errors.New("cannot read this Excel file. Version is too old")
			return
		}
		xls.version = int(version)

	case XLS_WORKSHEET:
		// do not use this version information for anything
		// it is unreliable (OpenOffice doc, 5.8), use only version information from the global stream

	default:
		// substream, e.g. chart
		// just skip the entire substream
		for {
			code := getUInt2d(xls.data, xls.pos)
			xls.readDefault()
			if code == XLS_TYPE_EOF || xls.pos >= xls.dataSize {
				break
			}
		}
	}
}

func (xls *XLS) readSst() {
	// offset within (spliced) record data
	pos := 0

	// get spliced record data
	splicedRecordData := xls.getSplicedRecordData()

	recordData := splicedRecordData.recordData
	spliceOffsets := splicedRecordData.spliceOffsets

	// offset: 0; size: 4; total number of strings in the workbook
	pos += 4

	// offset: 4; size: 4; number of following strings ($nm)
	nm := getInt4d(recordData, 4)
	pos += 4

	// loop through the Unicode strings (16-bit length)
	for i := 0; i < nm; i++ {
		// number of characters in the Unicode string
		numChars := getUInt2d(recordData, pos)
		pos += 2

		// option flags
		optionFlags := recordData[pos]
		pos++

		// bit: 0; mask: 0x01; 0 = compressed; 1 = uncompressed
		isCompressed := (optionFlags & 0x01) == 0

		// bit: 2; mask: 0x02; 0 = ordinary; 1 = Asian phonetic
		hasAsian := (optionFlags & 0x04) != 0

		// bit: 3; mask: 0x03; 0 = ordinary; 1 = Rich-Text
		hasRichText := (optionFlags & 0x08) != 0

		var formattingRuns uint16
		if hasRichText {
			// number of Rich-Text formatting runs
			formattingRuns = getUInt2d(recordData, pos)
			pos += 2
		}

		var extendedRunLength int
		if hasAsian {
			// size of Asian phonetic setting
			extendedRunLength = getInt4d(recordData, pos)
			pos += 4
		}

		// expected byte length of character array if not split
		length := int(numChars)
		if !isCompressed {
			length *= 2
		}

		// look up limit position
		var limitpos int
		for _, spliceOffset := range spliceOffsets {
			// it can happen that the string is empty, therefore we need
			// <= and not just <
			if pos <= spliceOffset {
				limitpos = spliceOffset
				break
			}
		}

		var retstr []byte
		if pos+length <= limitpos {
			// character array is not split between records
			retstr = recordData[pos : pos+length]
			pos += length
		} else {
			// character array is split between records

			// first part of character array
			retstr = append(retstr, recordData[pos:limitpos]...)
			bytesRead := limitpos - pos
			if !isCompressed {
				bytesRead /= 2
			}
			charsLeft := numChars - uint16(bytesRead)
			pos = limitpos

			// keep reading the characters
			for charsLeft > 0 {
				// look up next limit position, in case the string span more than one continue record
				for _, spliceOffset := range spliceOffsets {
					if pos < spliceOffset {
						limitpos = spliceOffset
						break
					}
				}

				// repeated option flags
				// OpenOffice.org documentation 5.21
				option := recordData[pos]
				pos++

				if isCompressed && option == 0 {
					// 1st fragment compressed
					// this fragment compressed
					length = min(int(charsLeft), limitpos-pos)
					retstr = append(retstr, recordData[pos:pos+length]...)
					charsLeft -= uint16(length)
					isCompressed = true
				} else if !isCompressed && option != 0 {
					// 1st fragment uncompressed
					// this fragment uncompressed
					length = min(int(charsLeft)*2, limitpos-pos)
					retstr = append(retstr, recordData[pos:pos+length]...)
					charsLeft -= uint16(length / 2)
					isCompressed = false
				} else if !isCompressed && option == 0 {
					// 1st fragment uncompressed
					// this fragment compressed
					length = min(int(charsLeft), limitpos-pos)
					for j := 0; j < length; j++ {
						retstr = append(retstr, recordData[pos+j], 0)
					}
					charsLeft -= uint16(length)
					isCompressed = false
				} else {
					// 1st fragment compressed
					// this fragment uncompressed
					newstr := make([]byte, 0, len(retstr)*2)
					for _, b := range retstr {
						newstr = append(newstr, b, 0)
					}
					retstr = newstr
					length = min(int(charsLeft)*2, limitpos-pos)
					retstr = append(retstr, recordData[pos:pos+length]...)
					charsLeft -= uint16(length / 2)
					isCompressed = false
				}

				pos += length
			}
		}

		// convert to UTF-8
		retstrStr := xls.decodeCodepage(string(retstr))

		// read additional Rich-Text information, if any
		var fmtRuns []map[string]uint16
		if hasRichText {
			// list of formatting runs
			for j := 0; j < int(formattingRuns); j++ {
				// first formatted character; zero-based
				charPos := getUInt2d(recordData, pos+j*4)
				// index to font record
				fontIndex := getUInt2d(recordData, pos+2+j*4)
				fmtRuns = append(fmtRuns, map[string]uint16{
					"charPos":   charPos,
					"fontIndex": fontIndex,
				})
			}
			pos += int(formattingRuns) * 4
		}

		// read additional Asian phonetics information, if any
		if hasAsian {
			pos += extendedRunLength
		}

		// store the shared sting
		xls.sst = append(xls.sst, retstrStr)
	}
}

func (xls *XLS) readSheet() {
	length := getUInt2d(xls.data, xls.pos+2)
	recordData := xls.data[xls.pos+4 : xls.pos+4+int(length)]

	// offset: 0; size: 4; absolute stream position of the BOF record of the sheet
	recOffset := getInt4d(xls.data, xls.pos+4)

	// move stream pointer to next record
	xls.pos += 4 + int(length)

	// offset: 4; size: 1; sheet state
	var sheetState int
	switch recordData[4] {
	case 0x00:
		sheetState = XLS_SHEET_STATE_VISIBLE
	case 0x01:
		sheetState = XLS_SHEET_STATE_HIDDEN
	case 0x02:
		sheetState = XLS_SHEET_STATE_VERYHIDDEN
	default:
		sheetState = XLS_SHEET_STATE_VISIBLE
	}

	// offset: 5; size: 1; sheet type
	sheetType := recordData[5]

	// offset: 6; size: var; sheet name
	var recName string
	recName = string(recordData[6:])
	/*
		if xls.version == XLS_BIFF8 {
			recName = readUnicodeStringShort(recordData[6:])
		} else if xls.version == XLS_BIFF7 {
			recName = readByteStringShort(recordData[6:])
		}
	*/

	xls.sheets = append(xls.sheets, &Sheet{
		offset:     recOffset,
		name:       recName,
		sheetState: sheetState,
		sheetType:  sheetType,
		rows:       make(map[int]*Row),
	})
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

func (xls *XLS) readLabel(sheet *Sheet) {
	recordData, row, col := xls.getRecord()

	if xls.version == XLS_BIFF8 {
		stringData := xls.readUnicodeStringLong(recordData[6:])
		sheet.setValue(row, col, stringData.value, CellDataTypeString)
	} else { //if xls.version == XLS_BIFF7 {
		stringData := xls.readByteStringLong(recordData[6:])
		sheet.setValue(row, col, stringData.value, CellDataTypeString)
	}
}

func (xls *XLS) readLabelSst(sheet *Sheet) {
	recordData, row, col := xls.getRecord()

	index := getInt4d(recordData, 6)

	if index < 0 || index >= len(xls.sst) {
		return
	}

	if false {
		// TODO: rich text
	} else {
		sheet.setValue(row, col, xls.sst[index], CellDataTypeString)
	}
}

func (xls *XLS) readNumber(sheet *Sheet) {
	recordData, row, col := xls.getRecord()

	numValue := extractNumber(recordData[6:14])
	sheet.setValue(row, col, numValue, CellDataTypeNumeric)
}

func (xls *XLS) readRK(sheet *Sheet) {
	recordData, row, col := xls.getRecord()

	rknum := getInt4d(recordData, 6)
	numValue := getIEEE754(rknum)
	sheet.setValue(row, col, numValue, CellDataTypeNumeric)
}

func (xls *XLS) readBoolErr(sheet *Sheet) {
	recordData, row, col := xls.getRecord()

	// offset: 6; size: 1; the boolean value or error value
	boolErr := recordData[6]
	// offset: 7; size: 1; 0=boolean; 1=error
	isError := recordData[7]

	var value interface{}
	switch isError {
	case 0: // boolean
		value = boolErr
		sheet.setValue(row, col, value, CellDataTypeBool)
		break
	case 1: // error type
		value = CheckErrorCode(ErrorCode(boolErr))
		sheet.setValue(row, col, value, CellDataTypeError)
		break
	}
}

func (xls *XLS) getRecord() ([]byte, int, int) {
	length := int(getUInt2d(xls.data, xls.pos+2))
	recordData := xls.data[xls.pos+4 : xls.pos+4+length]
	row := getUInt2d(recordData, 0)
	col := getUInt2d(recordData, 2)
	xls.pos += 4 + length
	return recordData, int(row), int(col)
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

func (xls *XLS) decodeCodepage(data string) string {
	return ConvertFrom(data, xls.CodePage)
}

// readByteStringShort reads a byte string (8-bit string length) from the given data.
func (xls *XLS) readByteStringShort(subData []byte) *stringConvertion {
	// offset: 0; size: 1; length of the string (character count)
	ln := int(subData[0])

	// offset: 1: size: var; character array (8-bit characters)
	value := xls.decodeCodepage(string(subData[1 : 1+ln]))

	return &stringConvertion{
		value: value,
		size:  1 + ln,
	}
}

func (xls *XLS) readByteStringLong(subData []byte) *stringConvertion {
	// offset: 0; size: 2; length of the string (character count)
	ln := int(getUInt2d(subData, 0))

	// offset: 2: size: var; character array (8-bit characters)
	value := xls.decodeCodepage(string(subData[2 : 2+ln]))

	return &stringConvertion{
		value: value,
		size:  2 + ln,
	}
}

// readUnicodeString reads a Unicode string with no string length field but with a known character count.
func (xls *XLS) readUnicodeString(subData []byte, characterCount int) *stringConvertion {
	// offset: 0: size: 1; option flags
	// bit: 0; mask: 0x01; character compression (0 = compressed 8-bit, 1 = uncompressed 16-bit)
	isCompressed := (subData[0] & 0x01) == 0

	// bit: 2; mask: 0x04; Asian phonetic settings
	//hasAsian := (subData[0] & 0x04) >> 2

	// bit: 3; mask: 0x08; Rich-Text settings
	//hasRichText := (subData[0] & 0x08) >> 3

	// offset: 1: size: var; character array
	// this offset assumes richtext and Asian phonetic settings are off which is generally wrong
	// needs to be fixed
	var value string
	if isCompressed {
		value = string(subData[1 : 1+characterCount])
	} else {
		utf16Data := make([]uint16, characterCount)
		for i := 0; i < characterCount; i++ {
			utf16Data[i] = uint16(subData[1+2*i]) | uint16(subData[2+2*i])<<8
		}
		value = string(utf16.Decode(utf16Data))
	}

	return &stringConvertion{
		value: value,
		size: 1 + characterCount*func() int {
			if isCompressed {
				return 1
			} else {
				return 2
			}
		}(),
	}
}

func (xls *XLS) readUnicodeStringShort(subData []byte) *stringConvertion {
	// offset: 0; size: 1; length of the string (character count)
	ln := int(subData[0])

	ret := xls.readUnicodeString(subData[1:], ln)

	ret.size += 1 // size in bytes of data structure
	return ret
}

func (xls *XLS) readUnicodeStringLong(subData []byte) *stringConvertion {
	// offset: 0; size: 2; length of the string (character count)
	ln := int(getUInt2d(subData, 0))

	ret := xls.readUnicodeString(subData[2:], ln)

	ret.size += 2 // size in bytes of data structure
	return ret
}

func (xls *XLS) getSplicedRecordData() *sstConvertion {
	var data []byte
	spliceOffsets := []int{0}

	i := 0

	for {
		i++

		// offset: 0; size: 2; identifier
		//identifier := getUInt2d(xls.data, xls.pos)
		// offset: 2; size: 2; length
		length := getUInt2d(xls.data, xls.pos+2)
		data = append(data, xls.data[xls.pos+4:xls.pos+4+int(length)]...)

		spliceOffsets = append(spliceOffsets, spliceOffsets[i-1]+int(length))

		xls.pos += 4 + int(length)
		nextIdentifier := getUInt2d(xls.data, xls.pos)
		if nextIdentifier != XLS_TYPE_CONTINUE {
			break
		}
	}

	return &sstConvertion{
		recordData:    data,
		spliceOffsets: spliceOffsets,
	}
}
