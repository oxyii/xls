package xls

import (
	"errors"
	"io"
	"os"
	"strings"
)

const (
	BIG_BLOCK_SIZE                 = 0x200
	SMALL_BLOCK_SIZE               = 0x40
	PROPERTY_STORAGE_BLOCK_SIZE    = 0x80
	SMALL_BLOCK_THRESHOLD          = 0x1000
	NUM_BIG_BLOCK_DEPOT_BLOCKS_POS = 0x2C
	ROOT_START_BLOCK_POS           = 0x30
	SMALL_BLOCK_DEPOT_BLOCK_POS    = 0x3C
	EXTENSION_BLOCK_POS            = 0x44
	NUM_EXTENSION_BLOCK_POS        = 0x48
	BIG_BLOCK_DEPOT_BLOCKS_POS     = 0x4C
	SIZE_OF_NAME_POS               = 0x40
	TYPE_POS                       = 0x42
	START_BLOCK_POS                = 0x74
	SIZE_POS                       = 0x78
)

type OLE struct {
	data []byte

	numBigBlockDepotBlocks int
	rootStartBlock         int
	sbdStartBlock          int
	extensionBlock         int
	numExtensionBlocks     int

	rootentry int

	bigBlockChain   []byte
	smallBlockChain []byte
	entry           []byte

	props []Property

	wrkbook int

	summaryInformation         int
	documentSummaryInformation int
}

type Property struct {
	name       string
	typ        int
	startBlock int
	size       int
}

func readOLE(filename string) (*OLE, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer func(file *os.File) {
		_ = file.Close()
	}(file)

	// Read the file identifier
	header := make([]byte, 8)
	_, err = file.Read(header)
	if err != nil {
		return nil, err
	}

	// Check OLE identifier
	identifierOle := []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}
	if !equal(header, identifierOle) {
		return nil, errors.New("the filename is not recognised as an OLE file")
	}

	// Get the file data
	data, err := io.ReadAll(file)
	if err != nil {
		return nil, err
	}

	ole := &OLE{data: append(header, data...)}

	ole.parseHeaders()

	return ole, nil
}

func (ole *OLE) parseHeaders() {
	// Total number of sectors used for the SAT
	ole.numBigBlockDepotBlocks = getInt4d(ole.data, NUM_BIG_BLOCK_DEPOT_BLOCKS_POS)

	// SecID of the first sector of the directory stream
	ole.rootStartBlock = getInt4d(ole.data, ROOT_START_BLOCK_POS)

	// SecID of the first sector of the SSAT (or -2 if not extant)
	ole.sbdStartBlock = getInt4d(ole.data, SMALL_BLOCK_DEPOT_BLOCK_POS)

	// SecID of the first sector of the MSAT (or -2 if no additional sectors are used)
	ole.extensionBlock = getInt4d(ole.data, EXTENSION_BLOCK_POS)

	// Total number of sectors used by MSAT
	ole.numExtensionBlocks = getInt4d(ole.data, NUM_EXTENSION_BLOCK_POS)

	// Read the big block depot blocks
	bigBlockDepotBlocks := make([]int, ole.numBigBlockDepotBlocks)
	pos := BIG_BLOCK_DEPOT_BLOCKS_POS
	bbdBlocks := ole.numBigBlockDepotBlocks

	if ole.numExtensionBlocks != 0 {
		bbdBlocks = (BIG_BLOCK_SIZE - BIG_BLOCK_DEPOT_BLOCKS_POS) / 4
	}

	for i := 0; i < bbdBlocks; i++ {
		bigBlockDepotBlocks[i] = getInt4d(ole.data, pos)
		pos += 4
	}

	for j := 0; j < ole.numExtensionBlocks; j++ {
		pos = (ole.extensionBlock + 1) * BIG_BLOCK_SIZE
		blocksToRead := min(ole.numBigBlockDepotBlocks-bbdBlocks, BIG_BLOCK_SIZE/4-1)

		for i := bbdBlocks; i < bbdBlocks+blocksToRead; i++ {
			bigBlockDepotBlocks[i] = getInt4d(ole.data, pos)
			pos += 4
		}

		bbdBlocks += blocksToRead
		if bbdBlocks < ole.numBigBlockDepotBlocks {
			ole.extensionBlock = getInt4d(ole.data, pos)
		}
	}

	// Read the big block chain
	pos = 0
	ole.bigBlockChain = make([]byte, 0)
	bbs := BIG_BLOCK_SIZE / 4
	for i := 0; i < ole.numBigBlockDepotBlocks; i++ {
		pos = (bigBlockDepotBlocks[i] + 1) * BIG_BLOCK_SIZE
		ole.bigBlockChain = append(ole.bigBlockChain, ole.data[pos:pos+4*bbs]...)
		pos += 4 * bbs
	}

	// Read the small block chain
	sbdBlock := ole.sbdStartBlock
	ole.smallBlockChain = make([]byte, 0)
	for sbdBlock != -2 {
		pos = (sbdBlock + 1) * BIG_BLOCK_SIZE
		ole.smallBlockChain = append(ole.smallBlockChain, ole.data[pos:pos+4*bbs]...)
		pos += 4 * bbs
		sbdBlock = getInt4d(ole.bigBlockChain, sbdBlock*4)
	}

	// Read the directory stream
	block := ole.rootStartBlock
	ole.entry = ole.readData(block)

	ole.readPropertySets()
}

func (ole *OLE) readPropertySets() {
	offset := 0
	entryLen := len(ole.entry)
	for offset < entryLen {
		d := ole.entry[offset : offset+PROPERTY_STORAGE_BLOCK_SIZE]
		nameSize := int(d[SIZE_OF_NAME_POS]) | (int(d[SIZE_OF_NAME_POS+1]) << 8)
		typ := int(d[TYPE_POS])
		startBlock := getInt4d(d, START_BLOCK_POS)
		size := getInt4d(d, SIZE_POS)
		name := strings.ReplaceAll(string(d[:nameSize]), "\x00", "")
		ole.props = append(ole.props, Property{name: name, typ: typ, startBlock: startBlock, size: size})
		upName := strings.ToUpper(name)
		if upName == "WORKBOOK" || upName == "BOOK" {
			ole.wrkbook = len(ole.props) - 1
		} else if upName == "ROOT ENTRY" || upName == "R" {
			ole.rootentry = len(ole.props) - 1
		} else if name == string([]byte{5})+"SummaryInformation" {
			ole.summaryInformation = len(ole.props) - 1
		} else if name == string([]byte{5})+"DocumentSummaryInformation" {
			ole.documentSummaryInformation = len(ole.props) - 1
		}
		offset += PROPERTY_STORAGE_BLOCK_SIZE
	}
}

func (ole *OLE) getStream(stream int) []byte {
	if stream == -1 {
		return nil
	}

	var streamData []byte

	if ole.props[stream].size < SMALL_BLOCK_THRESHOLD {
		rootdata := ole.readData(ole.props[ole.rootentry].startBlock)

		block := ole.props[stream].startBlock

		for block != -2 {
			pos := block * SMALL_BLOCK_SIZE
			streamData = append(streamData, rootdata[pos:pos+SMALL_BLOCK_SIZE]...)

			block = getInt4d(ole.smallBlockChain, block*4)
		}

		return streamData
	}

	numBlocks := ole.props[stream].size / BIG_BLOCK_SIZE
	if ole.props[stream].size%BIG_BLOCK_SIZE != 0 {
		numBlocks++
	}

	if numBlocks == 0 {
		return nil
	}

	block := ole.props[stream].startBlock

	for block != -2 {
		pos := (block + 1) * BIG_BLOCK_SIZE
		streamData = append(streamData, ole.data[pos:pos+BIG_BLOCK_SIZE]...)

		block = getInt4d(ole.bigBlockChain, block*4)
	}

	return streamData
}

func (ole *OLE) readData(block int) []byte {
	data := make([]byte, 0)
	for block != -2 {
		pos := (block + 1) * BIG_BLOCK_SIZE
		data = append(data, ole.data[pos:pos+BIG_BLOCK_SIZE]...)
		block = getInt4d(ole.bigBlockChain, block*4)
	}
	return data
}
