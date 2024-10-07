package xls

import (
	"encoding/binary"
	"math"
)

// ReadUInt2d reads a 16-bit unsigned integer from the given byte slice at the specified position.
func getUInt2d(data []byte, pos int) uint16 {
	if pos < 0 || pos+2 > len(data) {
		return 0
	}
	return binary.LittleEndian.Uint16(data[pos : pos+2])
}

/*
// Read 16-bit signed integer.
func getInt2d(data []byte, pos int) int {
	if pos < 0 || pos+2 > len(data) {
		return 0 //, errors.New(fmt.Sprintf("Parameter pos=%d is invalid.", pos))
	}

	var result int
	buf := bytes.NewReader(data[pos : pos+2])
	err := binary.Read(buf, binary.LittleEndian, &result)
	if err != nil {
		return 0 //, err
	}

	return result //, nil
}
*/
// Read 32-bit signed integer.
func getInt4d(data []byte, pos int) int {
	if pos < 0 {
		return 0 //, errors.New(fmt.Sprintf("Parameter pos=%d is invalid.", pos))
	}

	if len(data) < pos+4 {
		padding := make([]byte, pos+4-len(data))
		data = append(data, padding...)
	}

	// FIX: represent numbers correctly on 64-bit system
	// Changed to ensure correct result of the <<24 block on 32 and 64bit systems
	_or_24 := int(data[pos+3])
	var _ord_24 int
	if _or_24 >= 128 {
		// negative number
		_ord_24 = -int(256-_or_24) << 24
	} else {
		_ord_24 = int(_or_24&127) << 24
	}

	return int(data[pos]) | int(data[pos+1])<<8 | int(data[pos+2])<<16 | _ord_24 //, nil
}

func getIEEE754(rknum int) float64 {
	var value float64
	if (rknum & 0x02) != 0 {
		value = float64(rknum >> 2)
	} else {
		// The RK format calls for using only the most significant 30 bits
		// of the 64-bit floating point value. The other 34 bits are assumed
		// to be 0 so we use the upper 30 bits of rknum as follows...
		sign := (rknum & 0x80000000) >> 31
		exp := (rknum & 0x7ff00000) >> 20
		mantissa := 0x100000 | (rknum & 0x000ffffc)
		value = float64(mantissa) / math.Pow(2, float64(20-(exp-1023)))
		if sign != 0 {
			value = -1 * value
		}
	}
	if (rknum & 0x01) != 0 {
		value /= 100
	}
	return value
}

// extractNumber reads the first 8 bytes of a byte slice and returns an IEEE 754 float.
func extractNumber(data []byte) float64 {
	rknumhigh := int(binary.LittleEndian.Uint32(data[4:8]))
	rknumlow := int(binary.LittleEndian.Uint32(data[0:4]))
	sign := (rknumhigh & 0x80000000) >> 31
	exp := ((rknumhigh & 0x7ff00000) >> 20) - 1023
	mantissa := 0x100000 | (rknumhigh & 0x000fffff)
	mantissalow1 := (rknumlow & 0x80000000) >> 31
	mantissalow2 := rknumlow & 0x7fffffff
	value := float64(mantissa) / math.Pow(2, float64(20-exp))

	if mantissalow1 != 0 {
		value += 1 / math.Pow(2, float64(21-exp))
	}

	value += float64(mantissalow2) / math.Pow(2, float64(52-exp))
	if sign != 0 {
		value *= -1
	}

	return value
}

func equal(a, b []byte) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}
