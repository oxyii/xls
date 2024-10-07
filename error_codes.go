package xls

type ErrorCode int

const (
	ErrCodesNull  ErrorCode = iota // #NULL!
	ErrCodesDiv0                   // #DIV/0!
	ErrCodesValue                  // #VALUE!
	ErrCodesRef                    // #REF!
	ErrCodesName                   // #NAME?
	ErrCodesNum                    // #NUM!
	ErrCodesNA                     // #N/A
)

/*
	func GetErrorCodes() []ErrorCode {
		return []ErrorCode{
			ErrCodesNull,
			ErrCodesDiv0,
			ErrCodesValue,
			ErrCodesRef,
			ErrCodesName,
			ErrCodesNum,
			ErrCodesNA,
		}
	}
*/
func getErrorValues() map[ErrorCode]string {
	return map[ErrorCode]string{
		ErrCodesNull:  "#NULL!",
		ErrCodesDiv0:  "#DIV/0!",
		ErrCodesValue: "#VALUE!",
		ErrCodesRef:   "#REF!",
		ErrCodesName:  "#NAME?",
		ErrCodesNum:   "#NUM!",
		ErrCodesNA:    "#N/A",
	}
}

func CheckErrorCode(code ErrorCode) string {
	values := getErrorValues()
	if val, ok := values[code]; ok {
		return val
	}

	return values[ErrCodesNull]
}
