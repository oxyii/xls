package xls

import (
	"bytes"
	"golang.org/x/text/encoding"
	"golang.org/x/text/encoding/charmap"
	"golang.org/x/text/encoding/ianaindex"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/encoding/simplifiedchinese"
	"golang.org/x/text/encoding/traditionalchinese"
	"golang.org/x/text/encoding/unicode"
	"io"
	"strings"
)

var DefaultCodePage = charmap.Windows1252

func parseCodePage(value uint16) encoding.Encoding {
	usASCII, err := ianaindex.MIME.Encoding("US-ASCII")
	if err != nil {
		return DefaultCodePage
	}
	switch value {
	case 367:
		return usASCII //    ASCII
	case 437:
		return charmap.CodePage437 //    OEM US
	case 720: //    OEM Arabic
		return DefaultCodePage
	case 737: //    OEM Greek
		return DefaultCodePage
	case 775: //    OEM Baltic
		return DefaultCodePage
	case 850:
		return charmap.CodePage850 //    OEM Latin I
	case 852:
		return charmap.CodePage852 //    OEM Latin II (Central European)
	case 855:
		return charmap.CodePage855 //    OEM Cyrillic
	case 857:
		return DefaultCodePage //    OEM Turkish
	case 858:
		return charmap.CodePage858 //    OEM Multilingual Latin I with Euro
	case 860:
		return charmap.CodePage860 //    OEM Portugese
	case 861:
		return DefaultCodePage //    OEM Icelandic
	case 862:
		return charmap.CodePage862 //    OEM Hebrew
	case 863:
		return charmap.CodePage863 //    OEM Canadian (French)
	case 864:
		return DefaultCodePage //    OEM Arabic
	case 865:
		return charmap.CodePage865 //    OEM Nordic
	case 866:
		return charmap.CodePage866 //    OEM Cyrillic (Russian)
	case 869:
		return DefaultCodePage //    OEM Greek (Modern)
	case 874:
		return charmap.Windows874 //    ANSI Thai
	case 932:
		return japanese.ShiftJIS //    ANSI Japanese Shift-JIS
	case 936:
		return simplifiedchinese.GBK //    ANSI Chinese Simplified GBK
	case 949:
		return DefaultCodePage //    ANSI Korean (Wansung)
	case 950:
		return traditionalchinese.Big5 //    ANSI Chinese Traditional BIG5
	case 1200:
		return unicode.UTF16(unicode.LittleEndian, unicode.IgnoreBOM) //    UTF-16 (BIFF8)
	case 1250:
		return charmap.Windows1250 //    ANSI Latin II (Central European)
	case 1251:
		return charmap.Windows1251 //    ANSI Cyrillic
	case 0:
		//    CodePage is not always correctly set when the xls file was saved by Apple's Numbers program
		return DefaultCodePage
	case 1252:
		return charmap.Windows1252 //    ANSI Latin I (BIFF4-BIFF7)
	case 1253:
		return charmap.Windows1253 //    ANSI Greek
	case 1254:
		return charmap.Windows1254 //    ANSI Turkish
	case 1255:
		return charmap.Windows1255 //    ANSI Hebrew
	case 1256:
		return charmap.Windows1256 //    ANSI Arabic
	case 1257:
		return charmap.Windows1257 //    ANSI Baltic
	case 1258:
		return charmap.Windows1258 //    ANSI Vietnamese
	case 1361:
		return DefaultCodePage //    ANSI Korean (Johab)
	case 10000:
		return charmap.Macintosh //    Apple Roman
	case 10001:
		return DefaultCodePage //    Macintosh Japanese
	case 10002:
		return DefaultCodePage //    Macintosh Chinese Traditional
	case 10003:
		return DefaultCodePage //    Macintosh Korean
	case 10004:
		return DefaultCodePage //    Apple Arabic
	case 10005:
		return DefaultCodePage //    Apple Hebrew
	case 10006:
		return DefaultCodePage //    Macintosh Greek
	case 10007:
		return charmap.MacintoshCyrillic //    Macintosh Cyrillic
	case 10008:
		return simplifiedchinese.HZGB2312 //    Macintosh - Simplified Chinese (GB 2312)
	case 10010:
		return DefaultCodePage //    Macintosh Romania
	case 10017:
		return DefaultCodePage //    Macintosh Ukraine
	case 10021:
		return DefaultCodePage //    Macintosh Thai
	case 10029:
		return DefaultCodePage //    Macintosh Central Europe
	case 10079:
		return DefaultCodePage //    Macintosh Icelandic
	case 10081:
		return DefaultCodePage //    Macintosh Turkish
	case 10082:
		return DefaultCodePage //    Macintosh Croatian
	case 21010:
		return unicode.UTF16(unicode.LittleEndian, unicode.IgnoreBOM) //    UTF-16 (BIFF8) This isn't correct, but some Excel writer libraries erroneously use Codepage 21010 for UTF-16LE
	case 32768:
		return charmap.Macintosh //    Apple Roman
	case 32769:
		return DefaultCodePage //    ANSI Latin I (BIFF2-BIFF3)
	case 65000:
		return DefaultCodePage //    Unicode (UTF-7)
	case 65001:
		return unicode.UTF8 //    Unicode (UTF-8)
	default:
		return DefaultCodePage
	}
}

func ConvertFrom(str string, codePage encoding.Encoding) string {
	var out bytes.Buffer
	_, _ = io.Copy(&out, codePage.NewDecoder().Reader(strings.NewReader(str)))
	return out.String()
}
