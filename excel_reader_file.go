package excel

import (
	"github.com/xuri/excelize/v2"
)

type fileReader struct {
	*readerImpl
}

func NewFileReader(filepath string, options ...ReaderOption) (Reader, error) {
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		return nil, err
	}
	defer func() {
		_ = f.Close()
	}()

	return &fileReader{
		readerImpl: newReader(f, options...),
	}, nil
}
