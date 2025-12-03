package excel

import (
	"net/http"

	"github.com/xuri/excelize/v2"
)

type httpExcelReader struct {
	*readerImpl
}

func NewHttpReader(url string, options ...ReaderOption) (Reader, error) {
	resp, err := http.Get(url)
	if err != nil {
		return nil, err
	}
	defer func() {
		_ = resp.Body.Close()
	}()

	// 读取需要处理的源excel文件
	f, err := excelize.OpenReader(resp.Body)
	if err != nil {
		return nil, err
	}

	return &httpExcelReader{
		readerImpl: newReader(f, options...),
	}, nil
}
