package excel

import (
	"github.com/hdget/utils/text"
	"github.com/pkg/errors"
	"github.com/xuri/excelize/v2"
)

type Reader interface {
	ReadSheet(sheetName string) (*Sheet, error)
	ReadAllSheets() ([]*Sheet, error)
}

type ReaderOption func(*readerImpl)

func WithHeaderRowIndex(index int) ReaderOption {
	return func(r *readerImpl) {
		r.headerRowIndex = index
	}
}

type readerImpl struct {
	*excelize.File
	Sheets map[string]*Sheet
	// param
	headerRowIndex int
}

func newReader(file *excelize.File, options ...ReaderOption) *readerImpl {
	return &readerImpl{
		File:           file,
		Sheets:         make(map[string]*Sheet),
		headerRowIndex: 1,
	}
}

func (r *readerImpl) ReadSheet(sheetName string) (*Sheet, error) {
	rows, err := r.Rows(sheetName)
	if err != nil {
		return nil, errors.Wrapf(err, "read rows, sheet: %s", sheetName)
	}

	// 读取表头, 跳过r.option.headerRowIndex行
	for i := 0; i < r.headerRowIndex; i++ {
		rows.Next()
	}

	headerRow, err := rows.Columns()
	if err != nil {
		return nil, errors.Wrapf(err, "read header, sheet: %s", sheetName)
	}

	headerIndexes := make(map[string]int)
	headers := make([]string, len(headerRow))
	for i, colCell := range headerRow {
		header := text.CleanString(colCell)
		headerIndexes[header] = i
		headers[i] = header
	}

	s := &Sheet{Name: sheetName, HeaderIndexes: headerIndexes, Headers: headers, Rows: make([]*SheetRow, 0)}

	// 读取数据
	for rows.Next() {
		columns, err := rows.Columns()
		if err != nil {
			return nil, errors.Wrapf(err, "read data, sheet: %s", sheetName)
		}

		line := &SheetRow{Sheet: s, Columns: make([]string, 0)}
		line.Columns = append(line.Columns, columns...)

		s.Rows = append(s.Rows, line)
	}

	return s, nil
}

func (r *readerImpl) ReadAllSheets() ([]*Sheet, error) {
	sheets := make([]*Sheet, 0)
	for _, sheetName := range r.GetSheetList() {
		sheet, err := r.ReadSheet(sheetName)
		if err != nil {
			return nil, errors.Wrapf(err, "read sheet, sheet: %s", sheetName)
		}

		sheets = append(sheets, sheet)
	}
	return sheets, nil
}
