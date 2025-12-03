package excel

import (
	"github.com/pkg/errors"
)

type FileWriter interface {
	Writer
	Save(filename string) error // save excel file
}

type fileWriterImpl struct {
	*writerImpl
}

func NewFileWriter(options ...WriterOption) (FileWriter, error) {
	w, err := NewWriter(options...)
	if err != nil {
		return nil, err
	}

	return &fileWriterImpl{
		writerImpl: w.(*writerImpl),
	}, nil
}

func (w *fileWriterImpl) Save(filename string) error {
	err := w.SaveAs(filename)
	if err != nil {
		return errors.Wrap(err, "save excel file")
	}
	return nil
}
