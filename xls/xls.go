package xls

import (
	"encoding/binary"
	"github.com/luispater/xlsReader/cfb"
	"io"
)

// OpenFile - Open document from the file
func OpenFile(fileName string) (workbook Workbook, err error) {

	adaptor, err := cfb.OpenFile(fileName)

	if err != nil {
		return workbook, err
	}

	var book *cfb.Directory
	var root *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		fn := dir.Name()

		if fn == "Workbook" {
			if book == nil {
				book = dir
			}
		}
		if fn == "Book" {
			book = dir

		}
		if fn == "Root Entry" {
			root = dir
		}

	}

	if book != nil {
		size := binary.LittleEndian.Uint32(book.StreamSize[:])

		reader, errOpenObject := adaptor.OpenObject(book, root)

		if errOpenObject != nil {
			return workbook, errOpenObject
		}

		return readStream(reader, size)

	}

	return workbook, err
}

func OpenReader(reader io.ReadSeeker) (workbook Workbook, err error) {
	adaptor, err := cfb.OpenReader(reader)

	if err != nil {
		return workbook, err
	}

	var book *cfb.Directory
	var root *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		fn := dir.Name()

		if fn == "Workbook" {
			if book == nil {
				book = dir
			}
		}
		if fn == "Book" {
			book = dir

		}
		if fn == "Root Entry" {
			root = dir
		}

	}

	if book != nil {
		size := binary.LittleEndian.Uint32(book.StreamSize[:])

		readerOpenObject, errOpenObject := adaptor.OpenObject(book, root)

		if errOpenObject != nil {
			return workbook, errOpenObject
		}

		return readStream(readerOpenObject, size)

	}

	return workbook, err
}

func readStream(reader io.ReadSeeker, streamSize uint32) (workbook Workbook, err error) {

	stream := make([]byte, streamSize)

	_, err = reader.Read(stream)

	if err != nil {
		return workbook, nil
	}

	err = workbook.read(stream)

	if err != nil {
		return workbook, nil
	}

	for k := range workbook.sheets {
		sheet, errGetSheet := workbook.GetSheet(k)

		if errGetSheet != nil {
			return workbook, nil
		}

		err = sheet.read(stream)

		if err != nil {
			return workbook, nil
		}
	}

	return
}
