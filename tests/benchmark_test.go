package tests

import (
	"fillappgo/backend/processing"
	"log"
	"testing"
)

func BenchmarkProcessBook(b *testing.B) {
	logger := log.New(nil, "", log.LstdFlags)
	bookPath := "path/to/your/test/file.xlsx"
	date := "01.01.2023"

	for i := 0; i < b.N; i++ {
		_, _, err := processing.ProcessBook(bookPath, date, logger)
		if err != nil {
			b.Fatalf("Ошибка: %v", err)
		}
	}
}
