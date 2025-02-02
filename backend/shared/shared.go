package shared

import (
	"fillappgo/backend/consts"
	"github.com/shirou/gopsutil/process"
	"log"
	"math"
	"os"
	"path/filepath"
)

func KillExcel() error {
	processes, err := process.Processes()
	if err != nil {
		return err
	}
	for _, p := range processes {
		n, _ := p.Name()
		if n == consts.ExcelProcessName {
			return p.Kill()
		}
	}
	return nil
}

func OpenLogger() (*log.Logger, *os.File) {
	println("Открытие файла логов")
	execPath, err := os.Executable()
	if err != nil {
		log.Printf("Ошибка при получении пути исполняемого файла: %v\n", err)
	}

	programDir := filepath.Dir(execPath)
	logFilePath := filepath.Join(programDir, "shared", consts.LogFileName)

	file, err := os.OpenFile(logFilePath, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)

	if err != nil {
		log.Printf("Ошибка при открытии файла логов: %v\n", err)
	}

	logger := log.New(file, "INFO: ", log.Ldate|log.Ltime|log.Lshortfile)

	return logger, file
}

func TruncateToFourDecimals(value float64) float64 {
	return math.Floor(value*10000) / 10000
}
