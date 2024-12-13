package main

import (
	"context"
	"fillappgo/backend/Errors"
	"fillappgo/backend/consts"
	"fillappgo/backend/processing"
	"fillappgo/backend/readfiles"
	"fillappgo/backend/shared"
	"fmt"
	"github.com/wailsapp/wails/v2/pkg/runtime"
	"log"
)

// App struct
type App struct {
	ctx context.Context
}

// Book Путь к файлу книги
var Book string

// NewApp creates a new App application struct
func NewApp() *App {
	return &App{}
}

// startup is called when the app starts. The context is saved
// so we can call the runtime methods
func (a *App) startup(ctx context.Context) {
	a.ctx = ctx
}

// Greet returns a greeting for the given name
func (a *App) Greet(name string) string {
	return fmt.Sprintf("Hello %s, It's show time!", name)
}

func (a *App) SelectFile(dayOfWeek string) (string, error) {
	file, err := runtime.OpenFileDialog(a.ctx, runtime.OpenDialogOptions{})
	defer func() {
		if err := recover(); err != nil {
			fmt.Println(err)
		}
	}()
	if err != nil || len(file) == 0 {
		return "", fmt.Errorf(Errors.NewProgramError("0x1", "main", "Файл не выбран"))
	}

	logger, logFile := shared.OpenLogger()

	defer func() {
		println("Закрытие файла логов")
		err := logFile.Close()
		if err != nil {
			log.Printf("Ошибка при открытии файла логов: %v\n", err)
		}
	}()

	if string(file[len(file)-4:]) == "xlsx" {
		logger.Println("Начало работы ReadXlsx")
		err = readfiles.ReadXlsx(file, dayOfWeek, logger)
		if err != nil {
			return "", err
		}
	} else if string(file[len(file)-3:]) == "xls" {
		logger.Println("Начало работы ReadXls")
		err = readfiles.ReadXls(file, logger)
		if err != nil {
			return "", err
		}
	} else {
		logger.Println("Ошибка открытия файла накладной, неверный тип файла")
		return "", fmt.Errorf(Errors.NewProgramError("0x2", "main", "Не допустимый формат файла"))
	}

	return fmt.Sprintf(file), nil
}

func (a *App) SelectBook() (string, error) {
	file, err := runtime.OpenFileDialog(a.ctx, runtime.OpenDialogOptions{})
	defer func() {
		if err := recover(); err != nil {
			fmt.Println(err)
		}
	}()

	if err != nil || len(file) == 0 {
		fmt.Println(err)
		return "", fmt.Errorf(Errors.NewProgramError("1x1", "main", "Файл не выбран"))
	}
	if string(file[len(file)-4:]) != "xlsx" {
		return "", fmt.Errorf(Errors.NewProgramError("1x2", "main", "Не допустимый формат файла"))
	}
	Book = file
	println(file)
	return fmt.Sprintf(file), nil
}

func (a *App) StartFill(date string) ([]string, error) {
	err := shared.KillExcel()

	logger, file := shared.OpenLogger()
	if err != nil {
		println(err)
		logger.Printf("Ошибка при закрытии процесса Excel: %v\n", err)
		return nil, fmt.Errorf(Errors.NewProgramError("2X1", "main", "Не удалось завершить работу Excel. Закройте все вкладки с Excel"))
	}

	defer func() {
		println("Закрытие файла логов")
		err := file.Close()
		if err != nil {
			log.Printf("Ошибка при открытии файла логов: %v\n", err)
		}
	}()

	logger.Println("Начало работы ProcessBook")
	if len(readfiles.Products) < 1 || len(readfiles.ProductsGuard) < 1 {
		logger.Println("Ошибка при запуске, нет продуктов")
		return nil, fmt.Errorf(Errors.NewProgramError("2X2", "main", "Не было найдено ни одного продукта"))
	} else {
		return processing.ProcessBook(Book, date, logger)
	}
}

func (a *App) FillGuardAndCaramel(path string, caramel string, date string) error {
	defer func() {
		if err := recover(); err != nil {
			fmt.Println(err)
		}
	}()
	Book = path
	err := shared.KillExcel()

	if err != nil {
		return fmt.Errorf(Errors.NewProgramError("3X1", "main", "Не удалось завершить работу Excel. Закройте все вкладки с Excel"))
	}
	logger, file := shared.OpenLogger()

	defer func() {
		println("Закрытие файла логов")
		err := file.Close()
		if err != nil {
			log.Printf("Ошибка при открытии файла логов: %v\n", err)
		}
	}()

	logger.Println("Начало работы SetGuardAndCaramel")
	return processing.SetGuardAndCaramel(path, caramel, date, logger)
}

func (a *App) ResetProducts() {
	readfiles.Products = consts.Products{}
	readfiles.ProductsGuard = consts.Products{}
	readfiles.Standard = ""
	readfiles.StandardGuard = ""
}

func (a *App) CreateNewBook(path string) error {
	defer func() {
		if err := recover(); err != nil {
			fmt.Println(err)
		}
	}()
	Book = path
	err := shared.KillExcel()

	if err != nil {
		return fmt.Errorf(Errors.NewProgramError("4X1", "main", "Не удалось завершить работу Excel. Закройте все вкладки с Excel"))
	}
	logger, file := shared.OpenLogger()

	defer func() {
		println("Закрытие файла логов")
		err := file.Close()
		if err != nil {
			log.Printf("Ошибка при открытии файла логов: %v\n", err)
		}
	}()

	logger.Println("Начало работы CreateNewBook")
	return processing.CreateNewBook(Book, logger)
}
