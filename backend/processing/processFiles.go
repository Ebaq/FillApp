package processing

import (
	"fillappgo/backend/Errors"
	"fillappgo/backend/consts"
	"fillappgo/backend/readfiles"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"strconv"
	"strings"
)

var GuardProducts = consts.Products{
	{
		Name:     "Хлеб пшеничный из муки 1сорта",
		Amount:   4.3500,
		IsParsed: false,
	},
	{
		Name:     "Колбаса п/к",
		Amount:   1.4500,
		IsParsed: false,
	},
	{
		Name:     "Масло сливочное 72,5% порционное 15гр",
		Amount:   0.4350,
		IsParsed: false,
	},
	{
		Name:     "Сахар песок",
		Amount:   0.8700,
		IsParsed: false,
	},
	{
		Name:     "Печенье /ФилВоенторг/ весовое",
		Amount:   0.5800,
		IsParsed: false,
	},
	{
		Name:     "Кофе растворимый",
		Amount:   0.0435,
		IsParsed: false,
	},
}

// ProcessBook Обработка книги и запуск внесения значений
func ProcessBook(Book string, date string, logger *log.Logger) ([]string, error) {
	defer func() {
		if err := recover(); err != nil {
			fmt.Println(err)
		}
	}()
	var standardRow int
	var standardGuardRow int
	standardRowChan := make(chan int)
	standardGuardRowChan := make(chan int)
	logger.Println("Начало обработки продуктов с помощью словаря")
	products, err := readfiles.ParseErrors(readfiles.Products, logger)
	guardProducts, err := readfiles.ParseErrors(readfiles.ProductsGuard, logger)

	if err != nil {
		logger.Printf("Ошибка при обработке продуктов с помощью словаря: %v\n", err)
		fmt.Println(err)
	}

	logger.Println("Открытие книги")
	wb, err := excelize.OpenFile(Book)
	if err != nil {
		logger.Printf("Ошибка при открытии файла книги: %v\n", err)
		return nil, fmt.Errorf(Errors.NewProgramError("0X1", "processing", "Не удалось открыть файл"))
	}
	if err := wb.SaveAs(Book[:len(Book)-5] + "-copy.xlsx"); err != nil {
		logger.Printf("Ошибка при сохранении файла книги: %v\n", err)
		return nil, fmt.Errorf(Errors.NewProgramError("0X2", "processing", "Не удалось сохранить копию файла"))
	}
	defer func() {
		println("closing wb")
		if err := wb.Close(); err != nil {
			logger.Printf("Ошибка при закрытии файла книги: %v\n", err)
			fmt.Println(err)
		}
	}()

	logger.Println("Чтение строк книги")
	rows, err := wb.GetRows(consts.DefaultSheet)

	if err != nil {
		fmt.Println(err, "Closing wb")
		logger.Printf("Ошибка при чтении строк книги: %v\n", err)
		wb.Close()
		return nil, fmt.Errorf(Errors.NewProgramError("0X3", "processing", "Не удалось прочитать файл"))
	}

	logger.Println("Начало поиска строки с нормой")
	go findStandardRow(rows, date, readfiles.Standard, standardRowChan)
	go findStandardRow(rows, date, readfiles.StandardGuard, standardGuardRowChan)

	for i := 0; i < 2; i++ {
		select {
		case s := <-standardRowChan:
			standardRow = s
		case sg := <-standardGuardRowChan:
			standardGuardRow = sg
		}
	}

	if standardRow == -1 {
		logger.Println("Не была найдена строка с нормой")
		return nil, fmt.Errorf(Errors.NewProgramError("0X4", "processing", "Не удалось найти строку с нужной нормой"))
	}

	logger.Println("Начало записи количества продуктов в книгу")
	return writeAmounts(wb, Book, rows, logger, products, guardProducts, standardRow, standardGuardRow)
}

// Поиск строки с нужной нормой и датой
func findStandardRow(rows [][]string, date string, standard string, c chan<- int) int {
	resIndex := -1
	dateIndex := 0
	//TODO: Оптимизировать, начиная с конца, если дата > 15
	//if day > 15 {
	//	for i := len(rows) - 1; i >= 0; i-- {
	//		if len(rows[i]) > 0 {
	//			if rows[i][0] == date {
	//				dateIndex = i
	//				break
	//			}
	//		}
	//	}
	//} else {
	for i, row := range rows {
		if len(row) > 0 {
			println("row", row[0])
			if row[0] == date {
				dateIndex = i
				break
			}
		}
	}
	//}

	for i, row := range rows[dateIndex:len(rows)] {
		if strings.Contains(row[0], standard) {
			resIndex = dateIndex + i + 1
			break
		}
	}

	if c != nil {
		c <- resIndex
	}

	return resIndex
}

// Внесение количества продукта в ячейку с нормой под продуктом
func writeAmounts(wb *excelize.File, Book string, rows [][]string, logger *log.Logger, products consts.Products, guardProducts consts.Products, standardRow int, standardGuardRow int) ([]string, error) {
	k := len(rows[1]) - 5

	println("starting parsing rows")
	for i := 1; i < len(rows[1]); i += 5 {

		if k < i {
			break
		}

		productIndex, isContains := products.Contains(strings.TrimSpace(rows[1][i]))
		productKIndex, kIsContains := products.Contains(strings.TrimSpace(rows[1][k]))

		guardProductIndex, guardIsContains := guardProducts.Contains(strings.TrimSpace(rows[1][i]))
		guardProductKIndex, kGuardIsContains := guardProducts.Contains(strings.TrimSpace(rows[1][k]))

		if isContains {
			cell, _ := excelize.CoordinatesToCellName(i+3, standardRow)
			wb.SetCellFloat(consts.DefaultSheet, cell, products[productIndex].Amount, 4, 64)
			products[productIndex].IsParsed = true
		}

		if kIsContains {
			cell, _ := excelize.CoordinatesToCellName(k+3, standardRow)
			wb.SetCellFloat(consts.DefaultSheet, cell, products[productKIndex].Amount, 4, 64)
			products[productKIndex].IsParsed = true
		}

		if guardIsContains {
			cell, _ := excelize.CoordinatesToCellName(i+3, standardGuardRow)
			wb.SetCellFloat(consts.DefaultSheet, cell, guardProducts[guardProductIndex].Amount, 4, 64)
			guardProducts[guardProductIndex].IsParsed = true
		}

		if kGuardIsContains {
			cell, _ := excelize.CoordinatesToCellName(k+3, standardGuardRow)
			wb.SetCellFloat(consts.DefaultSheet, cell, guardProducts[guardProductKIndex].Amount, 4, 64)
		}

		k -= 5
	}

	logger.Println("Сохранении книги")
	err := wb.SaveAs(Book)

	if err != nil {
		logger.Printf("Ошибка при сохранении книги: %v\n", err)
		return nil, fmt.Errorf(Errors.NewProgramError("2X1", "processing", "Не удалось сохранить книгу"))
	}

	logger.Println("Сохранении измененной копии книги")
	err = wb.SaveAs(Book[:len(Book)-5] + "-edited-copy.xlsx")
	if err != nil {
		logger.Printf("Ошибка при сохранении измененной копии книги: %v\n", err)
	}

	var unusedProducts []string

	for _, prod := range products {
		if !prod.IsParsed {
			unusedProducts = append(unusedProducts, prod.Name)
		}
	}

	return unusedProducts, nil
}

// SetGuardAndCaramel Открытие книги и запуск внесения караула и карамели
func SetGuardAndCaramel(Book string, caramelAmount string, date string, logger *log.Logger) error {
	var caramelRow int
	var cm float64
	var guardRow int
	var err error
	caramelChan := make(chan int)
	guardChan := make(chan int)
	caramel := consts.Products{
		{
			Name:     "Карамель /ФилВоенторг/ /весовой",
			Amount:   0,
			IsParsed: false,
		},
	}

	if len(caramelAmount) > 0 {
		cm, err = strconv.ParseFloat(caramelAmount, 64)

		if err != nil {
			logger.Printf("Ошибка при приведении карамели к Float: %v\n", err)
			return fmt.Errorf(Errors.NewProgramError("1X7", "processing", "Не удалось прочитать количество карамели"))
		}

		caramel[0].Amount = cm
	}

	logger.Println("Открытие книги")
	wb, err := excelize.OpenFile(Book)
	if err != nil {
		logger.Printf("Ошибка при открытии книги: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("1X1", "processing", "Не удалось открыть файл"))
	}
	if err := wb.SaveAs(Book[:len(Book)-5] + "-copy.xlsx"); err != nil {
		logger.Printf("Ошибка при сохранении копии книги: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("1X2", "processing", "Не удалось сохранить копию файла"))
	}
	defer func() {
		println("closing wb")
		logger.Println("Закрытии книги")
		if err := wb.Close(); err != nil {
			logger.Printf("Ошибка при закрытии книги: %v\n", err)
			fmt.Println(err)
		}
	}()

	logger.Println("Начало чтения строк книги")
	rows, err := wb.GetRows(consts.DefaultSheet)

	if err != nil {
		logger.Printf("Ошибка при чтении строк книги: %v\n", err)
		wb.Close()
		return fmt.Errorf(Errors.NewProgramError("1X3", "processing", "Не удалось прочесть строки книги"))
	}

	if caramel[0].Amount > 0 {
		logger.Println("Начало поиска нормы караула и карамели")
		go findStandardRow(rows, date, consts.CaramelStandard, caramelChan)
		go findStandardRow(rows, date, consts.GuardStandard, guardChan)
	} else {
		logger.Println("Начало поиска нормы караула")
		guardRow = findStandardRow(rows, date, consts.GuardStandard, nil)
	}

	if caramel[0].Amount > 0 {
		for i := 0; i < 2; i++ {
			select {
			case c := <-caramelChan:
				caramelRow = c
			case g := <-guardChan:
				guardRow = g
			}
		}
	}

	if caramelRow == -1 || guardRow == -1 {
		logger.Printf("Не было найдено строки с нормой караула или карамели: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("1X4", "processing", "Не удалось найти строку с нужной нормой"))
	}

	close(caramelChan)
	close(guardChan)

	logger.Println("Начало заполнения караула и карамели")
	err = writeGuardAndCaramel(wb, rows, caramel, guardRow, caramelRow)

	if err != nil {
		logger.Printf("Ошибка при заполнении караула и карамели: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("1X5", "processing", "Не удалось сохранить файл"))
	}

	err = wb.SaveAs(Book)
	if err != nil {
		logger.Printf("Ошибка при сохранении книги: %v\n", err)
		return err
	}
	err = wb.SaveAs(Book[:len(Book)-5] + "-edited-copy.xlsx")
	if err != nil {
		logger.Printf("Ошибка при сохранении измененной копии книги: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("1X6", "processing", "Не удалось сохранить измененную копию файла"))
	}

	return nil

}

// Внесение количества из караула и внесение карамели в карамель
func writeGuardAndCaramel(wb *excelize.File, rows [][]string, caramel consts.Products, guardRow int, caramelRow int) error {
	k := len(rows[1]) - 5

	for i := 1; i < len(rows[1]); i += 5 {

		productIndex, isContains := GuardProducts.Contains(rows[1][i])
		productKIndex, kIsContains := GuardProducts.Contains(rows[1][k])

		if isContains {
			cell, _ := excelize.CoordinatesToCellName(i+3, guardRow)
			wb.SetCellFloat(consts.DefaultSheet, cell, GuardProducts[productIndex].Amount, 4, 64)
			GuardProducts[productIndex].IsParsed = true
		}
		if kIsContains {
			cell, _ := excelize.CoordinatesToCellName(k+3, guardRow)
			wb.SetCellFloat(consts.DefaultSheet, cell, GuardProducts[productKIndex].Amount, 4, 64)
			GuardProducts[productKIndex].IsParsed = true
		}

		if caramelRow > 0 && !caramel[0].IsParsed && caramel[0].Amount > 0 {
			_, isContains := caramel.Contains(rows[1][i])
			_, kIsContains := caramel.Contains(rows[1][k])
			if isContains {
				cell, _ := excelize.CoordinatesToCellName(i+3, caramelRow)
				wb.SetCellFloat(consts.DefaultSheet, cell, caramel[0].Amount, 4, 64)
				caramel[0].IsParsed = true
			}
			if kIsContains {
				cell, _ := excelize.CoordinatesToCellName(k+3, guardRow)
				wb.SetCellFloat(consts.DefaultSheet, cell, caramel[0].Amount, 4, 64)
				caramel[0].IsParsed = true
			}
		}

		k -= 5
	}

	return nil
}

func CreateNewBook(Book string, logger *log.Logger) error {
	defer func() {
		if err := recover(); err != nil {
			fmt.Println(err)
		}
	}()

	//logger.Println("Начало создания новой книги")

	logger.Println("Открытие книги")
	wb, err := excelize.OpenFile(Book)
	if err != nil {
		logger.Printf("Ошибка при открытии файла книги: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("3X1", "processing", "Не удалось открыть файл"))
	}

	if err := wb.SaveAs(Book[:len(Book)-5] + "-copy.xlsx"); err != nil {
		logger.Printf("Ошибка при сохранении файла книги: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("3X2", "processing", "Не удалось сохранить копию файла"))
	}

	if err := wb.SaveAs(Book[:len(Book)-5] + "-new.xlsx"); err != nil {
		logger.Printf("Ошибка при сохранении файла новой книги: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("3X3", "processing", "Не удалось сохранить новую книгу"))
	}

	logger.Println("Начало чтения строк книги")
	rows, err := wb.GetRows(consts.DefaultSheet)

	if err != nil {
		fmt.Println(err, "Closing wb")
		logger.Printf("Ошибка при чтении строк новой книги: %v\n", err)
		wb.Close()
		return fmt.Errorf(Errors.NewProgramError("3X5", "processing", "Не удалось прочитать файл новой книги"))
	}

	logger.Println("Начало работы функции getResults")
	results := getResults(rows, logger)

	logger.Println("Закрытие старой книги")
	if err := wb.Close(); err != nil {
		logger.Printf("Ошибка при закрытии файла книги: %v\n", err)
		fmt.Println(err)
	}

	newBookPath := Book[:len(Book)-5] + "-new.xlsx"

	logger.Println("Открытие новой книги")
	wb, err = excelize.OpenFile(newBookPath)

	if err != nil {
		logger.Printf("Ошибка при открытии файла новой книги: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("3X4", "processing", "Не удалось открыть новую книгу"))
	}

	logger.Println("Начало чтения строк новой книги")
	rows, err = wb.GetRows(consts.DefaultSheet)

	if err != nil {
		fmt.Println(err, "Closing wb")
		logger.Printf("Ошибка при чтении строк новой книги: %v\n", err)
		wb.Close()
		return fmt.Errorf(Errors.NewProgramError("3X5", "processing", "Не удалось прочитать файл новой книги"))
	}

	logger.Println("Начало работы функции clearBook")
	clearBook(wb, rows, results)

	return nil
}

func getResults(rows [][]string, logger *log.Logger) consts.Products {
	var resultIndex int
	var products consts.Products

	for i := len(rows) - 1; i > 0; i-- {
		if rows[i][0] == "Итого" {
			resultIndex = i
		}
	}

	for i := 5; i < len(rows[resultIndex]); i += 5 {
		println(rows[resultIndex][i])
		if rows[resultIndex][i] != "" {
			cell, err := strconv.ParseFloat(rows[resultIndex][i], 64)
			if err != nil {
				println(rows[resultIndex][i], err)
			}
			if cell > 0 {
				products = append(products, consts.Product{
					Name:     rows[1][i-4],
					Amount:   cell,
					IsParsed: false,
				})
			}
		}
	}

	return products
}

func clearBook(wb *excelize.File, rows [][]string, results consts.Products) {
	k := 1
	for i := 1; i < len(rows[2]); i++ {
		if rows[2][i] == "приход" || rows[2][i] == "расход" {
			for j := 2; j < len(rows[2][i]); j++ {
				if rows[j][i] != "" {
					cell, _ := excelize.CoordinatesToCellName(i, j)
					wb.SetCellValue(consts.DefaultSheet, cell, "")
				}
			}
		}
		if index, isContains := results.Contains(rows[1][k]); isContains {
			cell, _ := excelize.CoordinatesToCellName(k+4, 4)
			wb.SetCellFloat(consts.DefaultSheet, cell, results[index].Amount, 4, 64)
		}
		if k+5 < len(rows[1]) {
			k += 5
		}
	}
}
