package processing

import (
	"encoding/json"
	"fillappgo/backend/Errors"
	"fillappgo/backend/consts"
	"fillappgo/backend/readfiles"
	"fillappgo/backend/shared"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"math"
	"os"
	"os/exec"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
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

var filterProducts = consts.Products{{
	Name:     "Капуста белокочанная свежая",
	Amount:   0,
	IsParsed: false,
}, {
	Name:     "Капуста белокочанная маринованная",
	Amount:   0,
	IsParsed: false,
}, {
	Name:     "Капуста белокочанная сушеная",
	Amount:   0,
	IsParsed: false,
}, {
	Name:     "Капуста белокочанная квашеная",
	Amount:   0,
	IsParsed: false,
}, {
	Name:     "Лук репчатый свежий",
	Amount:   0,
	IsParsed: false,
}, {
	Name:     "Лук репчатый сушеный",
	Amount:   0,
	IsParsed: false,
}}

var literProducts = consts.Products{{
	Name:     "Молоко 3,2%",
	Amount:   1,
	IsParsed: false,
}, {
	Name:     "Молоко 3,2% /ФилВоенторг/ 0,200л",
	Amount:   0.2,
	IsParsed: false,
}, {
	Name:     "Сок фруктово-ягодный/ФилВоенторг/1,000л/1шт",
	Amount:   1,
	IsParsed: false,
}, {
	Name:     "Сок фруктово-ягодный /ФилВоенторг/ 0,200л",
	Amount:   0.2,
	IsParsed: false,
}}

// ProcessBook Обработка книги и запуск внесения значений
func ProcessBook(Book string, date string, logger *log.Logger) ([]string, []string, error) {
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
		return nil, nil, fmt.Errorf(Errors.NewProgramError("0X1", "processing", "Не удалось открыть файл"))
	}
	if err := wb.SaveAs(Book[:len(Book)-5] + "-КОПИЯ.xlsx"); err != nil {
		logger.Printf("Ошибка при сохранении файла книги: %v\n", err)
		return nil, nil, fmt.Errorf(Errors.NewProgramError("0X2", "processing", "Не удалось сохранить копию файла"))
	}
	defer func() {
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
		return nil, nil, fmt.Errorf(Errors.NewProgramError("0X3", "processing", "Не удалось прочитать файл"))
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
		return nil, nil, fmt.Errorf(Errors.NewProgramError("0X4", "processing", "Не удалось найти строку с нужной нормой"))
	}

	logger.Println("Начало записи количества продуктов в книгу")
	return writeAmounts(wb, Book, rows, logger, products, guardProducts, standardRow, standardGuardRow)
}

// Поиск строки с нужной нормой и датой
func findStandardRow(rows [][]string, date string, standard string, c chan<- int) int {
	resIndex := -1
	dateIndex := findDateRow(rows, date)

	for i, row := range rows[dateIndex:len(rows)] {
		if len(row) > 0 {
			if strings.Contains(row[0], standard) {
				resIndex = dateIndex + i + 1
				break
			}
		}
	}

	if c != nil {
		c <- resIndex
	}

	return resIndex
}

func findDateRow(rows [][]string, date string) (dateIndex int) {
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
			if row[0] == date {
				dateIndex = i
				break
			}
		}
	}
	//}
	return
}

// Внесение количества продукта в ячейку с нормой под продуктом
func writeAmounts(wb *excelize.File, Book string, rows [][]string, logger *log.Logger, products consts.Products, guardProducts consts.Products, standardRow int, standardGuardRow int) ([]string, []string, error) {
	k := len(rows[1]) - 5
	defer func() {
		if err := recover(); err != nil {
			logger.Printf("Ошибка в writeAmounts: %s", err)
		}
	}()

	mappedProducts := products.ToMap()
	mappedGuardProducts := guardProducts.ToMap()

	for i := 1; i < len(rows[1]); i += 5 {

		if k < i {
			break
		}

		productIndex, isContains := products.ContainsMap(strings.TrimSpace(rows[1][i]), mappedProducts)
		productKIndex, kIsContains := products.ContainsMap(strings.TrimSpace(rows[1][k]), mappedProducts)

		guardProductIndex, guardIsContains := guardProducts.ContainsMap(strings.TrimSpace(rows[1][i]), mappedGuardProducts)
		guardProductKIndex, kGuardIsContains := guardProducts.ContainsMap(strings.TrimSpace(rows[1][k]), mappedGuardProducts)

		if isContains {
			cell, _ := excelize.CoordinatesToCellName(i+3, standardRow)
			wb.SetCellValue(consts.DefaultSheet, cell, shared.TruncateToFourDecimals(products[productIndex].Amount))
			products[productIndex].IsParsed = true
		}

		if kIsContains {
			cell, _ := excelize.CoordinatesToCellName(k+3, standardRow)
			wb.SetCellValue(consts.DefaultSheet, cell, shared.TruncateToFourDecimals(products[productKIndex].Amount))
			products[productKIndex].IsParsed = true
		}

		if guardIsContains {
			cell, _ := excelize.CoordinatesToCellName(i+3, standardGuardRow)
			wb.SetCellValue(consts.DefaultSheet, cell, shared.TruncateToFourDecimals(guardProducts[guardProductIndex].Amount))
			guardProducts[guardProductIndex].IsParsed = true
		}

		if kGuardIsContains {
			cell, _ := excelize.CoordinatesToCellName(k+3, standardGuardRow)
			wb.SetCellValue(consts.DefaultSheet, cell, shared.TruncateToFourDecimals(guardProducts[guardProductKIndex].Amount))
			guardProducts[guardProductKIndex].IsParsed = true
		}

		k -= 5
	}

	logger.Println("Сохранении книги")
	err := wb.SaveAs(Book)

	logger.Println("Пересчет формул")
	RunPythonSave(Book)

	if err != nil {
		logger.Printf("Ошибка при сохранении книги: %v\n", err)
		return nil, nil, fmt.Errorf(Errors.NewProgramError("2X1", "processing", "Не удалось сохранить книгу"))
	}

	var unusedProducts []string
	var unusedGuardProducts []string

	for _, prod := range products {
		if !prod.IsParsed {
			unusedProducts = append(unusedProducts, prod.Name)
		}
	}

	for _, prod := range guardProducts {
		if !prod.IsParsed {
			unusedGuardProducts = append(unusedGuardProducts, prod.Name)
		}
	}

	logger.Printf("Products: %v Guard: %v", unusedProducts, unusedGuardProducts)
	return unusedProducts, unusedGuardProducts, nil
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
	if err := wb.SaveAs(Book[:len(Book)-5] + "-КОПИЯ.xlsx"); err != nil {
		logger.Printf("Ошибка при сохранении копии книги: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("1X2", "processing", "Не удалось сохранить копию файла"))
	}
	defer func() {
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

	RunPythonSave(Book)

	return nil

}

// Внесение количества из караула и внесение карамели в карамель
func writeGuardAndCaramel(wb *excelize.File, rows [][]string, caramel consts.Products, guardRow int, caramelRow int) error {
	k := len(rows[1]) - 5
	mappedProducts := GuardProducts.ToMap()

	for i := 1; i < len(rows[1]); i += 5 {

		productIndex, isContains := GuardProducts.ContainsMap(rows[1][i], mappedProducts)
		productKIndex, kIsContains := GuardProducts.ContainsMap(rows[1][k], mappedProducts)

		if isContains {
			cell, _ := excelize.CoordinatesToCellName(i+3, guardRow)
			wb.SetCellValue(consts.DefaultSheet, cell, shared.TruncateToFourDecimals(GuardProducts[productIndex].Amount))
			GuardProducts[productIndex].IsParsed = true
		}
		if kIsContains {
			cell, _ := excelize.CoordinatesToCellName(k+3, guardRow)
			wb.SetCellValue(consts.DefaultSheet, cell, shared.TruncateToFourDecimals(GuardProducts[productKIndex].Amount))
			GuardProducts[productKIndex].IsParsed = true
		}

		if caramelRow > 0 && !caramel[0].IsParsed && caramel[0].Amount > 0 {
			_, isContains := caramel.Contains(rows[1][i])
			_, kIsContains := caramel.Contains(rows[1][k])
			if isContains {
				cell, _ := excelize.CoordinatesToCellName(i+3, caramelRow)
				wb.SetCellValue(consts.DefaultSheet, cell, shared.TruncateToFourDecimals(caramel[0].Amount))
				caramel[0].IsParsed = true
			}
			if kIsContains {
				cell, _ := excelize.CoordinatesToCellName(k+3, caramelRow)
				wb.SetCellValue(consts.DefaultSheet, cell, shared.TruncateToFourDecimals(caramel[0].Amount))
				caramel[0].IsParsed = true
			}
		}

		k -= 5
	}

	return nil
}

func CreateNewBook(Book string, date string, logger *log.Logger) error {
	defer func() {
		if err := recover(); err != nil {
			fmt.Println(err)
		}
	}()
	var wg sync.WaitGroup

	//logger.Println("Начало создания новой книги")

	logger.Println("Открытие книги")
	wb, err := excelize.OpenFile(Book)
	if err != nil {
		logger.Printf("Ошибка при открытии файла книги: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("3X1", "processing", "Не удалось открыть файл"))
	}

	if err := wb.SaveAs(Book[:len(Book)-5] + "-КОПИЯ.xlsx"); err != nil {
		logger.Printf("Ошибка при сохранении файла книги: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("3X2", "processing", "Не удалось сохранить копию файла"))
	}

	dir := filepath.Dir(Book)

	newBookPath := filepath.Join(dir, "файл для сверки 2-87 ПУСТАЯ.xlsx")

	if err := wb.SaveAs(newBookPath); err != nil {
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

	logger.Println("Закрытие старой книги")
	if err := wb.Close(); err != nil {
		logger.Printf("Ошибка при закрытии файла книги: %v\n", err)
		fmt.Println(err)
	}

	logger.Println("Открытие новой книги")
	newWb, err := excelize.OpenFile(newBookPath)

	if err != nil {
		logger.Printf("Ошибка при открытии файла новой книги: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("3X4", "processing", "Не удалось открыть новую книгу"))
	}

	logger.Println("Начало чтения строк новой книги")
	rows, err = newWb.GetRows(consts.DefaultSheet)

	if err != nil {
		logger.Printf("Ошибка при чтении строк новой книги: %v\n", err)
		newWb.Close()
		return fmt.Errorf(Errors.NewProgramError("3X5", "processing", "Не удалось прочитать файл новой книги"))
	}

	inventRow := findInventRow(rows, date)
	logger.Println("Начало работы функции getResults")
	results := getResults(rows, inventRow, logger)

	wg.Add(2)
	logger.Println("Начало работы функции clearBook")
	go clearBook(newWb, rows, results, &wg)

	logger.Println("Начало работы функции replaceNumbers")
	go replaceNumbers(newWb, inventRow, &wg)

	wg.Wait()

	logger.Println("Сохранение новой книги")
	if err := newWb.SaveAs(newBookPath); err != nil {
		logger.Printf("Ошибка при сохранении файла новой книги: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("3X3", "processing", "Не удалось сохранить новую книгу"))
	}

	err = RunPythonSave(newBookPath)

	if err != nil {
		logger.Printf("Ошибка в пересчете формул: %s", err.Error())
	}

	return nil
}

func getResults(rows [][]string, inventRow int, logger *log.Logger) consts.Products {
	var products consts.Products

	//TODO: FIX
	for i := 4; i < len(rows[inventRow])-1; i += 5 {
		if rows[inventRow][i] != "" {
			cell, _ := strconv.ParseFloat(rows[inventRow][i], 64)
			//if cell > 0 {
			products = append(products, consts.Product{
				Name:     rows[1][i-3],
				Amount:   cell,
				IsParsed: false,
			})
		}
		//}
	}

	return products
}

func replaceNumbers(wb *excelize.File, inventRow int, wg *sync.WaitGroup) {
	defer wg.Done()
	colIndex := 5
	referenceRow := inventRow - 1
	for {
		// Адрес ячейки для строки с числами
		targetCell, err := excelize.CoordinatesToCellName(colIndex, inventRow)
		if err != nil {
			break // Завершаем цикл, если не удалось получить адрес
		}

		// Значение в целевой строке
		cellValue, _ := wb.GetCellValue(consts.DefaultSheet, targetCell)

		// Проверяем условия: не обрабатываем пустые строки, "инвентаризацию" и формулы
		if cellValue == "" || strings.ToLower(cellValue) == "инвентаризация" || strings.HasPrefix(cellValue, "=") {
			colIndex += 5
			continue
		}

		// Проверяем, является ли значение числом
		if _, err := strconv.ParseFloat(cellValue, 64); err != nil {
			colIndex += 5
			continue
		}

		// Адрес ячейки в верхней строке
		referenceCell, _ := excelize.CoordinatesToCellName(colIndex, referenceRow-1)

		// Получаем формулу из верхней строки
		formula, err := wb.GetCellFormula(consts.DefaultSheet, referenceCell)
		if err != nil || formula == "" {
			colIndex += 5
			continue
		}

		//println("Формула для ячейки: ", formula)

		// Устанавливаем формулу в текущую ячейку
		err = wb.SetCellFormula(consts.DefaultSheet, targetCell, fmt.Sprintf(`=%s`, formula))

		colIndex += 5
	}

}

func clearBook(wb *excelize.File, rows [][]string, results consts.Products, wg *sync.WaitGroup) {
	defer wg.Done()
	//for i := 1; i < len(rows[2])-1; i++ {
	//	if rows[2][i] == "приход" || rows[2][i] == "расход" {
	//		println(rows[2][i])
	//		println(len(rows[2][i]))
	//		println(len(rows))
	//		for j := 2; j < len(rows)-1; j++ {
	//			println(rows[j])
	//			if rows[j][i] != "" {
	//				cell, _ := excelize.CoordinatesToCellName(i, j)
	//				wb.SetCellValue(consts.DefaultSheet, cell, "")
	//			}
	//		}
	//	}
	//	//if k < len(rows[1])-1 {
	//	//	if index, isContains := results.Contains(rows[1][k]); isContains {
	//	//		cell, _ := excelize.CoordinatesToCellName(k+4, 4)
	//	//		wb.SetCellFloat(consts.DefaultSheet, cell, results[index].Amount, 4, 64)
	//	//	}
	//	//	k += 5
	//	//}
	//}
	//TODO: FIX

	if len(results) > 0 {
		// Обновление значений в ячейках
		for k := 1; k < len(rows[1]); k += 5 {
			cellValue := rows[1][k]

			if index, exists := results.Contains(cellValue); exists {

				if cell, err := excelize.CoordinatesToCellName(k+4, 5); err == nil {
					wb.SetCellValue(consts.DefaultSheet, cell, shared.TruncateToFourDecimals(results[index].Amount))
				}
			}
		}
	}

	// Поиск колонок с "приход" или "расход"
	var targetColIndexes []int

	if len(rows) > 2 { // Проверяем, что есть как минимум 3 строки
		for colIndex, cellValue := range rows[2] { // Проход по строке 3 (индекс 2)
			if cellValue == "приход" || cellValue == "расход" {
				targetColIndexes = append(targetColIndexes, colIndex)
			}
		}
	}

	if len(targetColIndexes) == 0 { // Если подходящих колонок не найдено
		return
	}

	// Очистка значений в найденных колонках
	for _, colIndex := range targetColIndexes { // Для каждой найденной колонки
		for rowIndex := 4; rowIndex < len(rows); rowIndex++ { // Начинаем со строки 4 (индекс 3)
			if colIndex < len(rows[rowIndex]) { // Проверяем, что индекс колонки в пределах строки
				cellValue := rows[rowIndex][colIndex]
				if cellValue != "" {
					if cellName, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1); err == nil {
						wb.SetCellValue(consts.DefaultSheet, cellName, nil) // Очищаем ячейку
					}
				}
			}
		}
	}
}

func findInventRow(rows [][]string, date string) int {
	resIndex := -1
	resCounter := 0
	dateIndex := findDateRow(rows, date)

	for i, row := range rows[dateIndex:len(rows)] {
		if len(row) > 0 {
			if strings.Contains(row[0], "Инвентаризация") {
				if resCounter == 1 {
					resIndex = dateIndex + i + 1
					break
				} else {
					resCounter++
				}
			}
		}
	}

	return resIndex
}

func RunPythonSave(path string) error {
	absPath, _ := filepath.Abs(path)
	wd, _ := os.Getwd()
	cmd := exec.Command("./shared/save.exe", absPath)
	cmd.Dir = wd

	//println(cmd.Dir)
	out, err := cmd.CombinedOutput()

	var res struct {
		Success bool   `json:"success,omitempty"`
		Error   string `json:"error,omitempty"`
	}

	if err != nil {
		return err
	}

	if err := json.Unmarshal(out, &res); err != nil {
		println(err.Error())
	}

	return err
}

func StartInventory(path string, logger *log.Logger) error {
	wb, err := excelize.OpenFile(path)

	if err != nil {
		logger.Printf("Ошибка в открытии книги: %v", err.Error())
		return err
	}

	rows, _ := wb.GetRows(consts.DefaultSheet)

	logger.Println("Начало работы функции getLastResults")
	products := getLastResults(rows)
	wb.Close()

	logger.Println("Начало работы функции createInventory")
	return createInventory(products, path, logger)

}

func createInventory(products consts.Products, path string, logger *log.Logger) error {
	invent := excelize.NewFile()
	defer invent.Close()
	var margin float64 = 0

	style, _ := invent.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size: float64(13),
		},
		Alignment: &excelize.Alignment{
			WrapText: true,
		},
	})

	index, _ := invent.NewSheet("Инвентаризация")
	invent.SetActiveSheet(index)
	err := invent.SetPageMargins("Инвентаризация", &excelize.PageLayoutMarginsOptions{
		Left:   &margin,
		Right:  &margin,
		Top:    &margin,
		Bottom: &margin,
		Header: &margin,
		Footer: &margin,
	})
	invent.SetColWidth("Инвентаризация", "A", "A", 70)
	//invent.SetColWidth("Инвентаризация", "B", "B", 15)

	if err != nil {
		println(err.Error())
		logger.Printf("Ошибка в выставлении стиля колонки: %s\n", err.Error())
	}

	err = invent.SetColStyle("Инвентаризация", "A", style)

	if err != nil {
		println(err.Error())
		logger.Printf("Ошибка в выставлении стиля колонки: %s\n", err.Error())
	}

	for row, product := range products {
		cellName, _ := excelize.CoordinatesToCellName(1, row)
		//cellAmount, _ := excelize.CoordinatesToCellName(2, row)

		if _, contains := filterProducts.Contains(product.Name); contains {
			amount := strconv.FormatFloat(math.Floor(product.Amount*100)/100, 'f', -1, 64)
			invent.SetCellValue("Инвентаризация", cellName, product.Name+" "+"("+amount+")")
		} else if i, literContains := literProducts.Contains(product.Name); literContains {
			words := strings.Split(product.Name, " ")
			var name string
			amount := strconv.FormatFloat(math.Floor(product.Amount*100)/100, 'f', -1, 64)
			if len(words) < 2 {
				name = strings.Join(words, " ") + " " + strconv.FormatFloat(literProducts[i].Amount, 'f', 1, 64) + "л" + " " + "(" + amount + ")"
			} else {
				name = strings.Join(words[:2], " ") + " " + strconv.FormatFloat(literProducts[i].Amount, 'f', 1, 64) + "л" + " " + "(" + amount + ")"
			}
			invent.SetCellValue("Инвентаризация", cellName, name)
		} else {
			var name string
			words := strings.Split(product.Name, " ")
			amount := strconv.FormatFloat(math.Floor(product.Amount*100)/100, 'f', -1, 64)
			if len(words) < 2 {
				name = strings.Join(words, " ") + " " + "(" + amount + ")"
			} else {
				name = strings.Join(words[:2], " ") + " " + "(" + amount + ")"
			}
			invent.SetCellValue("Инвентаризация", cellName, name)
		}
		//invent.SetCellValue("Инвентаризация", cellAmount, amount)
	}

	dir := filepath.Dir(path)
	savePath := filepath.Join(dir, "Данные для инвента.xlsx")
	invent.SaveAs(savePath)
	return nil
}

func getLastResults(rows [][]string) consts.Products {
	var products consts.Products
	indexRow := -1

	for k := len(rows) - 1; k >= 0; k-- {
		if len(rows[k]) > 0 {
			if rows[k][0] == "Итого" {
				indexRow = k
			}
		}
	}

	for i := 4; i < len(rows[indexRow])-1; i += 5 {
		if rows[indexRow][i] != "" {
			cell, _ := strconv.ParseFloat(rows[indexRow][i], 64)
			if cell > 0 {
				products = append(products, consts.Product{
					Name:     rows[1][i-3],
					Amount:   cell,
					IsParsed: false,
				})
			}
		}
	}

	return products
}
