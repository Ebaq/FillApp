package readfiles

import (
	"bytes"
	"encoding/json"
	"fillappgo/backend/Errors"
	"fillappgo/backend/consts"
	"fmt"
	"github.com/xbmlz/uniconv"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

var Products consts.Products
var Standard string
var ProductsGuard consts.Products
var StandardGuard string

// ReadXlsx Чтение накладных
func ReadXlsx(path string, dayOfWeek string, logger *log.Logger) error {
	productsChan := make(chan consts.Products)
	standardChan := make(chan string)

	logger.Println("Открытие накладной XLSX")
	file, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Println(err)
		logger.Printf("Ошибка при открытии накладной XLSX: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("0X1", "read", "Не удалось открыть файл"))
	}

	defer func() {
		println("closing file")
		if err := file.Close(); err != nil {
			logger.Printf("Ошибка при закрытии накладной XLSX: %v\n", err)
			fmt.Println(err)
		}
	}()

	logger.Println("Начало чтение строк накладной XLSX")
	rows, err := file.GetRows(dayOfWeek)

	if err != nil {
		logger.Printf("Ошибка при чтении строк накладной XLSX: %v\n", err)
		file.Close()
		return fmt.Errorf(Errors.NewProgramError("0X3", "read", "Не удалось прочитать файл"))
	}

	logger.Println("Начало поиска нормы и продуктов накландой XLSX")
	go findAllProducts(rows, productsChan)
	go findStandard(rows, standardChan)

	for i := 0; i < 2; i++ {
		select {
		case prod := <-productsChan:
			Products = prod
		case std := <-standardChan:
			Standard = std
		}
	}
	close(productsChan)
	close(standardChan)

	printVars()
	return nil
}

// Поиск продуктов и количества
func findAllProducts(rows [][]string, productsChan chan<- consts.Products) {
	var products consts.Products
	for _, row := range rows[10:] {
		if len(row) > 7 && row[0] != "0" && row[1] != "" && row[0] != "" && row[7] != "" {
			if amount, err := strconv.ParseFloat(row[7], 4); err == nil && amount > 0 {
				if ind, isContains := products.Contains(row[0]); isContains {
					am, _ := strconv.ParseFloat(strings.TrimSpace(row[7]), 64)
					products[ind].Amount += am

				} else {
					am, _ := strconv.ParseFloat(strings.TrimSpace(row[7]), 64)
					products = append(products, consts.Product{
						Name:     row[1],
						Amount:   am,
						IsParsed: false,
					})
				}
			}
		}
	}
	productsChan <- products
}

// Поиск нормы
func findStandard(rows [][]string, standardChan chan<- string) {
	standard := ""
std:
	for _, row := range rows {
		for _, cell := range row {
			if strings.Contains(cell, "Норма") {
				standard = cell
				standardChan <- standard
				break std
			}
		}
	}
}

func printVars() {
	println("Products")
	for _, prod := range Products {
		println(prod.Name, strconv.FormatFloat(prod.Amount, 'f', 4, 64))
	}
	println("Standard", Standard)

	println(len(Products))
}

// ReadXls TODO: Доделать функцию и добавить логи
func ReadXls(path string, logger *log.Logger) error {
	productsChan := make(chan consts.Products)
	//amountsChan := make(chan []float64)
	standardChan := make(chan string)
	p := uniconv.NewProcessor()
	logger.Println("Открытие конвертирования накладной XLS в XLSX")
	p.Start()
	defer p.Stop()

	c := uniconv.NewConverter()

	parts := strings.Split(path, "\\")
	fmt.Println(len(parts))
	fmt.Printf("parts before: %s\n", strings.Join(parts, "\\"))
	parts = parts[:len(parts)-1]
	parts = append(parts, `xlsx`)
	dir := strings.Join(parts, "\\")

	err := c.Convert(path, fmt.Sprintf("%s\\out.xlsx", dir))
	if err != nil {
		logger.Printf("Ошибка при конвертировании накладной XLS: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("1X1", "read", "Не удалось конвертировать накладную в формат XLSX"))
	}

	file, err := excelize.OpenFile(fmt.Sprintf("%s\\out.xlsx", dir))
	if err != nil {
		fmt.Println(err)
		logger.Printf("Ошибка при открытии накладной XLSX: %v\n", err)
		return fmt.Errorf(Errors.NewProgramError("1X2", "read", "Не удалось открыть накладную"))
	}
	if err := file.SaveAs(path[:len(path)-5] + "-copy.xlsx"); err != nil {
		fmt.Println(err)
		//return fmt.Errorf(Errors.NewProgramError("0X2", "read", "Не удалось сохранить копию файла"))
	}
	defer func() {
		println("closing file")
		if err := file.Close(); err != nil {
			logger.Printf("Ошибка при закрытии накладной XLS: %v\n", err)
			fmt.Println(err)
		}
	}()

	rows, err := file.GetRows("Накладная")

	if err != nil {
		logger.Printf("Ошибка при чтении строк накладной XLS: %v\n", err)
		fmt.Println(err, "Closing file")
		file.Close()
		return fmt.Errorf(Errors.NewProgramError("1X3", "read", "Не удалось прочитать файл"))
		//return fmt.Errorf(Errors.NewProgramError("0X3", "read", "Не удалось прочитать файл"))
	}

	logger.Println("Начало поиска нормы и продуктов накландой XLS")
	go findAllProducts(rows, productsChan)
	go findStandard(rows, standardChan)

	for i := 0; i < 2; i++ {
		select {
		case prod := <-productsChan:
			ProductsGuard = prod
		case std := <-standardChan:
			StandardGuard = std
		}
	}
	close(productsChan)
	close(standardChan)
	printVars()
	return nil
}

func ReadDictionary(logger *log.Logger) (consts.Dictionaries, error) {
	var dict consts.Dictionaries
	execPath, err := os.Executable()
	if err != nil {
		logger.Printf("Ошибка при поиске пути к исполняемому файлу: %v\n", err)
		return consts.Dictionaries{}, err
	}

	programDir := filepath.Dir(execPath)
	dictFilePath := filepath.Join(programDir, consts.DictionaryName)

	b, err := os.ReadFile(dictFilePath)

	if err != nil {
		logger.Printf("Ошибка при открытии файла со словарем: %v\n", err)
		return consts.Dictionaries{}, err
	}

	err = json.NewDecoder(bytes.NewBuffer(b)).Decode(&dict)
	if err != nil {
		logger.Printf("Ошибка при декодировании файла со словарем: %v\n", err)
		return consts.Dictionaries{}, err
	}

	return dict, nil
}

func ParseErrors(products consts.Products, logger *log.Logger) (consts.Products, error) {
	dict, err := ReadDictionary(logger)
	p := products

	if err != nil {
		return products, err
	}

	for i, product := range p {
		if original, isContains := dict.Contains(product.Name); isContains {
			p[i].Name = original
		}
	}

	return p, nil
}
