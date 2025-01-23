package readfiles

import (
	"bytes"
	"encoding/json"
	"fillappgo/backend/Errors"
	"fillappgo/backend/consts"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"os/exec"
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

	for _, prod := range Products {
		println(prod.Name, prod.Amount)
	}

	return nil
}

// Поиск продуктов и количества
func findAllProducts(rows [][]string, productsChan chan<- consts.Products) {
	var products consts.Products
	for _, row := range rows[10:] {
		if len(row) > 7 && row[0] != "0" && row[1] != "" && row[0] != "" && row[7] != "" {
			if amount, err := strconv.ParseFloat(row[7], 4); err == nil && amount > 0 {
				if ind, isContains := products.Contains(row[1]); isContains {
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

// ReadXls TODO: Доделать функцию и добавить логи
func ReadXls(path string, logger *log.Logger) error {
	absPath, _ := filepath.Abs(path)
	wd, _ := os.Getwd()
	cmd := exec.Command("./shared/xls.exe", absPath)
	cmd.Dir = wd

	//println(cmd.Dir)
	logger.Printf("Рабочая директория: %s", cmd.Dir)
	out, err := cmd.CombinedOutput()
	if err != nil {
		println(err.Error())
	}

	var res consts.PythonResult

	if err := json.Unmarshal(out, &res); err != nil {
		println(err.Error())
	}
	ProductsGuard = res.Products
	StandardGuard = res.Standard
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
	dictFilePath := filepath.Join(programDir, "shared", consts.DictionaryName)

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
	mappedErrors := dict.ToMap()

	if err != nil {
		return products, err
	}

	for i, product := range p {
		if original, isContains := dict.ContainsMap(product.Name, mappedErrors); isContains {
			p[i].Name = original
		}
	}

	return p, nil
}
