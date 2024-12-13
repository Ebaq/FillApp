package consts

import (
	"regexp"
)

// DefaultSheet Название листа с книгой
const DefaultSheet = "Книга кладовщика"
const GuardStandard = `Норма 1 п. 3 "ж" (караул)`
const CaramelStandard = "Норма 1  п.п. 6 (карамель)"
const DictionaryName = `dict.json`
const ExcelProcessName = "EXCEL.EXE"
const LogFileName = `logFile.log`

var filter = regexp.MustCompile("[^a-zA-Zа-яА-Я0-9]")

var Dictionary Dictionaries

type Product struct {
	Name     string
	Amount   float64
	IsParsed bool
}

type Products []Product

func (products Products) Contains(prod string) (index int, contains bool) {
	for i, product := range products {
		if clean(product.Name) == clean(prod) {
			return i, true
		}
	}
	return -1, false
}

type Dict struct {
	Name  string `json:name`
	Error string `json:error`
}

type Dictionaries struct {
	Dictionary []Dict `json:dictionary`
}

func (dictionaries Dictionaries) Contains(prod string) (original string, contains bool) {
	for _, product := range dictionaries.Dictionary {
		if clean(product.Error) == clean(prod) {
			return product.Name, true
		}
	}
	return "", false
}

func clean(s string) string {
	return filter.ReplaceAllString(s, "")
}
