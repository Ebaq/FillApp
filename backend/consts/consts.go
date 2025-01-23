package consts

import (
	"regexp"
	"strings"
)

// DefaultSheet Название листа с книгой
const DefaultSheet = "Книга кладовщика"
const GuardStandard = `Норма 1 п. 3 "ж" (караул)`
const CaramelStandard = "Норма 1  п.п. 6 (карамель)"
const DictionaryName = `dict.json`
const ExcelProcessName = "EXCEL.EXE"
const LogFileName = `logFile.log`
const PIN = "ea1857d3373f38d477455450d0d1afb381bdf169ec37169f979f0d4cee46bc8d"

var filter = regexp.MustCompile("[^a-zA-Zа-яА-Я0-9]")

var Dictionary Dictionaries

type Product struct {
	Name     string
	Amount   float64
	IsParsed bool
}

type PythonResult struct {
	Products Products `json:"products"`
	Standard string   `json:"standard"`
	Error    string   `json:"error,omitempty"`
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

func (products Products) ToMap() map[string]int {
	result := make(map[string]int, len(products))
	for i, product := range products {
		result[clean(product.Name)] = i
	}
	return result
}

func (products Products) ContainsMap(prod string, productMap map[string]int) (index int, contains bool) {
	println("checking", prod)
	if clean(prod) == "" {
		return -1, false
	}
	if index, contains := productMap[clean(prod)]; contains {
		println(index, contains)
		return index, true
	}
	println(index, contains)
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

func (dictionaries Dictionaries) ToMap() map[string]string {
	result := make(map[string]string, len(dictionaries.Dictionary))
	for _, dict := range dictionaries.Dictionary {
		result[clean(dict.Error)] = dict.Name
	}
	return result
}

func (dictionaries Dictionaries) ContainsMap(prod string, dictionaryMap map[string]string) (original string, contains bool) {
	cleanedProd := clean(prod)
	if original, exists := dictionaryMap[cleanedProd]; exists {
		return original, true
	}
	return "", false
}

func clean(s string) string {
	return strings.ToLower(filter.ReplaceAllString(s, ""))
}
