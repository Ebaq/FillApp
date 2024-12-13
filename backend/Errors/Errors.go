package Errors

import "encoding/json"

type ProgramError struct {
	StatusCode string
	Block      string
	Message    string
}

func NewProgramError(statusCode string, block string, message string) string {
	err, _ := json.Marshal(&ProgramError{
		StatusCode: statusCode,
		Block:      block,
		Message:    message,
	})

	return string(err)
}
