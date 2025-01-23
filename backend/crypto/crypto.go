package crypto

import (
	"crypto/sha256"
	"encoding/hex"
	"fillappgo/backend/Errors"
	"fillappgo/backend/consts"
	"fmt"
)

func hashPINWithSalt(pin string) string {
	hasher := sha256.New()
	hasher.Write([]byte(pin + "very_secure_salt22"))
	return hex.EncodeToString(hasher.Sum(nil))
}

func ComparePins(pin string) error {
	if hashPINWithSalt(pin) != consts.PIN {
		return fmt.Errorf(Errors.NewProgramError("0x1", "Crypto", "Неверный пин!"))
	}

	return nil
}
