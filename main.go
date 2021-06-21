package main

import (
	"fmt"
	"log"
	"os/exec"
)

func main() {
	fmt.Println("PDFに変換しています")
	err := exec.Command("powershell", "-NoProfile", "-ExecutionPolicy", "Unrestricted", ".\\excel2pdf.ps1", ".\\test.xlsx").Run()
	if err != nil {
		log.Fatal(err)
	}
}
