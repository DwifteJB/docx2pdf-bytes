package main

import (
	"log"
	"os"

	"github.com/DwifteJB/docx2pdf-bytes"
)

func main() {
	// read TestDocument.docx into bytes
	inputBytes, err := os.ReadFile("TestDocument.docx")

	if err != nil {
		log.Fatalf("Failed to read input file: %v", err)
	}

	outputBytes, err := docx2pdf.ConvertBytes(inputBytes)
	if err != nil {
		log.Fatalf("Conversion failed: %v", err)
	}

	// write outputBytes to TestDocument.pdf
	err = os.WriteFile("TestDocument.pdf", outputBytes, 0644)
	if err != nil {
		log.Fatalf("Failed to write output file: %v", err)
	}
}
