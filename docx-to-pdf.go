package docx2pdf

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/jung-kurt/gofpdf"
)

type Text struct {
	Text string `xml:",chardata"`
}

type RunProperties struct {
	Bold      bool   `xml:"b"`
	Italic    bool   `xml:"i"`
	FontSize  string `xml:"sz"`
	FontColor string `xml:"color"`
}

type Run struct {
	Properties RunProperties `xml:"rPr"`
	Texts      []Text        `xml:"t"`
}

type Paragraph struct {
	Alignment string `xml:"pPr>jc"` // Align: left, right, center
	Runs      []Run  `xml:"r"`
}

type TableCell struct {
	Text string `xml:"p>r>t"`
}

type TableRow struct {
	Cells []TableCell `xml:"tc"`
}

type Table struct {
	Rows []TableRow `xml:"tr"`
}

type Drawing struct {
	Image ImageRef `xml:"inline>graphic>graphicData>pic:pic>blipFill>blip"`
}

type ImageRef struct {
	ID string `xml:"embed,attr"`
}

type Document struct {
	Paragraphs []Paragraph `xml:"body>p"`
	Tables     []Table     `xml:"body>tbl"`
	Drawings   []Drawing   `xml:"body>drawing"`
}

func extractTextFromDocx(docxBytes []byte) (string, error) {
	reader, err := zip.NewReader(bytes.NewReader(docxBytes), int64(len(docxBytes)))
	if err != nil {
		return "", err
	}

	var documentXML string
	for _, file := range reader.File {
		if file.Name == "word/document.xml" {
			rc, err := file.Open()
			if err != nil {
				return "", err
			}
			defer rc.Close()

			bytes, err := io.ReadAll(rc)
			if err != nil {
				return "", err
			}
			documentXML = string(bytes)
			break
		}
	}

	return documentXML, nil
}

func setFontFromRun(pdf *gofpdf.Fpdf, run Run) {
	fontStyle := ""
	if run.Properties.Bold {
		fontStyle += "B"
	}
	if run.Properties.Italic {
		fontStyle += "I"
	}
	fontSize := 12.0 // default font size
	if run.Properties.FontSize != "" {
		// convert font size from half-points to points
		fontSizeValue, err := strconv.ParseFloat(run.Properties.FontSize, 64)
		if err == nil {
			fontSize = fontSizeValue / 2
		}
	}
	pdf.SetFont("Arial", fontStyle, fontSize)
	if run.Properties.FontColor != "" {
		pdf.SetTextColor(parseHexColor(run.Properties.FontColor))
	}
}

func parseHexColor(s string) (int, int, int) {
	var r, g, b int
	fmt.Sscanf(s, "%02x%02x%02x", &r, &g, &b)
	return r, g, b
}

func setParagraphAlignment(_ *gofpdf.Fpdf, alignment string) string {
	switch alignment {
	case "center":
		return "C"
	case "right":
		return "R"
	default:
		return "L"
	}
}

func processParagraph(pdf *gofpdf.Fpdf, para Paragraph) {
	align := setParagraphAlignment(pdf, para.Alignment)
	pdf.SetFont("Arial", "", 12)

	for _, run := range para.Runs {
		setFontFromRun(pdf, run)
		for _, text := range run.Texts {
			pdf.CellFormat(0, 6, text.Text, "", 1, align, false, 0, "")
		}
	}

	pdf.Ln(4) // Spasi antar paragraf
}

func processTable(pdf *gofpdf.Fpdf, table Table) {
	pdf.SetFont("Arial", "", 10)

	for _, row := range table.Rows {
		for _, cell := range row.Cells {
			pdf.CellFormat(40, 10, cell.Text, "1", 0, "C", false, 0, "")
		}
		pdf.Ln(-1) // Pindah ke baris berikutnya
	}
}

func addImageToPDF(pdf *gofpdf.Fpdf, imgPath string, x, y, width, height float64) {
	pdf.Image(imgPath, x, y, width, height, false, "", 0, "")
}

func extractImagesFromDocx(_ []byte, reader *zip.Reader) (map[string]string, error) {
	images := make(map[string]string)
	tempDir, err := os.MkdirTemp("", "docx_images")
	if err != nil {
		return nil, err
	}

	for _, file := range reader.File {
		if strings.HasPrefix(file.Name, "word/media/") {
			rc, err := file.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()

			imgBytes, err := io.ReadAll(rc)
			if err != nil {
				return nil, err
			}

			imgPath := filepath.Join(tempDir, filepath.Base(file.Name))
			err = os.WriteFile(imgPath, imgBytes, 0644)
			if err != nil {
				return nil, err
			}

			images[file.Name] = imgPath
		}
	}
	return images, nil
}

func createPDF(text string, images map[string]string) ([]byte, error) {
	pdf := gofpdf.New("P", "mm", "A4", "")
	pdf.AddPage()

	var doc Document
	err := xml.Unmarshal([]byte(text), &doc)
	if err != nil {
		return nil, err
	}

	for _, para := range doc.Paragraphs {
		processParagraph(pdf, para)
	}

	for _, table := range doc.Tables {
		processTable(pdf, table)
	}

	for _, drawing := range doc.Drawings {
		imgPath, exists := images["word/media/"+drawing.Image.ID]
		if exists {
			addImageToPDF(pdf, imgPath, 10, 10, 50, 50) // Example coordinates and size
		}
	}

	// create buffer to write pdf to
	var buf bytes.Buffer
	err = pdf.Output(&buf)
	if err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func ConvertBytes(inputBytes []byte) ([]byte, error) {
	reader, err := zip.NewReader(bytes.NewReader(inputBytes), int64(len(inputBytes)))
	if err != nil {
		return nil, err
	}

	text, err := extractTextFromDocx(inputBytes)
	if err != nil {
		return nil, err
	}

	images, err := extractImagesFromDocx(inputBytes, reader)
	if err != nil {
		return nil, err
	}

	pdfBytes, err := createPDF(text, images)
	if err != nil {
		return nil, err
	}

	return pdfBytes, nil
}
