package main

import (
	"bytes"
	"context"
	"encoding/base64"
	"errors"
	"fmt"
	"github.com/aws/aws-lambda-go/events"
	"github.com/aws/aws-lambda-go/lambda"
	"github.com/grokify/go-awslambda"
	split_to_cell "github.com/mhall58/excel-cell-rows/pkg/split-to-cell"
	"github.com/xuri/excelize/v2"
	"io"
	"log"
	"strings"
)

func main() {
	lambda.Start(handleRequest)
}

type customStruct struct {
	Content       string
	FileName      string
	FileExtension string
}

func handleRequest(ctx context.Context, req events.APIGatewayProxyRequest) (events.APIGatewayProxyResponse, error) {
	res := events.APIGatewayProxyResponse{}
	log.Println("This is matt")

	r, err := awslambda.NewReaderMultipart(req)
	if err != nil {
		log.Println("can't read request")
		return res, err
	}

	var (
		delimiter   string
		data        []byte
		filename    string
		splitColumn string
		checkColumn string
	)

	// Parse Form
	for {
		part, err := r.NextPart()

		if err == io.EOF {
			break
		}

		if err != nil {
			return res, err
		}

		content, err := io.ReadAll(part)
		if err != nil {
			return res, err
		}

		if part.FormName() == "upload" {
			data = content
			filename = part.FileName()
		}

		if part.FormName() == "delimiter" {
			delimiter = string(content)
			if delimiter == "nl" {
				delimiter = "\n"
			}
		}

		if part.FormName() == "split_column" {
			splitColumn = string(content)
		}

		if part.FormName() == "check_column" {
			checkColumn = string(content)
		}
	}

	file, err := excelize.OpenReader(bytes.NewReader(data))

	if err != nil {
		return res, err
	}

	if !strings.HasSuffix(filename, ".xlsx") {
		return res, errors.New("not an xlsx file")
	}

	split_to_cell.SplitCells(file, checkColumn, splitColumn, delimiter)

	buffer, err := file.WriteToBuffer()
	if err != nil {
		return res, err
	}

	filename = strings.ReplaceAll(filename, ".xlsx", "-split.xlsx")

	res = events.APIGatewayProxyResponse{
		StatusCode: 200,
		Headers: map[string]string{
			"Content-Type":        "application/octet-stream",
			"Content-Disposition": fmt.Sprintf("attachment; filename=\"%s\"", filename),
		},
		Body:            base64.StdEncoding.EncodeToString(buffer.Bytes()),
		IsBase64Encoded: true,
	}
	return res, nil
}
