#!/usr/bin/env bash
go build -o ./bin/Excel2Resource main.go
GOOS=windows GOARCH=amd64 go build -o ./bin/Excel2Resource.exe main.go
