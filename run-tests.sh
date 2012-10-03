#!/bin/sh

cd "$(dirname "$0")"
ExcelSpreadsheetOps Testing.xlsm \
  RunVBACode[ \
    'SendMessageToListener "Passing tests: " & Range("NumTestsPassing").Value' \
    'SendMessageToListener "Failing tests: " & Range("NumTestsFailing").Value' \
  ]
