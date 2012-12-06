#!/bin/sh

cd "$(dirname "$0")"

ExcelSpreadsheetOps Testing.xlsm \
  ImportAllModules . \
  RunMacro ReportTestResults
