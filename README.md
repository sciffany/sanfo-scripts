# Scripts for Sanfo

This repository contains Google Apps scripts written to automate daily and weekly procedures for Sanfo International Trading Inc to improve workflow and reduce operation costs.

## Daily Purchase Updates

Script name: `dailyPurchaseUpdate.gs`

Script searches for information by date, selects relevant data column, and transfers data to target spreadsheet. It can be configured to run daily, or ran with UI to let users manually make changes to the file. Also contains functions to reset weekly data and automatically refresh date headers and sheet names based on dates.

## Request for Payment and Cheque automator

Script name: `requestForPayment.gs`

App collects information from several spreadsheets and compiles them into single Request for Payment file every week, and then reformats data to write cheques for each request, and finally compiles several cheques into PDF files ready for printing.

## Payroll.gs

Script name: `payroll.gs`

App presents information and calculates monthly employee salary by checking attendance and overtime hours against day of week and configurable list of public and special holidays.

## Compilation of COH

Script name: `coh.gs`

App collects purchase data from several spreadsheets and updates central spreadsheet


