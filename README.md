# WhatsApp Auto Sender 📊

Automates sending Excel range screenshots to WhatsApp contacts.

## Features
- Copies a specific Excel range as an image
- Opens WhatsApp Desktop for each contact
- Pastes and sends the image automatically

## Requirements
- Windows OS
- WhatsApp Desktop installed
- Python 3.x

## Installation
pip install pywin32 pyautogui

## Usage
1. Update `RECIPIENTS` list with phone numbers and cell ranges
2. Update `EXCEL_PATH` with your file path
3. Run: python report_auto.py

## Note
Make sure WhatsApp Desktop is installed and logged in before running.