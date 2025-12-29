# DD2 Auto KeyPresser - Makefile
# Usage: make [target]

.PHONY: run build clean install dev help

# Default target
help:
	@echo DD2 Auto KeyPresser - Available commands:
	@echo.
	@echo   make install  - Install dependencies
	@echo   make dev      - Install dev dependencies
	@echo   make run      - Run the application
	@echo   make build    - Build exe file
	@echo   make clean    - Clean build artifacts
	@echo   make help     - Show this help

# Install production dependencies
install:
	pip install -r requirements.txt

# Install development dependencies
dev:
	pip install -r requirements.txt
	pip install pyinstaller

# Run the application
run:
	python dd2-keypresser.py

# Build exe file
build:
	pyinstaller dd2-keypresser.spec --clean --noconfirm
	@echo Build complete! Exe file: dist/dd2-keypresser.exe

# Clean build artifacts
clean:
	@if exist build rmdir /s /q build
	@if exist dist rmdir /s /q dist
	@if exist __pycache__ rmdir /s /q __pycache__
	@echo Cleaned build artifacts
