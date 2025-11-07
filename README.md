# 📘 EPUBtoPDFConverter

A lightweight, open-source Python application that converts **DRM-free EPUB files** into high-quality **PDF documents** using Google Chrome’s built-in “Save as PDF” functionality.

---

## 🚀 Features

- Converts **.epub** eBooks to **.pdf** automatically  
- Preserves **images, formatting, and CSS styles**  
- Handles EPUBs that have **nested directories or missing spine definitions**  
- Uses **Google Chrome (headless)** for consistent rendering  
- Temporary extraction handled automatically — no manual cleanup required  

---

## 🧰 Requirements

- **Python 3.8+**
- **Google Chrome** installed
- **Chrome WebDriver (chromedriver)** matching your Chrome version
- **Selenium** package (`pip install selenium`)

---

## 📂 How It Works

1. The script extracts the EPUB’s contents into a temporary directory.  
2. It parses the EPUB’s internal `container.xml` and `.opf` manifest to reconstruct reading order.  
3. It combines all `.xhtml` / `.html` content into a single temporary HTML file.  
4. It launches Chrome with a **“Save as PDF”** print profile and generates the final `.pdf`.  

---

## 🧾 Usage

```bash
python EPUBtoPDFConverter.py
