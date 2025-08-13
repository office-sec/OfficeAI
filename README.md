# 🚀 OfficeAI

> OfficeAI is an AI assistant plugin integrated into Microsoft Office and WPS, enabling seamless interaction with large language models for writing, polishing, translation, formatting, and data analysis in Word and Excel.  
> In Word, you can directly chat with AI in the right panel and insert results into documents with one click; in Excel, you only need natural language commands to automatically handle complex functions and chart generation.  
> **Currently offering free usage quota, lightweight setup, install and use immediately, significantly improving office productivity.**

[![Platform](https://img.shields.io/badge/Platform-Windows-blue)](https://office-ai.net) [![Free to Use](https://img.shields.io/badge/Price-Free-green)] 

**README Multi-language Version**: [English](/README_en.md) | [中文](/README_zh.md)

![OfficeAI Banner](screenshot/新增-1/offica-ai%20BANNER.png)
---

## 📚 Table of Contents

- [About OfficeAI](#about-officeai)  
- [Core Features](#core-features)  
  - [Word AI 📝](#word-ai-)  
  - [Excel AI 📊](#excel-ai-)  
- [Supported LLMs](#supported-llms)  
- [Installation](#installation)  
- [Usage Guide](#usage-guide)  
- [Resources](#resources)  
- [Development Roadmap](#development-roadmap)  
- [Contributing](#contributing)  
- [Contact Us](#contact-us)

---

## About OfficeAI

OfficeAI integrates seamlessly as a plugin into Microsoft Office and WPS, supporting intelligent document and spreadsheet processing. **Word AI** provides AI chat, writing, polishing, proofreading, intelligent formatting, and translation features, making document creation more efficient; **Excel AI** offers AI dialogue, formula generation, data analysis, and chart visualization features, making data processing more intelligent. You only need natural language to generate content, analyze data, and more—**no formulas, no templates required**.

---

## Core Features

### Word AI 📝
(Word AI is OfficeAI's highlight feature)

OfficeAI provides a comprehensive set of efficient and common AI features in Word/WPS plugins:

#### 💬 **AI Dialogue & Writing**
- **Natural AI Model Dialogue**: Engage in free multi-turn conversations with AI in Word/WPS. You can have multi-turn dialogues with popular large models. AI's response results can be exported to documents with one click or copied to clipboard
- **AI Writing**: You can customize prompts and save them as "Common Modes", call them with one click in any document to complete automatic continuation, summary extraction, language conversion and other tasks, supporting marketing copy, technical documentation, internal communication and various other types of article creation
- **AI Navigation**: Quickly activate ChatGPT, Claude, Gemini and other popular model dialogue entrances in Word/WPS, **no need to switch windows**, click and use immediately!

  **Dialogue & Writing Interface Demo**
  ![AI Writing Feature Demo](screenshot/新增-1/writing_compressed-1.gif)

#### 🎨 **Intelligent Formatting & Layout**
- **One-click Formatting**: Apply professional formatting with one click. Built-in multiple templates, support custom editing, users can perform efficient one-click formatting tasks, and can adjust effects in real-time through natural language dialogue, automatically generate outline directories. Users can also create custom templates and perform one-click formatting
- **Formatting Template Adjustment**: Support template copying, editing, customization, can set outline levels, Western fonts and other advanced formats

  **Intelligent Formatting & Layout Interface Demo**
  ![Intelligent Formatting Feature Demo](screenshot/新增-1/layout_compressed.gif)

#### ✍️ **Document Optimization & Processing**
- **Proofreading / Polishing**: Select text content (or entire document), Office AI can break through large model context length limitations, supporting long-text proofreading. AI automatically finds and corrects typos, grammar errors, etc.
- **AI Continuation**: Select text content, AI intelligently continues writing articles, results displayed in blue font, can confirm application or undo
- **AI Translation**: Support multi-language translation, built-in translation modes, can also customize translation modes

  **Document Optimization & Processing Interface Demo**
  ![Document Polishing Feature Demo](screenshot/Polish_compressed.gif)

#### 🛠️ **Toolbox Features**
- **AI Meeting Minutes**: Automatically generate professional meeting minutes format based on document content
- **AI Summary & Extraction**: Quickly extract key content from documents, saving reading time
- **Weekly Report Assistant**: Quickly organize weekly report key information into standard format
- **One-click Delete**: Delete extra blank lines, clear formatting, remove hyperlinks, etc.
- **Table Processing**: Insert content before tables, display headers across pages, select entire tables, etc.
- **Image to Text**: Recognize text in images and convert to editable text

More details: [WordAI Detailed Documentation](https://github.com/office-sec/OfficeAI/wiki/3-WordAI-Plugin)

---

## Excel AI 📊

OfficeAI provides a comprehensive set of efficient and common AI features in Excel/WPS plugins:

#### 💬 **AI Dialogue & Data Analysis**
- **Excel Dialogue**: Direct dialogue to generate data analysis
- **Intelligent Command Execution**: No formulas needed, easily handle complex instructions

  **AI Dialogue & Data Analysis Interface Demo**
  ![Excel AI Dialogue Feature Demo](screenshot/新增-1/Excel/chat%20to%20form.gif)

#### 📊 **Data Analysis & Visualization**
- **Data Visualization**: Automatically generate various charts and visualization views
- **Enterprise Privacy**: Local operation, ensuring data security
- **Data Cleaning**: One-click cleaning and optimization of table data

  **Data Visualization Feature Demo**
  ![Excel Data Visualization Demo](screenshot/新增-1/Excel/chart_compressed.gif)

#### 🎯 **Core Data Processing Features**
- **Formula Channel**: Natural language inquiry, AI automatically generates formulas
- **Cross-sheet Processing**: Cross-worksheet data processing and analysis
- **Batch Calculation**: Support addition, subtraction, multiplication, division and other batch operations
- **Data Extraction**: One-click extraction or filtering of various data
- **Intelligent Statistics**: Sum, average, maximum, minimum, etc.

  **Cross-sheet Processing Feature Demo**
  ![Excel AI Dialogue Feature Demo](screenshot/新增-1/Excel/chat%20to%20form.gif)

#### 🎨 **Format & Style Processing**
- **Number Format Conversion**: Chinese numbers, uppercase numbers, text format, etc.
- **Currency Format Processing**: Uppercase currency, ten thousand units, thousand separators
- **Cell Highlighting**: Intelligent highlighting, customizable effects
- **Quick Formatting**: Ten thousand units, currency symbols, Chinese numbers, etc.

  **Format & Style Processing Interface Demo**
  ![Excel Format Processing Feature Demo](screenshot/新增-1/Excel/format_compressed.gif)

#### 🛠️ **Toolbox Features**
- **Quick Input**: Gender, check/cross, yes/no, etc., support dropdown boxes
- **Rounding**: Round up, round down, round to even, etc.
- **Name Processing**: Modify, align, extract names
- **ID Card Processing**: Mask information, extract birth date, age, gender
- **Phone Number**: Mask, segment display, random generation
- **Table Translation**: Support multi-language table content translation, one-click language conversion

  **Translation Feature Demo**
  ![Excel Translation Feature Demo](screenshot/新增-1/Excel/trans_compressed.gif)

More details: [ExcelAI Detailed Documentation](https://github.com/office-sec/OfficeAI/wiki/4-ExcelAI-Plugin#432-extractfilter)

---

## Supported LLMs

Supports connection to various large language models, including:
- 🤖 Claude / Gemini / DeepSeek / OpenRouter 
Users can freely configure API Keys for flexible access.

---

## Installation

**System Requirements**: Windows 7/10/11 or later + Office 2013/2016/2019/Office 365

1. ⬇️ Download plugin installation package: [Install Now](https://office-ai.net/feature/wordai/)  
2. 📦 Extract and double-click to run the installer  
3. 🧑‍💻 Open Word or Excel or WPS, the plugin will load automatically  
4. ✨ Plugin entry is located in the top toolbar, click to start your intelligent office journey!

---

## Usage Guide

Please refer to the complete usage tutorial 👉 [https://github.com/office-sec/OfficeAI/wiki/3-WordAI-Plugin]

📌 **Tip**: The plugin's right-side floating window supports continuous dialogue and can embed local content context

---

## Resources

- 📄 [Product Website](https://office-ai.net)  
- 📚 [GitHub Wiki Detailed Documentation](https://github.com/office-sec/OfficeAI/wiki)
- 🔗 [Terms of Service & Privacy Policy](https://office-ai.net/policy)

---

## Development Roadmap

- 🔌 Support for macOS version  
- 📚 Long-text polishing, translation and other long-text interactions  
- 🎨 AI Intelligent Formatting: Automatically analyze document formats and generate custom templates
- 🧙‍♂️ Stronger context understanding capabilities  
- ✨ Support for more languages  
- 🤖 More convenient plugin API Key binding and operations  
- 📊 Excel feature enhancement and optimization

---

## Contributing

Welcome any form of contribution! Including but not limited to:

- 🐛 Bug reports  
- ✨ New feature suggestions  
- 🧑‍💻 PR submissions  
- 📖 Usage documentation supplements  

---

## Contact Us

- 🌐 Website: [https://office-ai.net](https://office-ai.net)    
- 📧 Email: service@office-ai.net
