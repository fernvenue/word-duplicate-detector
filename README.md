# Word Duplicate Content Detection Tool

[English](README.md) | [简体中文](README.zh_Hans.md) | [繁體中文](README.zh_Hant.md)

An open-source tool for detecting duplicate content in Word documents and adding intelligent annotations, with intelligent grouping and precise paragraph positioning features.

## ✨ Features

- **🔍 Automatic Detection**: Intelligently detect duplicate content in Word documents;
- **📝 Smart Annotations**: Add annotations to duplicate content, showing source paragraph numbers;
- **🔢 Intelligent Grouping**: Same content uses the same number for easy identification;
- **🎯 Precise Positioning**: Accurate paragraph positioning without page number estimation errors;
- **⚙️ Customizable Parameters**: Support for custom detection parameters (default 30 characters);
- **📄 Non-destructive**: Generate new annotated document while preserving original;
- **💻 Portable**: Portable operation, no installation required, no additional dependencies;
- **🌐 Multi-language Support**: Supports English, Simplified Chinese, Traditional Chinese;

## 🎮 Usage

Download the latest version directly from [Releases](https://github.com/fernvenue/word-duplicate-detector/releases/latest) and run, no installation required.

![image](./assets/images/3B96705F8401E35A.webp)

1. **Select Document**: Choose a Word document (.docx format) to analyze;
2. **Configure Parameters**: Set minimum duplicate length (recommended: 20-50 characters);
3. **Start Detection**: Click the *Start Detection* button;
4. **Review Results**: Check the generated annotated document;
5. **Navigate Duplicates**: Use duplicate numbers to quickly identify groups;

### 📊 Annotation Format

The tool adds annotations in the following format:

```
Duplicate #1: Repeats content from paragraph 25
```

- **Duplicate #N**: All duplicates of the same content use the same number;
- **Paragraph Y**: Exact paragraph number of the duplicate source;

### 🧠 Duplicate Numbering Logic

Same duplicate content gets assigned the same number, e.g.: If paragraph A repeats paragraph B, and paragraph C also repeats paragraph B, then both A and C will show the same duplicate number, making it easy to identify which paragraphs belong to the same duplicate group.

## 🚀 Build from Source

Python environment required.

Clone the repository:

```bash
git clone https://github.com/fernvenue/word-duplicate-detector.git
cd word-duplicate-detector
```

Run build script:

```bash
build.bat
```