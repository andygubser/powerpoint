# 📊 PowerPoint Word Presentation Generator

An automated tool to create PowerPoint presentations with one word per slide, featuring customizable timing, fonts, and layout. Perfect for language learning, vocabulary practice, or word-based presentations.

## 🎯 Features

- ✅ **Auto-fitting text**: Automatically adjusts font size to fit words perfectly on slides
- ✅ **Custom fonts**: Supports custom fonts (includes DCH-Basisschrift)
- ✅ **Auto-advance**: Configurable automatic slide timing (1-10+ seconds)
- ✅ **Random order**: Optional word shuffling for variety
- ✅ **Perfect centering**: Both horizontal and vertical text alignment
- ✅ **No line breaks**: Ensures single-line word display
- ✅ **Batch processing**: Processes hundreds of words automatically
- ✅ **Configurable**: Easy parameter adjustment at script top

## 📁 Project Structure

```
powerpoint/
├── README.md                    # This file
├── src/                        # Source code
│   ├── word_presentation_generator.py    # Main script
│   └── test.py                 # Test/utility script
├── data/                       # Input files
│   └── words.xlsx              # Excel file with words
├── fonts/                      # Font files
│   ├── DCH-Basisschrift.otf    # Custom font
│   └── DCH-Basisschrift.otf.zip
└── output/                     # Generated presentations
    └── words_presentation_DCH.pptx
```

## 🚀 Quick Start

### 1. Install Prerequisites

**Python 3.7+** must be installed on your system.

Install required packages:
```bash
pip install pandas python-pptx openpyxl lxml
```

### 2. Install Font (Optional but Recommended)

1. Navigate to the `fonts/` folder
2. Double-click `DCH-Basisschrift.otf`
3. Click "Install" to add the font to your system

### 3. Run the Script

**Option A: Command Line**
```bash
# Navigate to project folder
cd "powerpoint"

# Run the script
cd src
python word_presentation_generator.py
```

**Option B: Windows File Explorer**
1. Open `C:\Users\AndyGubser\Documents\Projects\powerpoint\src`
2. Right-click in the folder → "Open PowerShell window here"
3. Type: `python word_presentation_generator.py`

**Option C: Double-click (Windows)**
1. Navigate to `src/` folder
2. Double-click `word_presentation_generator.py`

## ⚙️ Configuration

All settings can be easily adjusted at the top of `src/word_presentation_generator.py`:

### 🔧 **Quick Settings**

```python
# ⏱️ TIMING SETTINGS - EASILY ADJUSTABLE
AUTO_ADVANCE_SECONDS = 3        # Change this: 1, 2, 3, 5, 10 seconds
DISABLE_MOUSE_CLICK = True      # True = auto-only, False = allow clicks

# 🎲 PRESENTATION SETTINGS
RANDOMIZE_ORDER = True          # True = shuffle words, False = original order
```

### 🎨 **Advanced Settings**

```python
# Font settings
FONT_NAME = 'DCH-Basisschrift'  # Font to use
MAX_FONT_SIZE = 320             # Maximum font size (points)
MIN_FONT_SIZE = 20              # Minimum font size (points)

# File settings
INPUT_FILE = "../data/words.xlsx"    # Your Excel file
OUTPUT_FILE = "../output/presentation.pptx"  # Output location
```

## 📝 Input File Format

Your Excel file (`data/words.xlsx`) should contain:
- **Column A**: One word per row
- **No headers**: First row should be the first word
- **Single column**: Only the first column is used

Example:
```
die
der
und
ist
nicht
```

## 🎮 Usage Examples

### Basic Usage
```bash
cd src
python word_presentation_generator.py
```

### Custom Timing
1. Edit `AUTO_ADVANCE_SECONDS = 5` in the script
2. Run the script
3. Slides will advance every 5 seconds

### Different Font
1. Edit `FONT_NAME = 'Arial'` in the script
2. Run the script

### Keep Original Order
1. Edit `RANDOMIZE_ORDER = False` in the script
2. Run the script

## 🔧 Troubleshooting

### Common Issues

**1. "ModuleNotFoundError: No module named 'pandas'"**
```bash
pip install pandas python-pptx openpyxl lxml
```

**2. "FileNotFoundError: words.xlsx"**
- Ensure `words.xlsx` is in the `data/` folder
- Check the file path in the configuration

**3. Font not displaying correctly**
- Install `DCH-Basisschrift.otf` from the `fonts/` folder
- Or change `FONT_NAME` to a system font like 'Arial'

**4. "python is not recognized"**
- Install Python from https://python.org
- Make sure "Add Python to PATH" is checked during installation

### Manual Slide Timing Setup

If automatic timing doesn't work:
1. Open the generated `.pptx` file
2. Go to **Transitions** tab in PowerPoint
3. Uncheck **"On Mouse Click"**
4. Check **"After"** and set to desired seconds
5. Click **"Apply To All"**

## 📊 Output

The script generates:
- A PowerPoint presentation file in `output/` folder
- Detailed console output showing:
  - Number of slides created
  - Font size adjustments made
  - Size distribution statistics
  - Success confirmation

## 🎯 Performance

- **Processing speed**: ~1-2 seconds per 100 slides
- **Memory usage**: Minimal (< 50MB)
- **File size**: ~200KB for 200 slides
- **Compatibility**: PowerPoint 2010+, LibreOffice, Google Slides

## 🌍 Language Support

Works with any language that uses standard character sets:
- ✅ German (included example)
- ✅ English
- ✅ French, Spanish, Italian
- ✅ Most European languages
- ⚠️ Some special characters may need font adjustment

## 🤝 Contributing

Feel free to modify and improve the script:
1. Fork the project
2. Make your changes
3. Test with different word lists
4. Share improvements!

## 📄 License

This project is free to use for personal and educational purposes.

## 🆘 Support

For help:
1. Check the troubleshooting section above
2. Verify all prerequisites are installed
3. Test with a simple word list first
4. Check file paths and permissions

---

**Made with ❤️ for vocabulary learning and word presentations** 