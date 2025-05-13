# SlideShare Downloader

## Overview

SlideShare Downloader is a Python-based GUI application that allows users to download SlideShare presentations as either PDF or PPT files. The application extracts images from a SlideShare URL and converts them into the chosen format, enabling offline access to SlideShare content.

## Features

* Download SlideShare presentations as **PDF** or **PPT**.
* Simple and intuitive GUI built using **Tkinter** and **ttkbootstrap**.
* Displays download progress using a circular progress bar.
* Open the download folder or the downloaded file directly from the application.
* Clear the input fields and reset the application state.

## Requirements

* Python 3.x
* ttkbootstrap
* BeautifulSoup4
* requests
* img2pdf
* python-pptx

Install the dependencies using:

```bash
pip install ttkbootstrap beautifulsoup4 requests img2pdf python-pptx
```

## Usage

1. **Run the application:**

   ```bash
   python slideshare_downloader.py
   ```
2. **Enter the SlideShare URL.**
3. **Select the desired output format (PDF or PPT).**
4. **Click the Download button.**
5. The download progress is displayed, and upon completion, you can open the folder or the downloaded file.

## Project Structure

* `slideshare_downloader.py` - Main application file.
* `requirements.txt` - Dependencies list.

## Error Handling

* Invalid URLs display an error message.
* Network or download errors are handled and displayed to the user.

## Author

* **Kunal Saraswat**
* Project No: 5
* Version: FINAL

## License

This project is open-source and available under the MIT License.
