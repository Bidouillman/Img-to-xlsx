# Img-to-xlsx
Usage
Install the required Python libraries:
pip install openpyxl pillow

Clone the repository:
git clone https://github.com/your-username/Img-to-xlsx.git
cd Img-to-xlsx
Replace 'your-img-path' with the path to your image file and 'your-xlxs-name' with your desired Excel file name in the img_to_xlsx.py file.

Run the script:
python img_to_xlsx.py
Description
This script takes an image and converts it into an Excel spreadsheet. Each pixel is represented as three cells, one for the red (R), one for the green (G), and one for the blue (B) values. The script also colors each cell with the corresponding RGB color... I will change a few things to make it easier to use.

Requirements
Python 3.x
Pillow (PIL Fork): Image processing library
openpyxl: Excel file manipulation library
