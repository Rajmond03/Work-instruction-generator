Title:

    WI Generator

Description:

    Generates work instruction documents from a predefined template to speed up document creation.

    The program reads process data from Excel, matches steps with images, and uses the OpenAI API to generate quality requirements, safety considerations, and required tools. It also updates document metadata.

Input:

    Input file: 
        - data/WI-adat.xlsx
        - images/(supported formats: .jpg; .jpeg; .png; .JPG; .JPEG; .PNG)
    Required columns (case-sensitive):
        - Sorszám
        - Lépés
        - Képnév

Output:

    - output/ (generated .docx files)
    - File naming: GWP-HU-PR-WI-xxx_<document_name>.docx
    
Setup:

    Install dependencies:
        pip install -r requirements.txt

    Create a .env file and add your OpenAI API key:
        OPENAI_API_KEY=your_api_key_here

Run:

    Double-click START_WI.bat

    or run from terminal:
    START_WI.bat

Used technologies:

    - Python
    - pandas (data processing)
    - openpyxl (Excel handling)
    - Pillow (image processing)
    - docxtpl (Word template rendering)
    - python-docx (document formatting)
    - openai (AI content generation)
    - tkinter (GUI)
    - deep-translator (translation)