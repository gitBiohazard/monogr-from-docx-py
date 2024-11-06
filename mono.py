from docx import Document
import zipfile
import os
from io import BytesIO
from PIL import Image # type: ignore
import xml.etree.ElementTree as ET
import re

class ABNTFormatter:
    def __init__(self, file_path, font="Times New Roman", font_size_title=12, 
                 font_size_subtitle=12, font_size_text=12, line_height=1.5):
        self.file_path = file_path
        self.font = font
        self.font_size_title = font_size_title
        self.font_size_subtitle = font_size_subtitle
        self.font_size_text = font_size_text
        self.line_height = line_height
        self.content = ""
        self.html_content = ""
        self.image_dir = "images"
        self.image_count = 1 

    def save_image(self, image_bytes, image_name):
        img = Image.open(BytesIO(image_bytes))
        image_path = os.path.join(self.image_dir, image_name)
        img.save(image_path)
        return image_path

    def extract_images(self):
        with zipfile.ZipFile(self.file_path, 'r') as docx_zip:
            image_files = [f for f in docx_zip.namelist() if f.startswith('word/media/')]
            
            if not os.path.exists(self.image_dir):
                os.makedirs(self.image_dir)

            image_paths = []
            for image_file in image_files:
                image_data = docx_zip.read(image_file)
                image_name = os.path.basename(image_file)
                image_path = self.save_image(image_data, image_name)
                image_paths.append(image_path)
            return image_paths

    def read_docx(self):
        doc = Document(self.file_path)
        content = ""
        is_indented = False
        
        image_paths = self.extract_images()

        for paragraph in doc.paragraphs:
            paragraph_text = paragraph.text.strip()

            for run in paragraph.runs:
                if 'graphic' in run._r.xml:
                    image_path = image_paths.pop(0)
                    figure_caption = f"<b>Figura {self.image_count}:</b> Nome da figura"
                    source_caption = f"Fonte: fonte da figura"
                    content += f'<div class="center">{figure_caption}</div>\n'
                    content += f'<div class="center"><img src="{image_path}" alt="Imagem" /></div>\n'
                    content += f'<div class="center">{source_caption}</div>\n'
                    self.image_count += 1
                    break

            if paragraph_text.lower() == "resumo":
                content += f"<div class='title center'>{paragraph_text}</div>\n"
            elif re.match(r'^\d{1,2}\s+[A-Za-zÀ-ÿ\s]+$', paragraph_text):
                content += f"<div class='title'>{paragraph_text}</div>\n"
                is_indented = False
            elif re.match(r'^\d{1,2}(\.\d{1,2})+\s+[A-Za-zÀ-ÿ\s]+$', paragraph_text):
                content += f"<div class='subtitle indented'>{paragraph_text}</div>\n"
                is_indented = True
            else:
                if is_indented:
                    content += f"<div class='text indented'>{paragraph_text}</div>\n"
                else:
                    content += f"<div class='text'>{paragraph_text}</div>\n"
        
        self.content = content.strip()
        return self.content
    
    def format_content_to_html(self):
        if not self.content:
            self.read_docx()
        
        self.html_content = f"""
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Documento ABNT</title>
            <style>
                body {{
                    font-family: '{self.font}', serif;
                    font-size: {self.font_size_text}px;
                    line-height: {self.line_height};
                    text-align: justify;
                    margin: 3cm 2cm 2cm 3cm;
                }}
                .title {{
                    font-size: {self.font_size_title}px;
                    font-weight: bold;
                    text-transform: uppercase;
                    text-align: left;
                    margin-top: 2cm;
                    margin-bottom: 1cm;
                }}
                .subtitle {{
                    font-size: {self.font_size_subtitle}px;
                    font-weight: bold;
                    text-transform: uppercase;
                    text-align: left;
                    margin-top: 1.5cm;
                    margin-bottom: 0.75cm;
                }}
                .text {{
                    font-size: {self.font_size_text}px;
                    text-indent: 1.25cm;
                    margin-bottom: 1.5em;
                }}
                .second-level {{
                    padding-left: 2.5cm;
                }}
                .indented {{
                    padding-left: 1.25cm;
                }}
                .center {{
                    text-align: center;
                }}
                img {{
                    max-width: 65%;
                    height: auto;
                    display: inline-block;
                    text-align: center;
                    margin: 1em 0;
                }}
            </style>
        </head>
        <body>
            {self.content.replace("\n", "")}
        </body>
        </html>
        """
        return self.html_content
    
    def save_html(self, output_path="output_abnt.html"):
        if not self.html_content:
            self.format_content_to_html()
        with open(output_path, "w", encoding="utf-8") as file:
            file.write(self.html_content)
        print(f"Documento formatado conforme ABNT salvo com sucesso como '{output_path}'.")
        
file_path = "doc.docx"  
formatter = ABNTFormatter(file_path)
formatter.read_docx()
formatter.format_content_to_html()
formatter.save_html("output_abnt.html")
