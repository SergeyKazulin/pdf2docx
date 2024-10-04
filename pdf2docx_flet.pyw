import flet as ft
from pdf2docx import Converter
import os

def main(page: ft.Page):
    page.title = "PDF to DOCX Converter"
    page.window_width = 600
    page.window_height = 450
    page.theme_mode = "dark"#"light"

    pdf_path = ft.Text()
    docx_folder = ft.Text()
    result_text = ft.Text()

    def select_pdf(e: ft.FilePickerResultEvent):
        if e.files:
            pdf_path.value = e.files[0].path
            page.update()

    def select_folder(e: ft.FilePickerResultEvent):
        if e.path:
            docx_folder.value = e.path
            page.update()

    def convert_pdf_to_docx(e):
        if not pdf_path.value or not docx_folder.value:
            result_text.value = "Please select both PDF file and output folder"
            result_text.color = "red"
            page.update()
            return

        try:
            pdf_file = pdf_path.value
            docx_file = os.path.join(docx_folder.value, os.path.splitext(os.path.basename(pdf_file))[0] + ".docx")
            
            cv = Converter(pdf_file)
            cv.convert(docx_file)
            cv.close()

            result_text.value = f"Conversion successful! File created:\n{docx_file}"
            result_text.color = "lightgreen"

        except Exception as ex:
            result_text.value = f"Error during conversion: {str(ex)}"
            result_text.color = "red"
        finally:
            page.update()

    pdf_picker = ft.FilePicker(on_result=select_pdf)
    folder_picker = ft.FilePicker(on_result=select_folder)

    page.overlay.extend([pdf_picker, folder_picker])

    converted_button = ft.ElevatedButton("Convert PDF to DOCX", color="#FAEEDD", on_click=convert_pdf_to_docx, bgcolor="#003153")
    converted_button_centred = ft.Container(converted_button, alignment=ft.alignment.center)

    page.add(
        ft.Text("PDF to DOCX Converter", size=24, weight="bold", color='lightblue', width=600, text_align='center'),
        ft.Text(value=f"{'*'*107}", size=12, color='lightblue'),
        ft.Row([
            ft.Text("Selected PDF File: ", color="#9ACEEB"),
            pdf_path
        ]),
        ft.ElevatedButton("Select PDF", color="#FAEEDD", on_click=lambda _: pdf_picker.pick_files(allow_multiple=False), bgcolor="blue"),
        #ft.Text(value="", size=8),
        ft.Divider(height=10),#, color="transparent"),   # выделение на фоне
        ft.Row([
            ft.Text("Selected Output Folder: ", color="#9ACEEB"),
            docx_folder
        ]),
        ft.ElevatedButton("Select Output Folder", color="#FAEEDD", on_click=lambda _: folder_picker.get_directory_path(), bgcolor="blue"),
        #ft.Text(value='', size=10),
        ft.Divider(height=20),#, color="transparent"),
        converted_button_centred,
        #ft.Text(value='', size=8),
        ft.Divider(height=10, color="transparent"),   # не видно на фоне
        result_text
    )

ft.app(target=main)
