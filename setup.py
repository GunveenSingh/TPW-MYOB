from cx_Freeze import setup, Executable

base = None

executables = [Executable("gunveenTPW.py", base=base)]

packages = ["pandas","time","csv","glob","os","re","pdfplumber","PyPDF2","io","reportlab.pdfgen","reportlab.lib.pagesizes","numpy","requests","warnings","socket","datetime","pysftp"]
options = {
    'build_exe': {
        'packages':packages,
    },
}

setup(
    name = "<first_ever>",
    options = options,
    version = "0.11",
    description = '',
    executables = executables
)
