import sys
from cx_Freeze import setup, Executable


build_exe_options = {
    "packages": ["pandas", "openpyxl"],
    "include_files": ["excel_finder_icon.ico"]
}

setup(
    name="Excel Finder",
    version="1.0",
    description="Aplicação para localizar palavras em um arquivo .xlsx e gerar um novo arquivo com os resultados da pesquisa.",
    options={"build_exe": build_exe_options},
    executables = [Executable("app.py", icon="excel_finder_icon.ico")]
)
