import sys
from cx_Freeze import setup, Executable

# Especifique o nome do arquivo principal da sua aplicação
target = Executable("appCA.py", base="Win32GUI")

# Lista de arquivos adicionais necessários (logo.webp, chromedriver.exe, icon.ico, config.py)
files = ['logo.webp', 'chromedriver.exe', 'icon.ico', 'config.py']

# Configurações do setup
setup(
    name="APP Correios Atende",
    version="2.0",
    description="Extração dos dados de postagem no Correios Atende",
    options={"build_exe": {"include_files": files}},
    executables=[target]
)
