from cx_Freeze import setup, Executable

# Lista de módulos que você está usando no seu script
modules = ["pandas", "datetime", "tkinter", "tkcalendar"]
packages = ["encodings"]
include_files = ["ganttgenius_ico.ico", "config.ini"]
# Executável a ser criado
executables = [
    Executable(
        script='main.py',
        base='Win32GUI',  # Isso remove a janela do console no Windows
        icon='ganttgenius_ico.ico'
    )
]

# Chamada da função setup
setup(
    name='GanttGenius',
    version='3.5',
    description='Análise e geração de planilhas Gantt com base em dados de duas tabelas específicas',
    options={
        "build_exe": {
            "packages": modules + packages,
            "include_files": include_files,
            "include_msvcr": True,
        }
    },
    executables=executables
)