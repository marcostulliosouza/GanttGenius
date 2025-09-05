from cx_Freeze import setup, Executable

modules = ["pandas", "datetime", "tkinter", "tkcalendar"]
packages = ["encodings"]
include_files = ["ganttgenius_ico.ico", "config.ini"]
executables = [
    Executable(
        script='main.py',
        base='Win32GUI',
        icon='ganttgenius_ico.ico'
    )
]
setup(
    name='GanttGenius',
    version='3.5.2',
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