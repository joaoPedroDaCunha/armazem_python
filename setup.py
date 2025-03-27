from cx_Freeze import setup, Executable

executables = [Executable("main.py", base="Win32GUI", icon="lago-de-dados.ico")]

setup(
    name="Armazem v1.4",
    version="0.1",
    description="Descrição do Meu Programa",
    executables=executables,
    options={
        "build_exe": {
            "include_files": ["print.py","programacao.py","salvar_dados.py","dados.xlsx"]
        }
    }
)