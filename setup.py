from cx_Freeze import setup, Executable

# Substitua "meuscript.py" pelo nome do seu script Python
executables = [Executable("main.py",base="Win32GUI",icon="lago-de-dados.ico")]

setup(
    name="Armazem v1.4",
    version="0.1",
    description="Descrição do Meu Programa",
    executables=executables
)
