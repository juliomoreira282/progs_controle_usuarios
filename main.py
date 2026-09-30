import ttkbootstrap as tb
from backend import GerenciadorExcel
from interface import InterfaceGestaoComputadores

if __name__ == "__main__":
    arquivo_excel = "controle_computadores_e_usuarios.xlsx"
    backend = GerenciadorExcel(arquivo_excel)
    root = tb.Window(themename="superhero")
    app = InterfaceGestaoComputadores(root, backend)
    root.mainloop()