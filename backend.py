import xlwings as xw
import shutil
import os
from datetime import datetime

class GerenciadorExcel:
    def __init__(self, arquivo):
        self.arquivo = arquivo

    def criar_backup(self):
        if not os.path.exists('backups_usuarios'):
            os.makedirs('backups_usuarios')
        timestamp = datetime.now().strftime("%Y%m%d_%H%m%S")
        caminho_backup = f"backups_usuarios/backup_{timestamp}.xlsx"
        try:
            shutil.copy(self.arquivo, caminho_backup)
        except Exception:
            pass

    def travar_planilha(self):
        try:
            with open(self.arquivo, 'r+'):
                pass
            return True
        except PermissionError:
            return False
        except FileNotFoundError:
            return True

    def carregar_dados(self):
        marcas_validas = []
        patrimonios_existentes = set()
        mapa_linhas = {}
        dados_memoria = {}

        app = xw.App(visible=False)
        try:
            wb = app.books.open(self.arquivo)
            sheet = wb.sheets['Dados Gerais']
            sheet_marcas = wb.sheets['Validação de Marcas']

            marcas = sheet_marcas.range('A2:A100').value
            if marcas:
                for m in marcas:
                    if m is not None:
                        marcas_validas.append(str(m).strip())

            last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
            if last_row >= 2:
                matriz = sheet.range(f'A2:G{last_row}').value
                if last_row == 2:
                    matriz = [matriz]

                for idx, linha_dados in enumerate(matriz):
                    linha_num = idx + 2
                    valor = linha_dados[0]

                    if valor is not None:
                        try:
                            patrimonio = str(int(float(valor))).strip()
                        except (ValueError, TypeError):
                            patrimonio = str(valor).strip()

                        patrimonios_existentes.add(patrimonio)
                        mapa_linhas[patrimonio] = linha_num
                        dados_memoria[patrimonio] = [
                            linha_dados[1], linha_dados[2], linha_dados[3], 
                            linha_dados[4], linha_dados[5], linha_dados[6]
                        ]
            wb.close()
            return True, marcas_validas, patrimonios_existentes, mapa_linhas, dados_memoria, ""
        except Exception as e:
            return False, [], [], set(), {}, {}, str(e)
        finally:
            app.quit()

    def salvar_adicoes(self, novos_itens):
        self.criar_backup()
        app = xw.App(visible=False)
        try:
            wb = app.books.open(self.arquivo)
            sheet = wb.sheets['Dados Gerais']

            last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
            linha_atual = last_row + 1 if last_row >= 1 else 2

            for item in novos_itens:
                c1 = sheet.range(f'A{linha_atual}')
                c1.number_format = '@'
                c1.value = item[0]

                sheet.range(f'B{linha_atual}').value = item[1]
                sheet.range(f'C{linha_atual}').value = item[2]
                sheet.range(f'D{linha_atual}').value = item[3]
                sheet.range(f'E{linha_atual}').value = item[4]
                sheet.range(f'F{linha_atual}').value = item[5]
                sheet.range(f'G{linha_atual}').value = item[6]

                linha_ref = linha_atual - 2 if (linha_atual - 2) > 1 else linha_atual - 1
                if linha_ref >= 2:
                    origem = sheet.range(f'A{linha_ref}:G{linha_ref}')
                    destino = sheet.range(f'A{linha_atual}:G{linha_atual}')
                    origem.copy()
                    destino.paste('formats')

                linha_atual += 1

            try:
                if sheet.api.ListObjects.Count > 0:
                    table = sheet.api.ListObjects(1)
                    table.Resize(sheet.range(f'A1:G{linha_atual - 1}').api)
            except Exception:
                pass

            wb.save()
            wb.close()
            return True, ""
        except Exception as e:
            return False, str(e)
        finally:
            app.quit()

    def salvar_substituicoes(self, itens_substituir, mapa_linhas):
        self.criar_backup()
        app = xw.App(visible=False)
        try:
            wb = app.books.open(self.arquivo)
            sheet = wb.sheets['Dados Gerais']

            for pat, novos_dados in itens_substituir, mapa_linhas:
                linha = mapa_linhas[pat]
                sheet.range(f'B{linha}').value = novos_dados[0]
                sheet.range(f'C{linha}').value = novos_dados[1]
                sheet.range(f'D{linha}').value = novos_dados[2]
                sheet.range(f'E{linha}').value = novos_dados[3]
                sheet.range(f'F{linha}').value = novos_dados[4]
                sheet.range(f'G{linha}').value = novos_dados[5]

            wb.save()
            wb.close()
            return True, ""
        except Exception as e :
            return False, str(e)
        finally:
            app.quit()

    def salvar_remocoes(self, itens_remover, mapa_linhas):
        self.criar_backup()
        app = xw.App(visible=False)
        try:
            wb = app.books.open(self.arquivo)
            sheet = wb.sheets['Dados Gerais']

            linhas_para_deletar = [mapa_linhas[pat] for pat in itens_remover]
            linhas_para_deletar.sort(reverse=True)

            for linha in linhas_para_deletar:
                sheet.range(f'{linha}:{linha}').delete()

            wb.save()
            wb.close()
            return True, ""
        except Exception as e:
            return False, str(e)
        finally:
            app.quit()