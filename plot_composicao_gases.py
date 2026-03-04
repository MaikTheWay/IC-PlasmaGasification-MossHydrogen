
import pandas as pd
import matplotlib.pyplot as plt
import os

def generate_composition_plots(file_path, sheet_name):
    try:
        # Carrega o arquivo Excel, definindo a segunda linha (índice 1) como o cabeçalho
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
    except FileNotFoundError:
        print(f"Erro: O arquivo {file_path} não foi encontrado.")
        return
    except Exception as e:
        print(f"Erro ao ler a planilha \'{sheet_name}\': {e}")
        return

    # Limpa os nomes das colunas (remove espaços em branco no início/fim)
    df.columns = df.columns.str.strip()

    # --- Gráfico 1: Composição (%WT) vs. Temperatura (K) para CO, H2, C(c) ---
    composition_wt_cols = ["CO", "H2", "C(c)"]
    if all(col in df.columns for col in composition_wt_cols):
        plt.figure(figsize=(12, 7))
        for col in composition_wt_cols:
            plt.plot(df[col], df["T"], marker=".", linestyle="-", label=col)
        plt.title(f"Composição (%WT) vs. Temperatura (K) para {sheet_name}")
        plt.xlabel("Composição (%WT)")
        plt.ylabel("Temperatura (K)")
        plt.legend()
        plt.grid(True)
        output_filename_1 = f"composicao_wt_vs_temp_{sheet_name}.png"
        plt.savefig(output_filename_1)
        plt.close()
        print(f"Gráfico \'{output_filename_1}\' gerado com sucesso.")
    else:
        print(f"Aviso: Colunas de composição (%WT) não encontradas. O primeiro gráfico não será gerado.")

    # --- Gráfico 2: Composição (%Vol) de H2 e CO vs. Temperatura (K) ---
    composition_vol_cols = ["H2", "CO"]
    if all(col in df.columns for col in composition_vol_cols):
        plt.figure(figsize=(12, 7))
        for col in composition_vol_cols:
            # Multiplica por 100 para obter a porcentagem em volume
            plt.plot(df[col] * 100, df["T"], marker=".", linestyle="-", label=f"{col} (%Vol)")
        plt.title(f"Composição (%Vol) de H2 e CO vs. Temperatura (K) para {sheet_name}")
        plt.xlabel("Composição (%Vol)")
        plt.ylabel("Temperatura (K)")
        plt.legend()
        plt.grid(True)
        output_filename_2 = f"composicao_vol_h2_co_vs_temp_{sheet_name}.png"
        plt.savefig(output_filename_2)
        plt.close()
        print(f"Gráfico \'{output_filename_2}\' gerado com sucesso.")
    else:
        print(f"Aviso: Colunas de H2 e/ou CO para composição (%Vol) não encontradas. O segundo gráfico não será gerado.")


if __name__ == "__main__":
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        excel_files = [f for f in os.listdir(script_dir) if f.endswith(".xlsx")]

        if not excel_files:
            print("Erro: Nenhum arquivo .xlsx encontrado na pasta do script.")
            exit()

        print("Arquivos Excel encontrados:")
        for i, fname in enumerate(excel_files):
            print(f"  {i+1}. {fname}")

        if len(excel_files) == 1:
            selected_excel_file = excel_files[0]
        else:
            while True:
                try:
                    choice = int(input("Escolha o número do arquivo Excel que deseja usar: "))
                    if 1 <= choice <= len(excel_files):
                        selected_excel_file = excel_files[choice - 1]
                        break
                    else:
                        print("Escolha inválida.")
                except ValueError:
                    print("Entrada inválida. Por favor, insira um número.")
        
        excel_file_path = os.path.join(script_dir, selected_excel_file)

        xls = pd.ExcelFile(excel_file_path)
        available_sheets = xls.sheet_names

        print("\nPlanilhas disponíveis:")
        for i, sheet in enumerate(available_sheets):
            print(f"  {i+1}. {sheet}")

        while True:
            try:
                choice = int(input("Escolha o número da planilha para análise: "))
                if 1 <= choice <= len(available_sheets):
                    selected_sheet = available_sheets[choice - 1]
                    break
                else:
                    print("Escolha inválida.")
            except ValueError:
                print("Entrada inválida. Por favor, insira um número.")

        print(f"\nProcessando a planilha: \'{selected_sheet}\'...")
        generate_composition_plots(excel_file_path, selected_sheet)

    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")
