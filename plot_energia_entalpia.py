
import pandas as pd
import matplotlib.pyplot as plt
import os

def generate_energy_plots(file_path, sheet_name):
    try:
        # Carrega o arquivo Excel, definindo a segunda linha (índice 1) como o cabeçalho
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
    except FileNotFoundError:
        print(f"Erro: O arquivo {file_path} não foi encontrado.")
        return
    except Exception as e:
        print(f"Erro ao ler a planilha \'{sheet_name}\\' : {e}")
        return

    # Limpa os nomes das colunas (remove espaços em branco no início/fim)
    df.columns = df.columns.str.strip()

    # Garante que as colunas necessárias para o cálculo de Etocha existam
    required_cols_for_etocha = ["T", "U", "M"]
    if not all(col in df.columns for col in required_cols_for_etocha):
        print(f"Erro: As colunas {required_cols_for_etocha} (Temperatura, Entalpia, Vazão Mássica) não foram encontradas na planilha \'{sheet_name}\\'.)")
        return

    # Converte a coluna 'T' para tipo numérico, tratando possíveis erros
    df["T"] = pd.to_numeric(df["T"], errors="coerce")
    df.dropna(subset=["T"], inplace=True) # Remove linhas onde a temperatura não é um número

    # Encontra o valor de U na temperatura de 400K
    u_400_row = df[df["T"] == 400]
    if u_400_row.empty:
        print(f"Erro: Não foi possível encontrar a temperatura de 400K na coluna 'T' da planilha \'{sheet_name}\\'.) Certifique-se de que a coluna 'U' representa a Entalpia.")
        return
    u_400 = u_400_row["U"].iloc[0]

    # Calcula DeltaH e Etocha
    df["DeltaH"] = df["U"] - u_400
    eta_tocha = 0.95
    df["Etocha"] = (df["M"] * df["DeltaH"]) / eta_tocha

    # --- Gráfico 1: Etocha (RW) vs. Temperatura (K) ---
    plt.figure(figsize=(12, 7))
    plt.plot(df["Etocha"], df["T"], marker="o", linestyle="-")
    plt.title(f"Etocha (RW) vs. Temperatura (K) para {sheet_name}")
    plt.xlabel("Etocha (RW)")
    plt.ylabel("Temperatura (K)")
    plt.grid(True)
    output_filename_1 = f"etocha_vs_temp_{sheet_name}.png"
    plt.savefig(output_filename_1)
    plt.close()
    print(f"Gráfico \'{output_filename_1}\' gerado com sucesso.")


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
        generate_energy_plots(excel_file_path, selected_sheet)

    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")
