import os.path
import pandas as pd


def read_files():
    amount_file = 0
    dict_files = {}
    name_dir_now = str(input('Qual endereço da pasta:'))
    if os.path.isdir(name_dir_now):
        # lista arquivos
        for name_file in os.listdir(name_dir_now):
            name_file_now = os.path.join(name_dir_now, name_file)
            # verifica se é arquivo e se tem extensão correta
            if os.path.isfile(name_file_now)\
                    and name_file.endswith(('.xls', '.xlsx')):
                amount_file += 1
                print(name_file_now)
                # carrega arquivo excel
                bd_sheet = pd.ExcelFile(name_file_now)
                # encontra planilha correta
                sheet_valid = [
                    value for value in bd_sheet.sheet_names
                    if value in ['Plan1', 'Planilha1']
                ]
                # cria dataframe
                df = pd.read_excel(name_file_now, sheet_valid[0])
                # cria coluna com nome do arquivo
                df['ARQUIVO'] = name_file
                if not df.empty:
                    # dados adicionado se já existe cidade
                    if name_file in dict_files:
                        dict_files[name_file] = dict_files[name_file].append(
                            df, ignore_index=True
                        )
                    # cria novo item no dicionário, cidade nova
                    else:
                        dict_files[name_file] = df
            else:
                print('não é arquivo ou não tem extensão correta')
        # SE PRECISAR criar dicionário para ordenar colunas
        # exemplo: data = {'PROBLEMAS': [1], 'EDUCAÇÃO': [1]}
        # criando dataframe VAZIO se precisar adicionar data como argumento
        df_final = pd.DataFrame()

        # percorre dicionario de dados, adicionando novas linhas no dataframe
        for key, df in dict_files.items():
            print(key)
            df_final = df_final.append(df, sort=False)
        # salva em xlsx
        with pd.ExcelWriter('allfiles.xlsx') as writer:
            df_final.to_excel(writer, sheet_name='Plan1')
            writer.save()
        # salva em xls, xlsx corrompe as vezes
        with pd.ExcelWriter('allfiles.xls') as writer:
            df_final.to_excel(writer, sheet_name='Plan1')
            writer.save()
    else:
        print('Pasta não existe')
    print('Quantidade de arquivos salvos ' + str(amount_file))


read_files()
