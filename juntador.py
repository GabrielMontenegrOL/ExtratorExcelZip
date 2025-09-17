import os
import zipfile

def extrair_excels_de_zips(pasta_zip, pasta_destino):
    # Cria a pasta destino se não existir
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)

    # Lista todos os arquivos na pasta_zip
    arquivos = os.listdir(pasta_zip)

    # Filtra arquivos ZIP
    arquivos_zip = [arq for arq in arquivos if arq.lower().endswith('.zip')]

    for arquivo_zip in arquivos_zip:
        caminho_zip = os.path.join(pasta_zip, arquivo_zip)
        with zipfile.ZipFile(caminho_zip, 'r') as zip_ref:
            # Lista arquivos dentro do zip
            for arquivo_interno in zip_ref.namelist():
                if arquivo_interno.lower().endswith(('.xls', '.xlsx')):
                    # Extrai para a pasta destino
                    print(f"Extraindo {arquivo_interno} de {arquivo_zip}")
                    zip_ref.extract(arquivo_interno, pasta_destino)

                    # Mover o arquivo para a raiz da pasta destino (pois pode vir com pastas dentro)
                    caminho_extraido = os.path.join(pasta_destino, arquivo_interno)
                    novo_caminho = os.path.join(pasta_destino, os.path.basename(arquivo_interno))
                    if caminho_extraido != novo_caminho:
                        os.rename(caminho_extraido, novo_caminho)

                    # Limpa pastas vazias criadas na extração
                    pasta_interna = os.path.dirname(caminho_extraido)
                    if pasta_interna and os.path.exists(pasta_interna):
                        try:
                            os.removedirs(pasta_interna)
                        except OSError:
                            pass

if __name__ == "__main__":
    pasta_com_zips = "C:/Users/gmob/Desktop/testes/zip_files"  # Alterar para o caminho da pasta que contém os ZIPs -> SALVE AQUI
    pasta_de_extracao = "C:/Users/gmob/Desktop/testes/todos_excel"  # Alterar para o caminho da pasta onde quer salvar os Excel extraídos -> TODOS

    extrair_excels_de_zips(pasta_com_zips, pasta_de_extracao)
    print("Extração concluída!")
