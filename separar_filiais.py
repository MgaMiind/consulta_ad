"""
Created on Thu Aug 17 09:53:07 2026

@author: Gabriel Campos
"""

import os
import re
import subprocess
import pandas as pd

# CONFIGURAÇÕES DE ARQUIVOS

ARQUIVO_ENTRADA = "relatorio_trend.xlsx"  
PASTA_SAIDA = "Relatorios_Por_Filial"     

# FUNÇÃO PARA CONSULTAR A LOCALIDADE DA MÁQUINA NO AD

def buscar_filial_no_ad(hostname):
    """
    Consulta a máquina no Active Directory via PowerShell e extrai a OU (Filial)
    """
    if pd.isna(hostname) or not str(hostname).strip():
        return "SEM HOSTNAME"

    hostname = str(hostname).strip()
    
    cmd = f'powershell -NoProfile -Command "(Get-ADComputer -Filter \\"Name -eq \'{hostname}\'\\" -Properties DistinguishedName).DistinguishedName"'
    
    try:

        resultado = subprocess.run(cmd, capture_output=True, text=True, shell=True)
        dn = resultado.stdout.strip()

        if not dn:
            return "NÃO ENCONTRADO NO AD"

        ous = re.findall(r'OU=([^,]+)', dn)

        if ous:
            return " / ".join(ous)

        else:
            return "AD - Raiz (Sem OU)"

    except Exception as e:
        return f"Erro na Consulta: {str(e)}"

# EXECUÇÃO PRINCIPAL

def processar_relatorio():
    print("=" * 60)
    print("      INICIANDO AUTOMAÇÃO DE SEPARAÇÃO POR FILIAL")
    print("=" * 60)

    if not os.path.exists(ARQUIVO_ENTRADA):
        print(f"\n[ERRO] O arquivo '{ARQUIVO_ENTRADA}' não foi encontrado na pasta atual.")
        print("Por favor, renomeie a planilha da Trend para 'relatorio_trend.xlsx' e tente novamente.")
        return

    if not os.path.exists(PASTA_SAIDA):
        os.makedirs(PASTA_SAIDA)

    print("\n1. Carregando as guias do Excel...")
    xl = pd.ExcelFile(ARQUIVO_ENTRADA)
    
    aba_maquinas_recorrentes = xl.sheet_names[0]
    #aba_problemas = xl.sheet_names[1]

    df_maquinas_recorrentes = xl.parse(aba_maquinas_recorrentes)
    #df_problemas = xl.parse(aba_problemas)

    coluna_host_1 = df_maquinas_recorrentes.columns[3]  # Pega a 1ª coluna
    #coluna_host_2 = df_problemas.columns[0]

    # Mapeamento de cache para não consultar a mesma máquina duas vezes no AD
    cache_filiais = {}

    def obter_filial_com_cache(host):
        if host not in cache_filiais:
            print(f"   -> Pesquisando no AD: {host}...")
            cache_filiais[host] = buscar_filial_no_ad(host)
        return cache_filiais[host]

    print(f"\n2. Identificando localidades da Aba 1 ({aba_maquinas_recorrentes})...")
    
    resultados_pesquisa = df_maquinas_recorrentes[coluna_host_1].apply(obter_filial_com_cache)
    
    df_maquinas_recorrentes.insert(loc=4, column='Localidade_AD', value=resultados_pesquisa)    
    #df_maquinas_recorrentes['Localidade_AD'] = df_maquinas_recorrentes[coluna_host_1].apply(obter_filial_com_cache)

    #print(f"\n3. Identificando localidades da Aba 2 ({aba_problemas})...")
    #df_problemas['Localidade_AD'] = df_problemas[coluna_host_2].apply(obter_filial_com_cache)

    todas_localidades = set(df_maquinas_recorrentes['Localidade_AD'])#.union(set(df_problemas['Localidade_AD']))

    print(f"\n3. Gerando arquivos por filial na pasta '{PASTA_SAIDA}'...")

    for local in todas_localidades:

        nome_arquivo_limpo = re.sub(r'[\\/*?:"<>|]', '_', str(local))
        caminho_saida = os.path.join(PASTA_SAIDA, f"Trend_{nome_arquivo_limpo}.xlsx")

        sub_df_maquinas_recorrentes = df_maquinas_recorrentes[df_maquinas_recorrentes['Localidade_AD'] == local]
        #sub_df_problemas = df_problemas[df_problemas['Localidade_AD'] == local]

        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            sub_df_maquinas_recorrentes.to_excel(writer, sheet_name=aba_maquinas_recorrentes, index=False)
            #sub_df_problemas.to_excel(writer, sheet_name=aba_problemas, index=False)

        print(f"   [OK] Criado: Trend_{nome_arquivo_limpo}.xlsx")
        
    print("\n4. Gerando planilha master (Relatório consolidado)...")
    
    arquivo_master = "relatório_trend_MASTER.xlsx"
    
    df_maquinas_recorrentes.to_excel(arquivo_master, index=False, engine='openpyxl')    
    
    print(f"   [OK] Planilha master gerada: {arquivo_master}")

    print("\n")
    print("=" * 60)
    print("  PROCESSO CONCLUÍDO COM SUCESSO!")
    print("=" * 60)

if __name__ == "__main__":
    processar_relatorio()