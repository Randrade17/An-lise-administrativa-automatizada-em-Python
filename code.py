# Analise-administrativa-automatizada-em-Python


O projeto Análise Administrativa Automatizada em Python tem como objetivo otimizar o trabalho de analistas administrativos por meio da automação de tarefas repetitivas e da consolidação inteligente de dados corporativos.

A aplicação lê e unifica relatórios em formatos CSV e Excel, realiza cálculos automáticos de indicadores financeiros e operacionais (como receita, custos, lucro e produtividade), e gera relatórios visuais e interativos com base nesses dados.

Além disso, o sistema exporta resultados em Excel e PDF, produz gráficos interativos com o uso do Plotly e pode ser expandido para envio automático de relatórios por e-mail.

Essa automação contribui para uma gestão mais ágil e precisa, reduzindo erros humanos e liberando tempo para análises estratégicas.


"""
Automação Administrativa em Python
Autor: Rafael Figueiredo (exemplo)
Descrição:
    Este script consolida relatórios administrativos e financeiros,
    calcula indicadores, gera gráficos e exporta relatórios automáticos.
"""

import os
import sys
import argparse
import pandas as pd
import plotly.express as px
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from openpyxl import Workbook

# -------------------- PRINCIPAIS FUNÇÕES -------------------- #

def carregar_dados(diretorio: str) -> pd.DataFrame:
    """
    Vai ler todos os arquivos CSV/Excel no diretório informado e retorna um DataFrame consolidado.
    """
    arquivos = [f for f in os.listdir(diretorio) if f.endswith(('.csv', '.xlsx'))]
    if not arquivos:
        raise FileNotFoundError("Nenhum arquivo CSV ou Excel encontrado no diretório informado.")
    
    dfs = []
    for arquivo in arquivos:
        caminho = os.path.join(diretorio, arquivo)
        try:
            if arquivo.endswith('.csv'):
                df = pd.read_csv(caminho)
            else:
                df = pd.read_excel(caminho)
            df['Origem_Arquivo'] = arquivo
            dfs.append(df)
        except Exception as e:
            print(f"[AVISO] Erro ao carregar {arquivo}: {e}")

    if not dfs:
        raise ValueError("Nenhum dado pôde ser carregado.")
    
    df_final = pd.concat(dfs, ignore_index=True)
    return df_final


def calcular_metricas(df: pd.DataFrame) -> dict:
    """
    Calcula indicadores administrativos e financeiros com base no DataFrame consolidado.
    """
    metricas = {}

    # Garantir que as colunas esperadas existam
    colunas = df.columns.str.lower()
    if 'receita' in colunas:
        metricas['Receita Total'] = df.loc[:, df.columns[colunas == 'receita'][0]].sum()
    if 'despesa' in colunas:
        metricas['Custo Operacional Total'] = df.loc[:, df.columns[colunas == 'despesa'][0]].sum()
    if 'funcionario' in colunas and 'producao' in colunas:
        metricas['Produtividade Média'] = df['producao'].sum() / df['funcionario'].nunique()
    
    if 'receita' in colunas and 'despesa' in colunas:
        metricas['Lucro Líquido'] = metricas['Receita Total'] - metricas['Custo Operacional Total']
    
    if 'vendas' in colunas and 'quantidade' in colunas:
        metricas['Ticket Médio'] = df['vendas'].sum() / df['quantidade'].sum()

    return metricas


def gerar_graficos(df: pd.DataFrame, pasta_saida: str):
    """
    Gera gráficos interativos com Plotly e salva como HTML.
    """
    if 'data' in df.columns:
        df['data'] = pd.to_datetime(df['data'], errors='coerce')
        df = df.dropna(subset=['data'])
        df['mês'] = df['data'].dt.to_period('M').astype(str)

        if 'receita' in df.columns:
            fig = px.line(df, x='mês', y='receita', title='Evolução da Receita Mensal')
            fig.write_html(os.path.join(pasta_saida, 'grafico_receita.html'))

        if 'despesa' in df.columns:
            fig = px.bar(df, x='mês', y='despesa', title='Custos Operacionais por Mês')
            fig.write_html(os.path.join(pasta_saida, 'grafico_despesa.html'))

        print("[INFO] Gráficos interativos gerados com sucesso.")


def exportar_excel(metricas: dict, df: pd.DataFrame, pasta_saida: str):
    """
    Exporta os dados consolidados e métricas para um arquivo Excel.
    """
    caminho_excel = os.path.join(pasta_saida, 'relatorio_administrativo.xlsx')
    with pd.ExcelWriter(caminho_excel, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados Consolidados')
        pd.DataFrame(metricas.items(), columns=['Métrica', 'Valor']).to_excel(writer, index=False, sheet_name='Indicadores')
    print(f"[INFO] Relatório Excel salvo em: {caminho_excel}")


def exportar_pdf(metricas: dict, pasta_saida: str):
    """
    Gera um PDF simples com as métricas principais.
    """
    caminho_pdf = os.path.join(pasta_saida, 'relatorio_administrativo.pdf')
    c = canvas.Canvas(caminho_pdf, pagesize=A4)
    c.setTitle("Relatório Administrativo")

    c.setFont("Helvetica-Bold", 16)
    c.drawString(200, 800, "Relatório Administrativo")
    c.setFont("Helvetica", 12)
    y = 760
    for k, v in metricas.items():
        c.drawString(100, y, f"{k}: {v:,.2f}")
        y -= 20
    c.save()
    print(f"[INFO] Relatório PDF salvo em: {caminho_pdf}")


# -------------------- FLUXO PRINCIPAL -------------------- #

def main():
    parser = argparse.ArgumentParser(description="Automação de relatórios administrativos")
    parser.add_argument('-i', '--input', required=True, help="Diretório contendo arquivos CSV/XLSX")
    parser.add_argument('-o', '--output', default='saida_relatorio', help="Pasta de saída para relatórios")
    args = parser.parse_args()
   
    os.makedirs(args.output, exist_ok=True)

    try:
        print("[INFO] Carregando dados...")
        df = carregar_dados(args.input)
        print("[INFO] Calculando métricas...")
        metricas = calcular_metricas(df)
        print("[INFO] Gerando gráficos...")
        gerar_graficos(df, args.output)
        print("[INFO] Exportando relatórios...")
        exportar_excel(metricas, df, args.output)
        exportar_pdf(metricas, args.output)
        print("[SUCESSO] Processo concluído!")
    except Exception as e:
        print(f"[ERRO] {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
