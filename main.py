# -*- coding: utf-8 -*-
"""
Gera relatorio XLSX analisando todos os XMLs da pasta:
 - Conta quantas vezes cada codigo da RELACAO-MODELOS.xlsx aparece em cada XML
 - Separa os resultados pela 3a coluna da RELACAO-MODELOS.xlsx (categoria/ambiente)
"""
import glob
import os
import re
from collections import defaultdict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MODELO_XLSX = 'RELACAO-MODELOS.xlsx'
SAIDA_XLSX = 'RELATORIO_MODELOS.xlsx'

# ---------------------------------------------------------------- estilos
AZUL = PatternFill('solid', fgColor='1F4E78')
AZUL_CLARO = PatternFill('solid', fgColor='DDEBF7')
CINZA = PatternFill('solid', fgColor='F2F2F2')
BRANCO = Font(color='FFFFFF', bold=True, size=11)
NEGRITO = Font(bold=True)
BORDA = Border(*[Side(style='thin', color='BFBFBF')] * 4)

CORES_CAT = {
    'COZINHA': 'FFF2CC',
    'COZINHA BASIC': 'FCE4D6',
    'COZINHA CONSTRUTOR': 'F8CBAD',
    'DORMITORIO': 'E2EFDA',
    'DORMITORIO BASIC': 'D9E1F2',
    'DORMITORIO CONSTRUTOR': 'BDD7EE',
    'LINEAR': 'EDEDED',
    'ESPECIAIS': 'E4DFEC',
}


def carregar_modelos(caminho):
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb.worksheets[0]
    modelos = []
    for r in range(2, ws.max_row + 1):
        item = ws.cell(row=r, column=1).value
        desc = ws.cell(row=r, column=2).value
        cat = ws.cell(row=r, column=3).value
        if item and str(item).strip():
            modelos.append({
                'item': str(item).strip(),
                'desc': str(desc).strip() if desc else '',
                'cat': str(cat).strip() if cat else 'SEM CATEGORIA',
            })
    return modelos


def main():
    modelos = carregar_modelos(MODELO_XLSX)

    xml_files = sorted(glob.glob('*.xml'))
    if not xml_files:
        print('Nenhum arquivo XML encontrado na pasta.')
        return

    # contagem[xml][item] = quantidade
    contagem = defaultdict(lambda: defaultdict(int))
    total_por_xml = defaultdict(int)

    for xf in xml_files:
        with open(xf, 'r', encoding='utf-8', errors='replace') as f:
            conteudo = f.read().upper()
        for m in modelos:
            padrao = r'(?<![A-Z0-9])' + re.escape(m['item'].upper()) + r'(?![A-Z0-9])'
            qtd = len(re.findall(padrao, conteudo))
            if qtd:
                contagem[xf][m['item']] += qtd
                total_por_xml[xf] += qtd

        # Identifica códigos especiais ES (ex: ES01, ES02, ES08, ES) que não estejam na planilha
        es_encontrados = set(re.findall(r'(?<![A-Z0-9])(ES\d+|ES)(?![A-Z0-9])', conteudo))
        for es_item in es_encontrados:
            if not any(m['item'].upper() == es_item for m in modelos):
                modelos.append({
                    'item': es_item,
                    'desc': f'ESPECIAL {es_item}',
                    'cat': 'ESPECIAIS',
                })
                padrao = r'(?<![A-Z0-9])' + re.escape(es_item) + r'(?![A-Z0-9])'
                qtd = len(re.findall(padrao, conteudo))
                if qtd:
                    contagem[xf][es_item] += qtd
                    total_por_xml[xf] += qtd

    # somente os codigos que realmente apareceram em pelo menos um XML
    modelos_encontrados = [
        m for m in modelos
        if sum(contagem[xf].get(m['item'], 0) for xf in xml_files) > 0
    ]
    if not modelos_encontrados:
        print('Nenhum codigo da planilha foi encontrado nos arquivos XML.')
        return

    wb = openpyxl.Workbook()

    # =========================================================== RESUMO
    ws = wb.active
    ws.title = 'Resumo por Categoria'
    ws['A1'] = 'RELATORIO DE MODELOS - ANALISE DE CODIGOS NOS ARQUIVOS XML'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws.merge_cells('A1:D1')

    linhas = [
        ('Planilha de modelos:', MODELO_XLSX),
        ('Arquivos XML analisados:', ', '.join(xml_files)),
        ('Total de codigos na planilha:', len(modelos)),
        ('Codigos encontrados (ocorrencia > 0):', len(modelos_encontrados)),
    ]
    for i, (k, v) in enumerate(linhas, start=2):
        ws.cell(row=i, column=1, value=k).font = NEGRITO
        ws.cell(row=i, column=2, value=v)

    cab = ['3a Coluna (Categoria/Ambiente)', 'Qtd de Codigos Encontrados',
           'Total de Ocorrencias', '% do Total']
    linha_cab = 6
    for c, titulo in enumerate(cab, start=1):
        cel = ws.cell(row=linha_cab, column=c, value=titulo)
        cel.fill = AZUL
        cel.font = BRANCO
        cel.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cel.border = BORDA

    total_geral = sum(total_por_xml.values())
    agg = defaultdict(lambda: {'achados': 0, 'ocorr': 0})
    for m in modelos:
        a = agg[m['cat']]
        achados = sum(contagem[xf].get(m['item'], 0) > 0 for xf in xml_files)
        if achados == 0:
            continue  # categoria sem nenhuma ocorrencia -> nao aparece
        a['achados'] += achados
        a['ocorr'] += sum(contagem[xf].get(m['item'], 0) for xf in xml_files)

    linha = linha_cab + 1
    for cat in sorted([c for c in agg if agg[c]['ocorr'] > 0],
                      key=lambda c: -agg[c]['ocorr']):
        a = agg[cat]
        ws.cell(row=linha, column=1, value=cat)
        ws.cell(row=linha, column=2, value=a['achados'])
        ws.cell(row=linha, column=3, value=a['ocorr'])
        ws.cell(row=linha, column=4,
                value=(a['ocorr'] / total_geral) if total_geral else 0)
        ws.cell(row=linha, column=4).number_format = '0.0%'
        for c in range(1, 5):
            cel = ws.cell(row=linha, column=c)
            cel.border = BORDA
            cel.fill = PatternFill('solid', fgColor=CORES_CAT.get(cat, 'FFFFFF'))
            if c in (2, 3):
                cel.alignment = Alignment(horizontal='center')
        linha += 1

    ws.cell(row=linha, column=1, value='TOTAL').font = NEGRITO
    ws.cell(row=linha, column=2,
            value=sum(a['achados'] for a in agg.values())).font = NEGRITO
    ws.cell(row=linha, column=3, value=total_geral).font = NEGRITO
    ws.cell(row=linha, column=4,
            value=1.0 if total_geral else 0).font = NEGRITO
    ws.cell(row=linha, column=4).number_format = '0.0%'
    for c in range(1, 5):
        ws.cell(row=linha, column=c).fill = AZUL_CLARO
        ws.cell(row=linha, column=c).border = BORDA

    for col, larg in zip('ABCD', (30, 26, 20, 14)):
        ws.column_dimensions[col].width = larg
    ws.freeze_panes = 'A7'

    # =========================================================== DETALHE
    ws2 = wb.create_sheet('Detalhe por Codigo')
    cabecalhos = ['Item', 'Descricao (Item)', '3a Coluna (Categoria/Ambiente)']
    cabecalhos += [f'Ocorrencias: {os.path.basename(x)}' for x in xml_files]
    cabecalhos += ['TOTAL', 'Qtd de XMLs com o codigo']

    for c, titulo in enumerate(cabecalhos, start=1):
        cel = ws2.cell(row=1, column=c, value=titulo)
        cel.fill = AZUL
        cel.font = BRANCO
        cel.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cel.border = BORDA

    linha = 2
    cat_atual = None
    for m in sorted(modelos_encontrados,
                    key=lambda x: (x['cat'],
                                   -sum(contagem[xf].get(x['item'], 0) for xf in xml_files))):
        if m['cat'] != cat_atual:
            cat_atual = m['cat']
            ws2.cell(row=linha, column=1, value='=== {} ==='.format(cat_atual))
            ws2.cell(row=linha, column=1).font = Font(bold=True, size=12)
            for c in range(1, len(cabecalhos) + 1):
                ws2.cell(row=linha, column=c).fill = PatternFill(
                    'solid', fgColor=CORES_CAT.get(cat_atual, 'FFFFFF'))
            linha += 1

        valores = [contagem[xf].get(m['item'], 0) for xf in xml_files]
        total = sum(valores)
        n_xmls = sum(1 for v in valores if v > 0)

        ws2.cell(row=linha, column=1, value=m['item']).font = NEGRITO
        ws2.cell(row=linha, column=2, value=m['desc'])
        ws2.cell(row=linha, column=3, value=m['cat'])
        col = 4
        for v in valores:
            ws2.cell(row=linha, column=col, value=v)
            col += 1
        ws2.cell(row=linha, column=col, value=total).font = NEGRITO
        ws2.cell(row=linha, column=col + 1, value=n_xmls)

        for c in range(1, len(cabecalhos) + 1):
            cel = ws2.cell(row=linha, column=c)
            cel.border = BORDA
            if c >= 4:
                cel.alignment = Alignment(horizontal='center')
        linha += 1

    larguras = [12, 46, 28] + [22] * len(xml_files) + [12, 20]
    for i, larg in enumerate(larguras, start=1):
        ws2.column_dimensions[get_column_letter(i)].width = larg
    ws2.freeze_panes = 'D2'
    ws2.auto_filter.ref = 'A1:{}{}'.format(
        get_column_letter(len(cabecalhos)), linha - 1)

    # =========================================================== POR AMBIENTE
    ws3 = wb.create_sheet('Contagem por Ambiente')
    wb.remove(ws3)
    ws3 = wb.create_sheet('Contagem por Ambiente')
    tit = ['3a Coluna (Ambiente)', 'Item', 'Descricao (Item)']
    tit += [f'{os.path.basename(x)}' for x in xml_files]
    tit += ['TOTAL']
    for c, t in enumerate(tit, start=1):
        cel = ws3.cell(row=1, column=c, value=t)
        cel.fill = AZUL
        cel.font = BRANCO
        cel.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cel.border = BORDA

    linha = 2
    for cat in sorted([c for c in agg if agg[c]['ocorr'] > 0],
                      key=lambda c: -agg[c]['ocorr']):
        ws3.cell(row=linha, column=1, value=cat).font = Font(bold=True, size=12)
        for c in range(1, len(tit) + 1):
            ws3.cell(row=linha, column=c).fill = PatternFill(
                'solid', fgColor=CORES_CAT.get(cat, 'FFFFFF'))
            ws3.cell(row=linha, column=c).border = BORDA
        linha += 1

        subs = sorted(
            [m for m in modelos_encontrados if m['cat'] == cat],
            key=lambda x: -sum(contagem[xf].get(x['item'], 0) for xf in xml_files))
        for m in subs:
            valores = [contagem[xf].get(m['item'], 0) for xf in xml_files]
            ws3.cell(row=linha, column=2, value=m['item']).font = NEGRITO
            ws3.cell(row=linha, column=3, value=m['desc'])
            col = 4
            for v in valores:
                ws3.cell(row=linha, column=col, value=v)
                col += 1
            ws3.cell(row=linha, column=col, value=sum(valores)).font = NEGRITO
            for c in range(2, len(tit) + 1):
                cel = ws3.cell(row=linha, column=c)
                cel.border = BORDA
                if c >= 4:
                    cel.alignment = Alignment(horizontal='center')
            linha += 1
        linha += 1

    for i, larg in enumerate([28, 12, 46] + [22] * len(xml_files) + [12], start=1):
        ws3.column_dimensions[get_column_letter(i)].width = larg
    ws3.freeze_panes = 'D2'

    wb.save(SAIDA_XLSX)

    print('Relatorio gerado com sucesso: {}'.format(SAIDA_XLSX))
    print('\nTotal de ocorrencias por XML:')
    for xf in xml_files:
        print('  {}: {}'.format(os.path.basename(xf), total_por_xml[xf]))
    print('\nTotal de ocorrencias por 3a coluna:')
    for cat in sorted([c for c in agg if agg[c]['ocorr'] > 0],
                      key=lambda c: -agg[c]['ocorr']):
        print('  {:22s} codigos={:3d}  ocorrencias={}'.format(
            cat, agg[cat]['achados'], agg[cat]['ocorr']))
    print('\nTOTAL GERAL de ocorrencias: {}'.format(total_geral))
    print('Codigos com ocorrencia > 0: {} de {}'.format(
        len(modelos_encontrados), len(modelos)))


if __name__ == '__main__':
    main()
