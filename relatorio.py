"""Exportação do relatório de consultas para Excel."""
import argparse
import csv
import json
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from bot import ler_cpfs


def montar_dados(base, saida):
    cpfs = ler_cpfs(base, 'CPF')
    progresso = saida / 'progresso.json'
    registros = json.loads(progresso.read_text(encoding='utf-8')) if progresso.exists() else {}
    historico = []
    log = saida / 'consultas.csv'
    if log.exists():
        with log.open(encoding='utf-8-sig', newline='') as stream:
            historico = [r for r in csv.DictReader(stream, delimiter=';') if r['cpf'] in cpfs]
    ultimos = {r['cpf']: r for r in historico}
    consultas, cartoes = [], []
    for cpf in cpfs:
        itens = registros.get(cpf, [])
        if cpf in registros and not itens:
            raise ValueError('Progresso contém consulta sem cartões; revise os dados.')
        total = Decimal('0')
        for item in itens:
            valor = Decimal(item['disponivel_saque'].replace('.', '').replace(',', '.'))
            total += valor
            cartoes.append([cpf, item['cartao'], item['status_cartao'], float(valor)])
        motivo = '' if itens else ultimos.get(cpf, {}).get('observacao', 'Consulta ainda não executada')
        consultas.append([cpf, 'Concluído' if itens else 'Pendente', len(itens), float(total) if itens else None, motivo])
    return consultas, cartoes, historico


def exportar(base, saida, destino):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.formatting.rule import CellIsRule, FormulaRule

    consultas, cartoes, historico = montar_dados(base, saida)
    wb = Workbook()
    resumo = wb.active
    resumo.title = 'Resumo'
    resumo.append(['Relatório de consultas BMG'])
    resumo.append(['Gerado em', datetime.now().strftime('%d/%m/%Y %H:%M:%S')])
    resumo.append(['Base de entrada', base.name])
    resumo.append(['CPFs na base', len(consultas)])
    resumo.append(['CPFs concluídos', sum(r[1] == 'Concluído' for r in consultas)])
    resumo.append(['CPFs pendentes', sum(r[1] == 'Pendente' for r in consultas)])
    resumo.append(['Cartões encontrados', len(cartoes)])
    resumo.append(['Disponível para saque — total', sum(r[3] for r in cartoes)])
    resumo['B8'].number_format = '"R$" #,##0.00'
    resumo.append(['Escopo', 'Somente os CPFs da planilha selecionada.'])
    resumo.append(['Valores', 'Valores observados no portal; soma por cartão, sem garantia de saque.'])
    resumo.append(['Atualização', 'Resultados retomados podem vir de execuções anteriores.'])
    resumo.append(['Data da consulta', 'Não registrada nos dados antigos; data acima é de geração do relatório.'])
    resumo.append(['Pendências', 'Sem resultado confirmado. Valores em branco não significam saldo zero.'])

    def tabela(nome, headers, rows, widths):
        sheet = wb.create_sheet(nome)
        sheet.append(headers)
        for row in rows:
            sheet.append(row)
        sheet.freeze_panes = 'A2'
        sheet.sheet_view.showGridLines = False
        if rows:
            tab = Table(displayName=nome.replace(' ', '') + 'Tabela', ref=sheet.dimensions)
            tab.tableStyleInfo = TableStyleInfo(name='TableStyleMedium9', showRowStripes=True)
            sheet.add_table(tab)
        for cell in sheet[1]:
            cell.fill = PatternFill('solid', fgColor='C45A11')
            cell.font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
            cell.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)
        sheet.row_dimensions[1].height = 34
        for col, width in widths.items():
            sheet.column_dimensions[col].width = width
        for row in sheet.iter_rows(min_row=2):
            sheet.row_dimensions[row[0].row].height = 32
            for cell in row:
                cell.font = Font(name='Calibri', size=11)
                cell.alignment = Alignment(vertical='center', wrap_text=True)
                # Texto exportado nunca deve ser interpretado como fórmula.
                if isinstance(cell.value, str):
                    cell.data_type = 's'
            row[0].number_format = '@'
            cpf = row[0].value
            if isinstance(cpf, str) and len(cpf) == 11 and cpf.isdigit():
                row[0].value = f'{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}'
            row[0].alignment = Alignment(vertical='center', horizontal='left')
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.orientation = 'landscape'
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.print_title_rows = '1:1'
        return sheet

    consulta = tabela('Consultas', ['CPF', 'Situação da consulta', 'Quantidade de cartões', 'Disponível saque — total (R$)', 'Pendência / motivo'], consultas, {'A':19, 'B':24, 'C':24, 'D':31, 'E':95})
    detalhes = tabela('Cartoes', ['CPF', 'Tipo de cartão', 'Status do cartão', 'Disponível para saque (R$)'], cartoes, {'A':22,'B':28,'C':36,'D':32})
    pendentes = [r for r in consultas if r[1] == 'Pendente']
    tabela('Pendencias', ['CPF', 'Situação', 'Quantidade confirmada', 'Total confirmado (R$)', 'Motivo'], pendentes, {'A':19,'B':22,'C':26,'D':28,'E':95})
    tabela('Historico', ['CPF', 'Registro da execução', 'Detalhes'], [[r['cpf'], r['status'], r['observacao']] for r in historico], {'A':19,'B':30,'C':100})
    for sheet, coluna in [(consulta, 'D'), (detalhes, 'D')]:
        for cell in sheet[coluna][1:]:
            cell.number_format = '"R$" #,##0.00'
    if cartoes:
        detalhes.conditional_formatting.add(f'D2:D{len(cartoes)+1}', CellIsRule(operator='greaterThan', formula=['0'], fill=PatternFill('solid', fgColor='E2F0D9')))
    if consultas:
        consulta.conditional_formatting.add(f'A2:E{len(consultas)+1}', FormulaRule(formula=['$B2="Pendente"'], fill=PatternFill('solid', fgColor='FFF2CC')))
    for sheet in (consulta, detalhes):
        for cell in sheet['D'][1:]:
            cell.alignment = Alignment(horizontal='right', vertical='center')
    resumo.sheet_view.showGridLines = False
    resumo.column_dimensions['A'].width = 38
    resumo.column_dimensions['B'].width = 100
    for row in resumo:
        resumo.row_dimensions[row[0].row].height = 27
        for cell in row:
            cell.font = Font(name='Calibri', size=11)
            cell.alignment = Alignment(vertical='center', wrap_text=True)
    resumo['A1'].font = Font(name='Calibri', size=17, bold=True, color='C45A11')
    resumo.row_dimensions[1].height = 36
    # Abre diretamente na lista legível de cartões; o resumo fica disponível na primeira aba.
    wb.active = wb.sheetnames.index('Cartoes')
    destino.parent.mkdir(parents=True, exist_ok=True)
    temporario = destino.with_name(destino.stem + '.tmp.xlsx')
    try:
        wb.save(temporario)
        temporario.replace(destino)
    finally:
        if temporario.exists():
            temporario.unlink()
    return len(consultas), len(cartoes), len(pendentes)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('base', type=Path)
    parser.add_argument('destino', type=Path)
    args = parser.parse_args()
    contagem = exportar(args.base, Path(__file__).parent / 'saida', args.destino)
    print(f'Relatório salvo: {contagem[0]} CPFs, {contagem[1]} cartões e {contagem[2]} pendências.')
