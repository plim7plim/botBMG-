"""Consulta de cartões BMG com login manual e entrada CSV/XLSX."""
import argparse
import csv
import json
import re
import time
import os
import socket
import subprocess
import tempfile
from pathlib import Path

LOGIN = 'https://portalfranqueado.help.com.br/cas/login?service=https%3A%2F%2Fdnabmg.help.com.br%2Flogin%2Fcas'
CONSULTA = 'https://dnabmg.help.com.br/#/servicos/consulta-cartao'


def pedir_confirmacao(mensagem):
    print('BOT_PROMPT|' + mensagem, flush=True)
    return input()


def extrair_cartoes(texto):
    texto = texto.replace('\u00ad', '').replace('\u200b', '').replace('\ufeff', '')
    # Exclui cartões históricos para não misturar contratos antigos e atuais.
    texto = re.split(r'Cartões\s*-\s*Histórico', texto, flags=re.I)[0]
    blocos = re.split(r'(?=\bBMG\s+(?:MULTI|BENEF[IÍ]CIO|CARD INSS|CARD CONV[EÊ]NIO)\b)', texto, flags=re.I)
    resultado = []
    for bloco in blocos:
        nome = re.match(r'BMG\s+(?:MULTI|BENEF[IÍ]CIO|CARD INSS|CARD CONV[EÊ]NIO)\b', bloco, re.I)
        if not nome:
            continue
        status = re.search(r'Status\s*Cartão\s*:\s*(.*?)\s*Status\s*Conta\s*:', bloco, re.I | re.S)
        saque = re.search(r'Disponível\s*Saque\s*:\s*R\$\s*([\d.]+,\d{2})', bloco, re.I)
        if not status or not saque:
            raise ValueError('Resultado incompleto: status ou saque ausente')
        resultado.append({'cartao': nome.group().upper(), 'status_cartao': ' '.join(status.group(1).split()), 'disponivel_saque': saque.group(1)})
    if not resultado:
        raise ValueError('Nenhum cartão reconhecido; confira a tela')
    if len(re.findall(r'Status\s*Cartão\s*:', texto, re.I)) != len(resultado):
        raise ValueError('Há cartões não reconhecidos; revise o mapeamento antes de salvar')
    return resultado


def normalizar_cpf(valor):
    texto = str(valor or '').strip()
    if re.fullmatch(r'\d+\.0', texto):
        texto = texto[:-2]
    digitos = re.sub(r'[.\-\s]', '', texto)
    if not digitos.isdigit() or len(digitos) > 11 or not digitos:
        raise ValueError('CPF deve conter até 11 dígitos')
    return digitos.zfill(11)


def ler_cpfs(arquivo, coluna):
    if arquivo.suffix.lower() == '.xlsx':
        from openpyxl import load_workbook
        workbook = load_workbook(arquivo, read_only=True, data_only=True)
        try:
            linhas = list(workbook.active.iter_rows(values_only=True))
        finally:
            workbook.close()
    elif arquivo.suffix.lower() == '.csv':
        with arquivo.open(encoding='utf-8-sig', newline='') as stream:
            amostra = stream.read(4096)
            stream.seek(0)
            try:
                dialeto = csv.Sniffer().sniff(amostra, delimiters=',;\t')
            except csv.Error:
                dialeto = csv.excel
            linhas = list(csv.reader(stream, dialeto))
    else:
        raise ValueError('Use uma planilha .xlsx ou .csv')
    if not linhas:
        return []
    cabecalho = [str(item or '').strip().casefold() for item in linhas[0]]
    if coluna.casefold() not in cabecalho:
        raise ValueError(f'Coluna {coluna!r} não encontrada')
    indice = cabecalho.index(coluna.casefold())
    cpfs = []
    for numero, linha in enumerate(linhas[1:], start=2):
        valor = linha[indice] if indice < len(linha) else None
        if valor is None or str(valor).strip() == '':
            continue
        try:
            cpf = normalizar_cpf(valor)
        except ValueError as erro:
            raise ValueError(f'Linha {numero}: {erro}') from erro
        if cpf not in cpfs:
            cpfs.append(cpf)
    return cpfs


def sessao_expirada(page):
    return '/cas/login' in page.url or page.get_by_role('button', name='Entrar', exact=True).is_visible()


def campo_consulta(page, seletor=None):
    if seletor:
        return page.locator(seletor)
    campo = page.get_by_label(re.compile(r'^CPF\s*:?$', re.I))
    if campo.count() == 1:
        return campo
    # Algumas versões do portal exibem o rótulo sem associá-lo ao input.
    candidatos = page.locator('input:visible:not([type=hidden]):not([type=button]):not([type=submit]):not([type=checkbox]):not([type=radio])')
    if candidatos.count() != 1:
        raise ValueError('Campo CPF ambíguo. Informe --campo-cpf com o seletor real.')
    return candidatos


def abrir_consulta(page, seletor=None):
    if page.url != CONSULTA:
        page.goto(CONSULTA, wait_until='domcontentloaded')
    page.get_by_role('button', name='Pesquisar', exact=True).wait_for(state='visible', timeout=20000)
    campo_consulta(page, seletor).wait_for(state='visible', timeout=15000)
    return page.url


def preparar_formulario(page, url_consulta, seletor=None):
    if page.url != url_consulta:
        page.goto(url_consulta, wait_until='domcontentloaded')
    elif re.search(r'Status\s*Cartão\s*:', page.locator('body').inner_text(), re.I):
        # Recarrega somente quando precisa limpar o resultado anterior.
        page.reload(wait_until='domcontentloaded')
    if sessao_expirada(page):
        raise PermissionError('Sessão expirada')
    page.get_by_role('button', name='Pesquisar', exact=True).wait_for(state='visible', timeout=20000)
    campo = campo_consulta(page, seletor)
    if re.search(r'Status\s*Cartão\s*:', page.locator('body').inner_text(), re.I):
        raise ValueError('Formulário contém resultado anterior; execução interrompida')
    return campo


def aguardar_login_manual(page, seletor=None):
    while True:
        pedir_confirmacao('Conclua o login e o reCAPTCHA no navegador. Depois pressione Enter aqui: ')
        if sessao_expirada(page):
            print('O portal ainda está na tela de acesso. O login não foi confirmado; consultas aguardando.')
            continue
        return abrir_consulta(page, seletor)


def cpf_resultado(texto):
    encontrados = re.findall(r'CPF\s*:\s*(\d{3}\.?\d{3}\.?\d{3}-?\d{2})\s*-', texto, re.I)
    return normalizar_cpf(encontrados[0]) if len(encontrados) == 1 else None


def aguardar_resultado(page, timeout, cpf_esperado=None):
    limite = time.monotonic() + timeout
    anterior = None
    estavel_desde = None
    ultimo_erro = 'Nenhum resultado recebido'
    while time.monotonic() < limite:
        if sessao_expirada(page):
            raise PermissionError('Sessão expirada')
        texto = page.locator('body').inner_text()
        try:
            if cpf_esperado and cpf_resultado(texto) != cpf_esperado:
                raise ValueError('O CPF do resultado ainda não corresponde ao CPF pesquisado')
            cartoes = extrair_cartoes(texto)
        except ValueError as erro:
            ultimo_erro = str(erro)
            cartoes = None
        # Avisos e relógios fora dos cartões não devem reiniciar a espera.
        if cartoes and cartoes == anterior:
            if estavel_desde is not None and time.monotonic() - estavel_desde >= 2:
                return cartoes
        else:
            estavel_desde = time.monotonic() if cartoes else None
        anterior = cartoes
        page.wait_for_timeout(250)
    raise TimeoutError(f'Resultado não reconhecido: {ultimo_erro}')


def salvar_csv(registros, saida):
    temporario = saida / 'cartoes.tmp'
    with temporario.open('w', encoding='utf-8-sig', newline='') as resultado:
        tabela = csv.writer(resultado, delimiter=';')
        tabela.writerow(['cpf', 'cartao', 'status_cartao', 'disponivel_saque'])
        for documento, itens in registros.items():
            for item in itens:
                tabela.writerow([documento, item['cartao'], item['status_cartao'], item['disponivel_saque']])
    temporario.replace(saida / 'cartoes.csv')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('planilha', type=Path)
    parser.add_argument('--coluna', default='CPF')
    parser.add_argument('--validar', action='store_true', help='Apenas validar a entrada')
    parser.add_argument('--campo-cpf', help='Seletor CSS do campo CPF, se não houver label acessível')
    parser.add_argument('--intervalo', type=float, default=1)
    parser.add_argument('--timeout', type=float, default=45, help='Prazo máximo de cada consulta em segundos')
    args = parser.parse_args()
    cpfs = ler_cpfs(args.planilha, args.coluna)
    print(f'{len(cpfs)} CPFs únicos carregados.')
    if args.validar or not cpfs:
        return
    from playwright.sync_api import sync_playwright
    saida = Path(__file__).parent / 'saida'
    saida.mkdir(exist_ok=True)
    chrome = next((p for p in [
        Path(os.environ.get('PROGRAMFILES', 'C:/Program Files')) / 'Google/Chrome/Application/chrome.exe',
        Path(os.environ.get('LOCALAPPDATA', '')) / 'Google/Chrome/Application/chrome.exe',
        Path('C:/Program Files (x86)/Google/Chrome/Application/chrome.exe'),
    ] if p.is_file()), None)
    if chrome is None:
        raise FileNotFoundError('Google Chrome não encontrado neste computador.')
    with socket.socket() as porta_local:
        porta_local.bind(('127.0.0.1', 0))
        porta = porta_local.getsockname()[1]
    with tempfile.TemporaryDirectory(prefix='botbmg-login-', ignore_cleanup_errors=True) as perfil:
        processo_chrome = subprocess.Popen([str(chrome), f'--user-data-dir={perfil}',
            f'--remote-debugging-port={porta}', '--remote-debugging-address=127.0.0.1',
            '--start-maximized', '--no-first-run', LOGIN], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        try:
            pedir_confirmacao('Faça login manualmente no Chrome. Só depois pressione Enter aqui: ')
            if processo_chrome.poll() is not None:
                raise RuntimeError('O navegador foi fechado antes de iniciar as consultas.')
            inicio_conexao = time.monotonic()
            print('Conectando ao Chrome já autenticado…', flush=True)
            with sync_playwright() as playwright:
                browser = playwright.chromium.connect_over_cdp(f'http://127.0.0.1:{porta}', timeout=10000)
                print(f'Chrome conectado em {time.monotonic() - inicio_conexao:.1f}s. Preparando consultas…', flush=True)
                executar_consultas(browser, args, cpfs, saida)
        finally:
            if processo_chrome.poll() is None:
                processo_chrome.terminate()
                try:
                    processo_chrome.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    processo_chrome.kill()


def executar_consultas(browser, args, cpfs, saida):
        context = browser.contexts[0]
        paginas = [p for p in context.pages if 'help.com.br' in p.url]
        if len(paginas) != 1:
            raise ValueError('Mantenha somente uma aba do portal BMG aberta no navegador de consultas.')
        page = paginas[0]
        try:
            if sessao_expirada(page):
                url_consulta = aguardar_login_manual(page, args.campo_cpf)
            else:
                print('Abrindo formulário de consulta…', flush=True)
                url_consulta = abrir_consulta(page, args.campo_cpf)
            inicio_sessao = time.monotonic()
            progresso = saida / 'progresso.json'
            registros = json.loads(progresso.read_text(encoding='utf-8')) if progresso.exists() else {}
            salvar_csv(registros, saida)
            falhas = 0
            print('Consultas automáticas iniciadas. Ctrl+C interrompe; o progresso salvo permite retomar.')
            with (saida / 'consultas.csv').open('a', encoding='utf-8-sig', newline='') as stream:
                writer = csv.writer(stream, delimiter=';')
                if stream.tell() == 0:
                    writer.writerow(['cpf', 'status', 'observacao'])
                    stream.flush()
                pendentes = [cpf for cpf in cpfs if cpf not in registros]
                for numero, cpf in enumerate(pendentes, start=1):
                    if cpf in registros:
                        continue
                    try:
                        if time.monotonic() - inicio_sessao >= 85 * 60 or '/cas/login' in page.url:
                            page.goto(LOGIN)
                            url_consulta = aguardar_login_manual(page, args.campo_cpf)
                            inicio_sessao = time.monotonic()
                        for tentativa in range(2):
                            try:
                                print(f'Iniciando consulta {numero}/{len(pendentes)}…', flush=True)
                                campo = preparar_formulario(page, url_consulta, args.campo_cpf)
                                campo.fill(cpf, timeout=15000)
                                if re.search(r'Status\s*Cartão\s*:', page.locator('body').inner_text(), re.I):
                                    raise ValueError('Formulário contém resultado anterior; execução interrompida')
                                page.get_by_role('button', name='Pesquisar', exact=True).click()
                                print(f'Consulta {numero}/{len(pendentes)} enviada. Aguardando resultado do portal…', flush=True)
                                cartoes = aguardar_resultado(page, args.timeout, cpf)
                                break
                            except Exception:
                                if tentativa == 1 or not sessao_expirada(page):
                                    raise
                                print('Sessão expirada. Renove o login no navegador.')
                                url_consulta = aguardar_login_manual(page, args.campo_cpf)
                                inicio_sessao = time.monotonic()
                        registros[cpf] = cartoes
                        temporario = progresso.with_suffix('.tmp')
                        temporario.write_text(json.dumps(registros, ensure_ascii=False, indent=2), encoding='utf-8')
                        temporario.replace(progresso)
                        salvar_csv(registros, saida)
                        writer.writerow([cpf, 'extraido_automaticamente', f'{len(cartoes)} cartões'])
                        stream.flush()
                        print(f'Consulta {numero}/{len(pendentes)} salva: {len(cartoes)} cartões.')
                        time.sleep(max(0, args.intervalo))
                    except Exception as erro:
                        writer.writerow([cpf, 'erro', f'{type(erro).__name__}: {erro}'])
                        stream.flush()
                        falhas += 1
                        diagnostico = saida / 'diagnostico'
                        diagnostico.mkdir(exist_ok=True)
                        try:
                            page.screenshot(path=str(diagnostico / f'consulta_{numero}.png'), full_page=True)
                            (diagnostico / f'consulta_{numero}.txt').write_text(page.locator('body').inner_text(), encoding='utf-8')
                        except Exception:
                            pass
                        print(f'Falha {numero}/{len(pendentes)}: {erro}. CPF pendente; diagnóstico salvo em saida/diagnostico.')
                        if falhas >= 3 or sessao_expirada(page) or isinstance(erro, (PermissionError, FileNotFoundError)):
                            pedir_confirmacao('Execução interrompida após falhas. Confira a tela. Pressione Enter para fechar: ')
                            break
                        print('Seguindo para o próximo CPF. A consulta com erro poderá ser retomada.')
                print(f'Execução finalizada: {len([cpf for cpf in cpfs if cpf in registros])}/{len(cpfs)} CPFs salvos; {falhas} falhas nesta execução. Resultados: saida/cartoes.csv')
        finally:
            context.close()
            browser.close()


if __name__ == '__main__':
    main()
