import unittest
from unittest.mock import MagicMock, patch
from bot import extrair_cartoes, aguardar_resultado, salvar_csv, cpf_resultado, preparar_formulario, CONSULTA
from pathlib import Path
from tempfile import TemporaryDirectory


class ExtracaoTest(unittest.TestCase):
    def test_cartao_convenio_nao_aguarda_timeout(self):
        texto = 'BMG MULTI Status Cartão: Normal Status Conta: Normal Disponível Saque: R$0,00 BMG CARD CONVENIO Status Cartão: Bloqueado Status Conta: Normal Disponí\u00advel Saque: R$12,34'
        itens = extrair_cartoes(texto)
        self.assertEqual(len(itens), 2)
        self.assertEqual(itens[1]['cartao'], 'BMG CARD CONVENIO')
        self.assertEqual(itens[1]['disponivel_saque'], '12,34')

    def test_primeira_consulta_reutiliza_formulario_limpo(self):
        page = MagicMock()
        page.url = CONSULTA
        page.locator.return_value.inner_text.return_value = 'CPF: Pesquisar'
        with patch('bot.sessao_expirada', return_value=False), patch('bot.campo_consulta'):
            preparar_formulario(page, CONSULTA)
        page.goto.assert_not_called()
        page.reload.assert_not_called()

    def test_consulta_seguinte_limpa_resultado_uma_vez(self):
        page = MagicMock()
        page.url = CONSULTA
        page.locator.return_value.inner_text.side_effect = ['Status Cartão: Normal', 'CPF: Pesquisar']
        with patch('bot.sessao_expirada', return_value=False), patch('bot.campo_consulta'):
            preparar_formulario(page, CONSULTA)
        page.goto.assert_not_called()
        page.reload.assert_called_once()

    def test_resultado_antigo_persistente_interrompe(self):
        page = MagicMock()
        page.url = CONSULTA
        page.locator.return_value.inner_text.return_value = 'Status Cartão: Normal'
        with patch('bot.sessao_expirada', return_value=False), patch('bot.campo_consulta'):
            with self.assertRaisesRegex(ValueError, 'resultado anterior'):
                preparar_formulario(page, CONSULTA)

    def test_caractere_invisivel_do_portal(self):
        itens = extrair_cartoes('BMG BENEFÍCIO Status Cartão: Normal Status Conta: Normal Disponí\u00advel Saque: R$4,72')
        self.assertEqual(itens[0]['disponivel_saque'], '4,72')

    def test_cpf_resultado_nao_confunde_com_formulario(self):
        self.assertEqual(cpf_resultado('CPF: Pesquisar CPF: 001.234.567-89 - CLIENTE'), '00123456789')
        self.assertIsNone(cpf_resultado('CPF: Pesquisar'))

    def test_resultado_de_outro_cpf_nao_e_salvo(self):
        page = MagicMock()
        page.url = 'https://dnabmg.help.com.br/'
        page.get_by_role.return_value.is_visible.return_value = False
        page.locator.return_value.inner_text.return_value = 'CPF: 001.234.567-89 - CLIENTE BMG MULTI Status Cartão: Normal Status Conta: Normal Disponível Saque: R$0,00'
        with patch('bot.time.monotonic', side_effect=range(100)):
            with self.assertRaisesRegex(TimeoutError, 'CPF do resultado'):
                aguardar_resultado(page, 10, '99999999999')

    def test_pagina_dinamica_nao_impede_salvamento(self):
        page = MagicMock()
        page.url = 'https://dnabmg.help.com.br/'
        page.get_by_role.return_value.is_visible.return_value = False
        page.locator.return_value.inner_text.side_effect = [
            f'BMG MULTI Status Cartão: Bloqueado Status Conta: Normal Disponível Saque: R$0,00 Relógio {n}'
            for n in range(20)
        ]
        with patch('bot.time.monotonic', side_effect=range(100)):
            resultado = aguardar_resultado(page, 30)
        self.assertEqual(resultado[0]['status_cartao'], 'Bloqueado')
        with TemporaryDirectory() as pasta:
            salvar_csv({'00123456789': resultado}, Path(pasta))
            self.assertIn('00123456789;BMG MULTI;Bloqueado;0,00', (Path(pasta) / 'cartoes.csv').read_text(encoding='utf-8-sig'))

    def test_varios_cartoes_e_historico(self):
        texto = '''Cartões
BMG MULTI
Status Cartão: Bloqueado
Status Conta: Normal
Disponível Saque:
R$0,00
BMG BENEFÍCIO
Status Cartão: Cartão Extraviado
Status Conta: Normal
Disponível Saque: R$138,31
BMG CARD INSS
Status Cartão: Cartão Bloqueado
Status Conta: Normal
Disponível Saque: R$326,11
Cartões - Histórico
BMG MULTI
Status Cartão: Antigo
Status Conta: Normal
Disponível Saque: R$999,00'''
        itens = extrair_cartoes(texto)
        self.assertEqual(len(itens), 3)
        self.assertEqual([x['disponivel_saque'] for x in itens], ['0,00', '138,31', '326,11'])
        self.assertEqual(itens[1]['status_cartao'], 'Cartão Extraviado')

    def test_nao_interpretar_ausencia_como_saldo_zero(self):
        with self.assertRaises(ValueError):
            extrair_cartoes('BMG MULTI Status Cartão: Bloqueado Status Conta: Normal')
        with self.assertRaises(ValueError):
            extrair_cartoes('Acesso ao Sistema')

    def test_nao_salvar_resultado_parcial_com_cartao_desconhecido(self):
        with self.assertRaises(ValueError):
            extrair_cartoes('BMG MULTI Status Cartão: Bloqueado Status Conta: Normal Disponível Saque: R$0,00 BMG OUTRO Status Cartão: Ativo Status Conta: Normal Disponível Saque: R$10,00')


if __name__ == '__main__':
    unittest.main()
