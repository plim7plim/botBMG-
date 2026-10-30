# botBMG

Aplicativo Windows para consultar cartões no portal BMG a partir de uma lista de CPFs e exportar resultados em Excel.

## Requisitos

Windows 10/11, Python 3.12 ou superior com Tkinter e disponível no PATH, Google Chrome instalado, internet e conta autorizada no portal.

## Como usar

1. Baixe em **Code → Download ZIP** e extraia a pasta.
2. Abra **Iniciar botBMG.cmd** ou **Abrir botBMG.vbs**.
3. Clique em **Escolher planilha**. Use CSV UTF-8 ou XLSX com coluna `CPF` na primeira linha; no XLSX, a primeira aba é lida. Há um modelo vazio em `entrada/exemplo.csv`.
4. Clique em **1. Abrir Chrome para login**. Na primeira utilização, aguarde a instalação automática das dependências.
5. Faça login e resolva o reCAPTCHA manualmente no navegador.
6. Após entrar no portal, clique em **2. Login feito • Iniciar bot**.
7. Use **Baixar planilha de consultas** ao terminar para salvar o Excel.

O bot se conecta ao navegador apenas após confirmar o login, por uma porta local em 127.0.0.1. Usa perfil temporário separado do Chrome pessoal, sem preencher senhas ou resolver CAPTCHA. O portal pode recusar a autenticação mesmo com login manual.

## Entrada e resultados

Formate os CPFs como texto para preservar zeros à esquerda. A leitura remove pontuação, completa zeros e elimina duplicatas. Verifica o formato, sem verificar os dígitos verificadores.

O relatório Excel tem abas **Resumo**, **Consultas**, **Cartoes**, **Pendencias** e **Historico**, com status e disponível para saque por cartão, totais por CPF, motivos de erro e filtros. Inclui somente a lista selecionada. Pendências ficam sem valor financeiro, sem serem tratadas como saldo zero. A data mostrada é de geração do relatório: dados antigos não têm data individual de consulta e podem vir de execuções anteriores. A soma por cartão não garante saque.

## Progresso e falhas

`saida/progresso.json` guarda os resultados. Novas execuções pulam CPFs salvos, inclusive em outra planilha. Preserve o arquivo para retomar. `saida/cartoes.csv` contém os cartões e `saida/consultas.csv` registra as tentativas.

O bot valida o CPF do resultado e aguarda dados completos e estáveis. Reconhece BMG MULTI, BMG BENEFÍCIO, BMG CARD INSS e BMG CARD CONVENIO. Outros tipos, casos sem cartões e telas desconhecidas exigem revisão e ficam pendentes. Após três falhas na execução, interrompe para revisão.

A sessão informada dura 90 minutos. Há renovação manual preventiva aos 85 minutos e recuperação quando aparece a tela de acesso. **Parar** encerra o processo e o navegador de consultas, preservando resultados já salvos.

Em falhas, texto e captura da página ficam em `saida/diagnostico`. Entradas, saídas e diagnósticos contêm dados pessoais e devem ficar em local restrito, fora do Git. O repositório não inclui bases reais, resultados, credenciais ou sessões.

## Terminal

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\python bot.py entrada\clientes.xlsx --validar
.\.venv\Scripts\python bot.py entrada\clientes.xlsx
```

Usa o Chrome instalado, sem baixar Chromium. `--campo-cpf` aceita seletor CSS do campo; `--timeout` define prazo da consulta (45 segundos); `--intervalo` define pausa entre CPFs (1 segundo).

## Testes

```powershell
python -m unittest test_extracao.py
```

Testes locais cobrem extração, caracteres invisíveis, CPF do resultado e limpeza do formulário. Não substituem validação no portal autenticado. O tempo de consulta depende da resposta do portal.
