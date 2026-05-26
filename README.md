# ConcrematFVale — Automação de Ingresso de Notas de Serviço

Automação desenvolvida para o setor financeiro da **Concremat S/A** com o objetivo de eliminar o processo manual de ingresso de Notas Fiscais de Serviço no portal de gestão fiscal corporativo.

---

## Problema

O setor financeiro precisava realizar o upload de pares de arquivos (XML + PDF) de notas de serviço em um portal web, para 11 contratos distintos, registrando manualmente cada protocolo gerado em uma planilha de controle. O processo era repetitivo, suscetível a erro humano e consumia tempo operacional significativo.

## Solução

Script Python com Selenium WebDriver que automatiza todo o fluxo:

1. Autenticação no portal fiscal corporativo
2. Seleção do contrato via menu interativo no terminal
3. Localização automática dos arquivos XML e PDF no diretório configurado
4. Upload dos arquivos e preenchimento do campo de gestor responsável
5. Submissão do formulário e tratamento de alertas
6. Extração do número de protocolo gerado e registro em planilha Excel

---

## Stack

| Camada | Tecnologia |
|--------|-----------|
| Linguagem | Python 3.13 |
| Automação Web | Selenium WebDriver · ChromeDriver |
| Manipulação de Planilhas | openpyxl |
| Controle de Versão | Git / GitHub |

---

## Estrutura do Projeto

```
ConcrematFVale/
├── NotaDeServico.py     # Script principal de automação
├── dados.py             # Configurações sensíveis (não versionado — ver abaixo)
├── relatorio.xlsx       # Planilha de controle de protocolos
├── executar.bat         # Script de setup e execução para Windows
└── requirements.txt     # Dependências Python
```

> `dados.py` contém credenciais e diretórios de cada contrato. **Não está incluído no repositório por razões de segurança.** Para execução completa, o arquivo deve ser fornecido separadamente.

---

## Como Executar

### Pré-requisitos

- Python 3.13+
- Google Chrome instalado
- ChromeDriver compatível com a versão do Chrome

### Setup

```bash
# 1. Clone o repositório
git clone https://github.com/yBelloni/ConcrematFVale.git
cd ConcrematFVale

# 2. Instale as dependências
pip install -r requirements.txt

# 3. Adicione o arquivo dados.py com as configurações do ambiente
# (solicitar ao responsável)

# 4. Execute
python NotaDeServico.py
```

**Ou via Windows:** execute `executar.bat` — verifica Python 3.13, instala dependências automaticamente e inicia o script.

---

## Fluxo de Execução

```
Iniciar → Selecionar contrato (1–11)
        → Login automático no portal
        → Loop por arquivo:
            ├── Informar número da nota
            ├── Localizar XML + PDF no diretório
            ├── Upload dos arquivos no formulário
            ├── Submeter e capturar protocolo
            ├── Registrar protocolo e número da nota no Excel
            └── Continuar ou encerrar
```

---

## Segurança

- Credenciais isoladas em `dados.py`, fora do versionamento (`.gitignore`)
- Expiração de uso configurável via `efil()` para ambientes controlados
- Nenhuma credencial ou dado sensível exposto no histórico de commits




