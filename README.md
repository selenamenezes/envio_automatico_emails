# Automação de Transito (ETL + Relatórios por E-mail) — Python

Script em Python para processar uma planilha (Excel) com dados de trânsito/pedidos, gerar arquivos de saída segmentados por **tipo** e **setor** e disparar **relatórios via e-mail** usando o Outlook.

## Visão geral
O fluxo principal do script:
1. **Carrega** a planilha Excel (sheet `RELAÇÃO_PEDIDOS`) e filtra os dados.
2. **Gera arquivos `.xlsx`** agrupados por:
   - Tipo (ex.: UTD, TÉCNICA, COMERCIAL)
   - Setor (ex.: CENTRO-OESTE, METROPOLITANO, NORTE, etc.)
3. Cria também uma versão **sem recebedor/dependências** (variação de filtro).
4. **Envia e-mails** para destinatários cadastrados no `config.py`, anexando os arquivos gerados.
5. Inclui funções opcionais para **atualizar consultas no Excel** e tratar alertas de **OCORRENCIA DE MATERIAL [OM]**.

## Estrutura esperada do projeto
- `script.py` (ou o nome do arquivo onde está esse código)
- `config.py` (**obrigatório**): contém caminhos e dados de e-mails/corpos das mensagens
- Arquivos/planilhas de entrada e pastas de saída (definidos em `config.py`)

> O script importa `config as c` e depende fortemente desse arquivo.

## Dependências
- `pandas`
- `openpyxl` (para salvar Excel; usado explicitamente em uma função)
- `pywin32` (integração com Excel/Outlook via `win32com.client`)
- Bibliotecas padrão: `ctypes`, `time`, `os`, `re`, `datetime`

Instalação (exemplo):
```bash
pip install pandas openpyxl pywin32
```

## Requisitos de ambiente
- Sistema operacional: **Windows**
- Outlook instalado e logado (envio via Outlook/COM)
- Excel (opcional) para a função `atualizar_transito_zero_2025()` (refresh de conexões/consultas)

## Como funciona (principais funções)

### 1) `atualizar_transito_zero_2025()`
- Atualiza conexões/consultas dentro da planilha informada em `c.PLANILHA_TRANSITO`.
- Usa `win32com` para localizar conexões por nome (ex.: `Consulta - ZMM94`, etc.).

> Esta função está comentada no fluxo principal do script.

### 2) `filtro_transito(df, tipos, setores)`
- Normaliza colunas:
  - `SETOR`, `SETOR_EXPEDIDOR`, `TIPO`, `Tipo de TR`, `ATIVO`
- Aplica filtros por combinações de:
  - `Tipo de TR` (ex.: `DE-PARA`, `EXPEDIÇÃO`, `REVERSA`, `SUCATA`, `REFORMA`, etc.)
  - `ATIVO == SIM`
  - `TIPO` (UTD/COMERCIAL/TÉCNICA)
- Salva saídas em duas pastas:
  - `c.PASTA_TRANSITO_REDE`
  - `c.PASTA_TRANSITO_BD`
- Nome do arquivo: `{tipo}_{setor}.xlsx`

### 3) `filtro_transito_sem_recebedor(df, tipos, setores)`
- Mesmo conceito do `filtro_transito`, porém salva com sufixo:
  - `{tipo}_{setor}_SEM_RECEBEDOR.xlsx`
- Usa pastas:
  - `c.PASTA_SEM_DEP_REDE`
  - `c.PASTA_SEM_DEP`

### 4) `filtro_om_pendente_encerramento(df, setores)` (opcional)
- (Opcional) Filtra o sheet `OM` (carregado como `DF_OM` no topo).
- Valida existência de colunas necessárias.
- Gera arquivos para encerramento de OM:
  - `ENCERRAR_OM_{setor}.xlsx`

> Está comentada no fluxo principal.

### 5) `enviar_emails(tipo_nome, emails_dict, arquivos_dict, arquivos_sem_dep_dict, cc)`
- Envia e-mail via Outlook para cada setor.
- Destinatários e anexos vêm de dicionários:
  - `emails_dict`: `{SETOR: [email1, email2, ...]}`
  - `arquivos_dict`: `{SETOR: caminho_arquivo}`
  - `arquivos_sem_dep_dict`: `{SETOR: caminho_arquivo_sem_dep}`
- Monta e envia:
  - Assunto: `[TRANSITO] {tipo_nome} - {SETOR} - {data_hoje}`
  - Corpo: `c.CORPO_TRANSITO`
  - CC: `cc` (normalmente `c.CC`)
  - Anexos: arquivo principal + arquivo “sem recebedor” (se existirem)

> Também há suporte opcional para anexar imagem se `c.CAMINHO_IMAGEM` existir.

### 6) `enviar_emails_om(emails_dict, arquivos_dict, cc)` (opcional)
- Envia alertas para OM com assunto e corpo definidos via `c.CORPO_OM`.
- Está comentada no fluxo principal.

## Fluxo principal (o que executa agora)
No final do script (bloco não comentado):
```py
filtro_transito(DF, TIPO, SETORES)
filtro_transito_sem_recebedor(DF, TIPO, SETOR_EXPEDIDOR)
time.sleep(50)
enviar_emails("UTD", UTD, ARQUIVOS_UTD, ARQUIVOS_UTD_SEM_DEP, c.CC)
time.sleep(600)
enviar_emails("TÉCNICA", TECNICA, ARQUIVOS_TECNICA, ARQUIVOS_TECNICA_SEM_DEP, c.CC)
time.sleep(600)
enviar_emails("COMERCIAL", COMERCIAL, ARQUIVOS_COMERCIAL, ARQUIVOS_COMERCIAL_SEM_DEP, c.CC)
```
Ou seja: gera arquivos + dispara e-mails para **UTD**, **TÉCNICA** e **COMERCIAL**.

## Arquivo `config.py` (o que precisa conter)
O script espera que `config.py` forneça, no mínimo, estes itens:

- Caminho da planilha:
  - `PLANILHA_TRANSITO`
- Pastas de saída:
  - `PASTA_TRANSITO_REDE`, `PASTA_TRANSITO_BD`
  - `PASTA_SEM_DEP_REDE`, `PASTA_SEM_DEP`
  - `PASTA_OM` (se usar OM)
- Destinatários e anexos por setor:
  - `CONTATO_UTD`, `CONTATO_TECNICA`, `CONTATO_COMERCIAL`
  - `ARQUIVOS_UTD`, `ARQUIVOS_TECNICA`, `ARQUIVOS_COMERCIAL`
  - `ARQUIVOS_UTD_SEM_DEP`, `ARQUIVOS_TECNICA_SEM_DEP`, `ARQUIVOS_COMERCIAL_SEM_DEP`
  - (opcional) `ARQUIVOS_OM`
- Conteúdo e configuração do e-mail:
  - `CORPO_TRANSITO`, `CORPO_OM`
  - `CC`
  - (opcional) `CAMINHO_IMAGEM`

## Como executar
1. Ajuste o `config.py` com caminhos válidos e dicionários de e-mails/anexos.
2. Garanta que a planilha em `c.PLANILHA_TRANSITO` exista e contenha as abas:
   - `RELAÇÃO_PEDIDOS`
   - (opcional) `OM`
3. Execute:
```bash
python script.py
```

## Observações / boas práticas
- O envio de e-mails usa `time.sleep(...)` para espaçar disparos.
- Se quiser ativar funções opcionais:
  - descomente `atualizar_transito_zero_2025()`
  - descomente o bloco `filtro_om_pendente_encerramento(...)` e `enviar_emails_om(...)`
- Garanta que as colunas usadas nos filtros existam no Excel.
- Verifique permissões de escrita nas pastas de saída definidas no `config.py`.

## Licença
Este projeto está **sem licença definida**.
