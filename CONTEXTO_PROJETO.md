# Projeto: Painel APS - Automação de Extração e Dashboard (Power BI)

## Objetivo geral
Automatizar a extração mensal de dados de financiamento da saúde (município de
São Gabriel do Oeste - MS, IBGE 500769) que hoje é feita manualmente via
download de planilhas em dois sites do Ministério da Saúde, unificar os dados
e alimentar um painel Power BI para o gestor.

Fluxo final desejado:
1. Script Python extrai dados via API (sem necessidade de abrir navegador).
2. Dados salvos em CSV em uma pasta fixa (`saida/`).
3. Power BI conecta nessa pasta (fonte "Pasta"), com atualização agendada.
4. Link do relatório publicado é compartilhado com o gestor.
5. (Futuro) Avaliar `.bat` + Task Scheduler para rodar o script
   automaticamente todo mês, e possivelmente migrar armazenamento para
   banco de dados (Postgres/RDS na AWS) se o volume crescer.

## Status atual
- Script principal: `extrair_dados.py` (na raiz do projeto).
- Pastas de saída: `./saida/` (CSVs) e `./saida/brutos_fns/` (arquivos
  brutos opcionais da planilha detalhada FNS).
- Partes 1-3 (e-Gestor resumo+pagamentos, FNS ação detalhada) **já testadas
  e funcionando** na máquina do usuário.
- Parte 4 (FNS "Planilha Detalhada" via POST) foi implementada mas **ainda
  não testada** - requer `pip install xlrd`. Pode precisar de ajuste no
  parser (`_parsear_planilha_detalhada`) dependendo do formato real do
  arquivo retornado (pode vir como HTML disfarçado de .xls, como observado
  no arquivo `PlanilhaDetalhada.xls` originalmente enviado pelo usuário).

## Parâmetros fixos do município (exemplo usado nos testes)
- UF: MS (sigla) / coUf IBGE: 50
- Município: SAO GABRIEL DO OESTE / coMunicipioIbge: 500769
- CNPJ do Fundo Municipal de Saúde: 13.659.627/0001-09 (13659627000109)
- Esfera administrativa: MUNICIPAL

## Endpoints descobertos (todos públicos, sem autenticação)

### 1) e-Gestor APS - Pagamento (GET, retorna JSON)
```
GET https://relatorioaps-prd.saude.gov.br/financiamento/pagamento
    ?unidadeGeografica=MUNICIPIO
    &coUf=50
    &coMunicipio=500769
    &nuParcelaInicio=202601   (formato AAAAMM)
    &nuParcelaFim=202612
    &tipoRelatorio=COMPLETO
```
Resposta JSON com duas listas principais:
- `resumosPlanosOrcamentarios[]`: uma linha por (parcela x plano
  orçamentário). Campos: sgUf, coUfIbge, coMunicipioIbge, noMunicipio,
  nuCompCnes, nuParcela, dsPlanoOrcamentario, dsEsferaAdministrativa,
  vlIntegral, vlAjuste, vlDesconto, vlEfetivoRepasse, vlImplantacao,
  vlAjusteImplantacao, vlDescontoImplantacao, vlTotalImplantacao.
  -> Equivale à aba "Resumo" do Excel original.
- `pagamentos[]`: uma linha por parcela, com ~80 colunas detalhadas
  (eSF, eAP, eMulti, ACS, Saúde Bucal/CEO/LRPD, classificações de
  qualidade/vínculo, população, tetos de equipes, PSE, etc.)
  -> Muito mais completo que o Excel original baixado manualmente.

Uma única chamada cobre todo o intervalo de parcelas (não precisa
iterar mês a mês).

### 2) FNS - Entidades (GET, retorna JSON)
Usado para obter o CNPJ/razão social da entidade (fundo municipal de
saúde) a partir do município. Precisa ser chamado por ano+mês
(geralmente o resultado é o mesmo todo mês, mas o script chama 1x e
reutiliza).
```
GET https://consultafns.saude.gov.br/recursos/consulta-detalhada/entidades
    ?ano=2026
    &count=10
    &estado=MS
    &mes=1            (1-12)
    &municipio=500769
    &page=1
    &tipoConsulta=2   (2 = Fundo a Fundo)
```
Resposta: `resultado.dados[0]` contém: uf, municipio, cpfCnpj
(13659627000109), razaoSocial, codigoMunicipioIBGE, esferaAdministrativa,
cpfCnpjFormatado (13.659.627/0001-09), etc.

### 3) FNS - Ação Detalhada (GET, retorna JSON, paginado)
```
GET https://consultafns.saude.gov.br/recursos/consulta-detalhada/detalhe-acao
    ?ano=2026
    &count=10
    &cpfCnpjUg=13659627000109
    &estado=MS
    &municipio=500769
    &page=1            (iterar até totalPaginas)
    &tipoConsulta=2
```
Resposta: `resultado.dados[]` com cada "ação" (id, descricao,
formaRepasse, quantidadeParcelas, componenteBloco.nome, blocoPacto.nome,
grupoAcao.nome, valorTotal, valorDescontoTotal, valorLiquido).
Também vem `resultado.totalPaginas` para controlar a paginação.
Em um teste real (município MS/500769, ano 2026), totalPaginas = 3
(25 ações no total, count=10/página).

Esse endpoint precisa ser chamado **mês a mês** (filtro ano+mes).

### 4) FNS - Planilha Detalhada (POST, retorna arquivo - NÃO TESTADO)
```
POST https://consultafns.saude.gov.br/recursos/consulta-detalhada/planilha-detalhada/
Content-Type: application/json

Body (exemplo real capturado via DevTools):
{
  "ano": "2026",
  "coAcao": "",
  "coBloco": "",
  "coComponente": "",
  "coGrupoAcao": "",
  "coMesAno": 1,                     // mês (1-12)
  "coMunicipioIbge": "500769",
  "coPlanoOrcamentario": "",
  "dtFinalOb": "",
  "dtInicioOb": "",
  "formaRepasse": "",
  "noMunicipio": "SAO GABRIEL DO OESTE",
  "noRazaoSocial": "FUNDO MUNICIPAL DE SAUDE DE SAO GABRIEL DO OESTE",
  "nuCnpj": "13.659.627/0001-09",
  "nuCpfCnpjUg": "13659627000109",
  "sgUf": "MS",
  "tipoConsulta": 2
}
```
GET nessa URL retorna 405 (Method Not Allowed) - confirmado que é POST.
A resposta provavelmente é um arquivo `.xls` no formato JasperReports
(mesmo padrão do arquivo `PlanilhaDetalhada.xls` enviado originalmente),
que internamente é um "Composite Document File" lido via `xlrd`.

Estrutura observada no arquivo de exemplo (`PlanilhaDetalhada.xls`):
- Linhas 0-6: cabeçalho com metadados (Município/UF, Mês/Ano, IBGE,
  CPF/CNPJ, Entidade).
- Linha 7: cabeçalho da tabela com colunas: Bloco, Grupo, Ação Detalhada,
  Competência/Parcela, Nº OB, Data OB, Banco OB, Agência OB, Conta OB,
  Valor Total, Desconto, Valor Líquido, Observação, Processo, Tipo
  Repasse, Nº Proposta.
- Linhas seguintes: dados (uma linha por OB/lançamento), até linha
  "Total Geral".

O script já tem uma função `_parsear_planilha_detalhada()` que tenta
fazer esse parsing via `xlrd.open_workbook(file_contents=...)`,
localizando a linha de cabeçalho pela presença de "Bloco" e
"Ação Detalhada". **PRECISA SER VALIDADO** com uma resposta real -
se o arquivo vier como HTML em vez de binário OLE, o parser vai
precisar ser adaptado (provavelmente usar `pandas.read_html` ou
BeautifulSoup em vez de xlrd).

## Próximos passos sugeridos (ordem de prioridade)
1. Instalar `xlrd` (`pip install xlrd`) e testar a Parte 4
   (`extrair_fns_planilha_detalhada`). Se der erro no parsing, salvar o
   arquivo bruto (`salvar_arquivo_bruto=True`) e inspecionar o conteúdo
   real para ajustar o parser.
2. Validar os CSVs gerados nas 4 partes (conferir nomes de colunas,
   tipos de dados, valores nulos).
3. Criar um `.bat` que ativa o ambiente Python (se houver venv) e roda
   `extrair_dados.py` com os parâmetros do mês corrente (ex: calcular
   automaticamente `--ano` e `--parcela-inicio/--parcela-fim` com base
   na data atual).
4. Agendar esse `.bat` no Task Scheduler do Windows (mensal).
5. Montar o modelo no Power BI Desktop:
   - Fonte "Pasta" apontando para `saida/`.
   - Relacionamentos entre as 4 tabelas (chave comum: ano/mes/parcela +
     município).
   - Medidas DAX para os KPIs (ver seção "Indicadores sugeridos" abaixo).
6. Publicar no Power BI Service e compartilhar o link com o gestor.
7. (Opcional/futuro) Migrar armazenamento para banco de dados (Postgres
   ou S3+Athena na AWS) se for necessário histórico de múltiplos
   municípios/anos, com upsert para evitar duplicidade.

## Indicadores / Dashboard sugeridos (resumo da análise já feita)

### Página 1 - Visão Geral
- KPIs: total repassado no mês (vlEfetivoRepasse), variação % vs mês
  anterior, acumulado no ano, total de descontos no período.
- Gráfico de linha: evolução mensal do vlEfetivoRepasse (12 meses).
- Gráfico de composição (barras/pizza): repasse por
  `dsPlanoOrcamentario` (eSF/eAP, Saúde Bucal, eMulti, ACS, etc.)
- Destaque para meses com `vlDesconto` significativo.

### Página 2 - Detalhamento APS (de `pagamentos[]`)
- Tabela com totais por programa/mês: vlTotalEsf, vlTotalEmulti,
  vlTotalAcsDireto, vlPagamentoEsb40h (+ qualidade), vlPagamentoCeoMunicipal,
  vlPagamentoLrpdMunicipal.
- Indicadores de qualidade/vínculo (dsClassificacaoQualidadeEsfEap,
  dsClassificacaoVinculoEsfEap, dsClassificacaoQualidadeEmulti) -
  evolução mensal (BOM/ÓTIMO/SUFICIENTE/REGULAR) - impacta receita.
- Cobertura: credenciado vs homologado vs pago (qtEsfCredenciado,
  qtEsfHomologado, qtEsfTotalPgto, etc.) - identificar equipes não
  remuneradas.

### Página 3 - Recursos do FNS (de `detalhe-acao`)
- Gráfico de barras: valor líquido por `blocoPacto.nome` (Atenção
  Primária, Vigilância, Assistência Farmacêutica, MAC, Gestão do SUS).
- Tabela de ações ordenada por valor (descricao, valorTotal,
  valorDescontoTotal, valorLiquido).
- `totalBloco` (vem pré-agregado na resposta) pode ser usado direto
  para visão "ano corrente" sem re-somar.

### Página 4 - Repasses/OBs (de `planilha-detalhada`)
- Tabela detalhada (auditoria/conciliação): Nº OB, Data OB,
  Banco/Agência/Conta, Valor Total, Desconto, Valor Líquido, Ação,
  com slicers por mês e bloco.
- Gráfico opcional: valor líquido por mês de emissão da OB (verificar
  atrasos entre competência e pagamento).

### O que evitar
- Gráficos individuais para cada uma das ~80 colunas de `pagamentos`
  (muitas são metas/tetos estáticos - melhor como tabela de referência).
- Duplicar a mesma informação em formatos diferentes (ex: mesmo total
  em pizza E em barras).

## Observações técnicas gerais
- Todos os endpoints são públicos, sem autenticação/login.
- Usar User-Agent de navegador comum nos headers (alguns servidores
  bloqueiam requisições sem User-Agent).
- Adicionar pequenas pausas (`time.sleep`) entre chamadas por cortesia
  com o servidor, especialmente no loop mês a mês do FNS.
- Tratar erros por mês individualmente (try/except no loop) - meses sem
  dados ainda disponíveis (ex: mês corrente) podem retornar erro e não
  devem interromper o processamento dos demais meses.
