"""
ETL - Extração de dados de financiamento da saúde (APS + FNS)
Fontes:
  1) e-Gestor APS  -> relatorioaps-prd.saude.gov.br
  2) FNS detalhada -> consultafns.saude.gov.br

Saída: arquivos CSV em ./saida/ (um por fonte/tabela), prontos para
carregar em banco de dados ou Power BI.

Uso:
    python extrair_dados.py --uf MS --co-uf 50 --municipio 500769 \
        --ano 2026 --parcela-inicio 202601 --parcela-fim 202612
"""

import argparse
import csv
import json
import sys
import time
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError

try:
    import xlrd
except ImportError:
    xlrd = None  # só é necessário para o parsing da "Planilha Detalhada" do FNS

OUT_DIR = Path(__file__).parent / "saida"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; ETL-Saude/1.0)",
    "Accept": "application/json",
}


def get_json(url, tentativas=3, espera=2):
    """Faz GET e retorna o JSON decodificado, com retries simples."""
    ultimo_erro = None
    for i in range(tentativas):
        try:
            req = Request(url, headers=HEADERS)
            with urlopen(req, timeout=30) as resp:
                data = resp.read()
                return json.loads(data)
        except (HTTPError, URLError, json.JSONDecodeError) as e:
            ultimo_erro = e
            print(f"  [aviso] tentativa {i+1}/{tentativas} falhou: {e}")
            time.sleep(espera)
    raise RuntimeError(f"Falha ao buscar {url}: {ultimo_erro}")


def post_bytes(url, payload_dict, tentativas=3, espera=2):
    """Faz POST com corpo JSON e retorna os bytes brutos da resposta
    (usado quando a resposta é um arquivo, ex: .xls)."""
    body = json.dumps(payload_dict).encode("utf-8")
    req_headers = dict(HEADERS)
    req_headers["Content-Type"] = "application/json"
    req_headers["Accept"] = "*/*"

    ultimo_erro = None
    for i in range(tentativas):
        try:
            req = Request(url, data=body, headers=req_headers, method="POST")
            with urlopen(req, timeout=60) as resp:
                return resp.read(), resp.headers.get("Content-Type", "")
        except (HTTPError, URLError) as e:
            ultimo_erro = e
            print(f"  [aviso] tentativa {i+1}/{tentativas} falhou: {e}")
            time.sleep(espera)
    raise RuntimeError(f"Falha ao buscar (POST) {url}: {ultimo_erro}")


def salvar_csv(linhas, caminho):
    """Salva lista de dicts em CSV, criando a pasta se necessário."""
    if not linhas:
        print(f"  [aviso] nenhuma linha para salvar em {caminho}")
        return
    caminho.parent.mkdir(parents=True, exist_ok=True)
    # une todas as chaves possíveis (alguns registros podem variar)
    campos = []
    for linha in linhas:
        for k in linha.keys():
            if k not in campos:
                campos.append(k)
    novo_arquivo = not caminho.exists()
    with open(caminho, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=campos)
        if novo_arquivo:
            writer.writeheader()
        writer.writerows(linhas)
    print(f"  -> {len(linhas)} linha(s) gravada(s) em {caminho}")


# ---------------------------------------------------------------------------
# 1) e-Gestor APS
# ---------------------------------------------------------------------------
def extrair_egestor(co_uf, co_municipio, parcela_inicio, parcela_fim):
    print("[e-Gestor APS] extraindo financiamento/pagamento ...")
    url = (
        "https://relatorioaps-prd.saude.gov.br/financiamento/pagamento"
        f"?unidadeGeografica=MUNICIPIO&coUf={co_uf}&coMunicipio={co_municipio}"
        f"&nuParcelaInicio={parcela_inicio}&nuParcelaFim={parcela_fim}"
        "&tipoRelatorio=COMPLETO"
    )
    dados = get_json(url)

    resumos = dados.get("resumosPlanosOrcamentarios", [])
    pagamentos = dados.get("pagamentos", [])

    salvar_csv(resumos, OUT_DIR / "egestor_resumo_planos.csv")
    salvar_csv(pagamentos, OUT_DIR / "egestor_pagamentos_detalhado.csv")
    return dados


# ---------------------------------------------------------------------------
# 2) FNS - entidade (CNPJ do fundo de saúde)
# ---------------------------------------------------------------------------
def obter_entidade(ano, mes, uf, co_municipio, tipo_consulta=2):
    print("[FNS] buscando dados da entidade ...")
    url = (
        "https://consultafns.saude.gov.br/recursos/consulta-detalhada/entidades"
        f"?ano={ano}&count=10&estado={uf}&mes={mes}"
        f"&municipio={co_municipio}&page=1&tipoConsulta={tipo_consulta}"
    )
    dados = get_json(url)
    registros = dados.get("resultado", {}).get("dados", [])
    if not registros:
        raise RuntimeError("Nenhuma entidade encontrada para os filtros informados.")
    entidade = registros[0]
    print(f"  -> Entidade: {entidade.get('razaoSocial')} (CNPJ {entidade.get('cpfCnpjFormatado')})")
    return entidade


# ---------------------------------------------------------------------------
# 4) FNS - "Planilha Detalhada" (tabela com Nº OB, Data OB, Banco, Valor Líquido...)
# ---------------------------------------------------------------------------
def extrair_fns_planilha_detalhada(ano, mes, entidade, salvar_arquivo_bruto=False):
    """Faz o POST para gerar a Planilha Detalhada do FNS e extrai as linhas
    da tabela de repasses (Bloco, Ação, Nº OB, Data OB, Valor Líquido, etc.)
    """
    print(f"[FNS] extraindo planilha detalhada - {mes:02d}/{ano} ...")
    url = "https://consultafns.saude.gov.br/recursos/consulta-detalhada/planilha-detalhada/"
    payload = {
        "ano": str(ano),
        "coAcao": "",
        "coBloco": "",
        "coComponente": "",
        "coGrupoAcao": "",
        "coMesAno": mes,
        "coMunicipioIbge": entidade["codigoMunicipioIBGE"],
        "coPlanoOrcamentario": "",
        "dtFinalOb": "",
        "dtInicioOb": "",
        "formaRepasse": "",
        "noMunicipio": entidade["municipio"],
        "noRazaoSocial": entidade["razaoSocial"],
        "nuCnpj": entidade["cpfCnpjFormatado"],
        "nuCpfCnpjUg": entidade["cpfCnpj"],
        "sgUf": entidade["uf"],
        "tipoConsulta": entidade.get("tipoConsulta", 2),
    }

    conteudo, content_type = post_bytes(url, payload)

    if salvar_arquivo_bruto:
        bruto_dir = OUT_DIR / "brutos_fns"
        bruto_dir.mkdir(parents=True, exist_ok=True)
        caminho_bruto = bruto_dir / f"planilha_detalhada_{ano}_{mes:02d}.xls"
        with open(caminho_bruto, "wb") as f:
            f.write(conteudo)

    if xlrd is None:
        print("  [aviso] biblioteca 'xlrd' não instalada (pip install xlrd) - "
              "arquivo bruto salvo, mas não processado.")
        return []

    linhas = _parsear_planilha_detalhada(conteudo, ano, mes)
    salvar_csv(linhas, OUT_DIR / "fns_planilha_detalhada.csv")
    return linhas


def _parsear_planilha_detalhada(conteudo_bytes, ano, mes):
    """Faz o parsing do .xls (formato JasperReports / xlrd) extraindo a
    tabela de OBs. Estrutura observada:
      - linhas 1-5: cabeçalho (Município/UF, Mês/Ano, IBGE, CPF/CNPJ, Entidade)
      - linha com células 'Bloco' / 'Grupo' / 'Ação Detalhada' = header da tabela
      - linhas seguintes = dados, até a primeira linha totalmente vazia
    """
    import io
    wb = xlrd.open_workbook(file_contents=conteudo_bytes)
    sh = wb.sheet_by_index(0)

    # localizar a linha de cabeçalho da tabela (procura por 'Bloco')
    header_row_idx = None
    for r in range(sh.nrows):
        valores = sh.row_values(r)
        if "Bloco" in valores and "Ação Detalhada" in valores:
            header_row_idx = r
            break

    if header_row_idx is None:
        print("  [aviso] cabeçalho da tabela não encontrado na planilha detalhada.")
        return []

    cabecalhos = sh.row_values(header_row_idx)
    # mapeia índice de coluna -> nome (ignora colunas em branco)
    colunas = {i: str(c).strip() for i, c in enumerate(cabecalhos) if str(c).strip()}

    linhas = []
    for r in range(header_row_idx + 1, sh.nrows):
        valores = sh.row_values(r)
        if not any(str(v).strip() for v in valores):
            continue  # linha vazia
        if str(valores[0]).strip().lower().startswith("total"):
            continue  # linha de total geral, não é um lançamento individual

        linha = {"ano": ano, "mes": mes}
        for idx, nome_col in colunas.items():
            linha[nome_col] = valores[idx]
        linhas.append(linha)

    return linhas



def extrair_fns_acao_detalhada(ano, mes, uf, co_municipio, cnpj, tipo_consulta=2):
    print(f"[FNS] extraindo ação detalhada - {mes:02d}/{ano} ...")
    todas_linhas = []
    pagina = 1
    while True:
        url = (
            "https://consultafns.saude.gov.br/recursos/consulta-detalhada/detalhe-acao"
            f"?ano={ano}&count=10&cpfCnpjUg={cnpj}&estado={uf}"
            f"&municipio={co_municipio}&page={pagina}&tipoConsulta={tipo_consulta}"
        )
        dados = get_json(url)
        resultado = dados.get("resultado", {})
        registros = resultado.get("dados", [])

        for r in registros:
            todas_linhas.append({
                "ano": ano,
                "mes": mes,
                "id_acao": r.get("id"),
                "descricao_acao": r.get("descricao"),
                "forma_repasse": r.get("formaRepasse"),
                "qtd_parcelas": r.get("quantidadeParcelas"),
                "componente_bloco": (r.get("componenteBloco") or {}).get("nome"),
                "bloco_pacto": (r.get("blocoPacto") or {}).get("nome"),
                "grupo_acao": (r.get("grupoAcao") or {}).get("nome"),
                "valor_total": r.get("valorTotal"),
                "valor_desconto_total": r.get("valorDescontoTotal"),
                "valor_liquido": r.get("valorLiquido"),
            })

        total_paginas = resultado.get("totalPaginas", 1)
        print(f"  -> página {pagina}/{total_paginas} ({len(registros)} registros)")
        if pagina >= total_paginas:
            break
        pagina += 1
        time.sleep(0.5)  # cortesia com o servidor

    salvar_csv(todas_linhas, OUT_DIR / "fns_acao_detalhada.csv")
    return todas_linhas


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------
def parcelas_para_meses(parcela_inicio, parcela_fim):
    """Converte intervalo AAAAMM-AAAAMM em lista de (ano, mes)."""
    ini = int(str(parcela_inicio))
    fim = int(str(parcela_fim))
    ano, mes = divmod(ini, 100)
    fim_ano, fim_mes = divmod(fim, 100)
    resultado = []
    while (ano, mes) <= (fim_ano, fim_mes):
        resultado.append((ano, mes))
        mes += 1
        if mes > 12:
            mes = 1
            ano += 1
    return resultado


def main():
    parser = argparse.ArgumentParser(description="ETL Saúde - APS + FNS")
    parser.add_argument("--uf", required=True, help="Sigla do estado, ex: MS")
    parser.add_argument("--co-uf", required=True, help="Código IBGE do estado, ex: 50")
    parser.add_argument("--municipio", required=True, help="Código IBGE do município, ex: 500769")
    parser.add_argument("--ano", required=True, type=int, help="Ano de referência para consulta FNS")
    parser.add_argument("--parcela-inicio", required=True, help="AAAAMM, ex: 202601")
    parser.add_argument("--parcela-fim", required=True, help="AAAAMM, ex: 202612")
    parser.add_argument("--tipo-consulta", default=2, type=int, help="tipoConsulta FNS (2 = Fundo a Fundo)")
    args = parser.parse_args()

    OUT_DIR.mkdir(parents=True, exist_ok=True)

    # 1) e-Gestor APS (uma chamada cobre o intervalo de parcelas todo)
    extrair_egestor(args.co_uf, args.municipio, args.parcela_inicio, args.parcela_fim)

    # 2) FNS - precisa ser por mês (a API filtra por ano+mes)
    meses = parcelas_para_meses(args.parcela_inicio, args.parcela_fim)
    entidade = None
    for ano, mes in meses:
        try:
            if entidade is None:
                entidade = obter_entidade(ano, mes, args.uf, args.municipio, args.tipo_consulta)
                entidade["tipoConsulta"] = args.tipo_consulta
            extrair_fns_acao_detalhada(ano, mes, args.uf, args.municipio, entidade["cpfCnpj"], args.tipo_consulta)
            extrair_fns_planilha_detalhada(ano, mes, entidade)
        except Exception as e:
            print(f"  [erro] {ano}-{mes:02d}: {e}", file=sys.stderr)

    print("\nConcluído. Arquivos gerados em:", OUT_DIR.resolve())


if __name__ == "__main__":
    main()
