"""Exporta o catálogo de endereços (contatos) do Thunderbird pra arquivos
.vcf (vCard), um por catálogo (ex: "Catálogo Pessoal.vcf").

Lê direto do banco SQLite que o Thunderbird usa pra guardar contatos
(abook.sqlite, e qualquer outro catálogo que o usuário tenha criado,
abook-N.sqlite) - sem depender de nenhuma biblioteca externa, só sqlite3
da biblioteca padrão do Python.

Abre o banco em modo só-leitura (não trava nem grava nada nele) e confere
se as tabelas esperadas existem antes de consultar - se o schema não bater
(versão muito diferente do Thunderbird, catálogo vazio/corrompido), avisa
e pula esse catálogo em vez de quebrar a exportação inteira.

IMPORTANTE: feche o Thunderbird antes de rodar, pra garantir que o banco
não está no meio de uma escrita.
"""
import sqlite3
from pathlib import Path

from exportar_eml import sanitizar_nome, caminho_longo
from vcard_util import contato_vazio, contato_para_vcard

# Nome da propriedade do Thunderbird -> nome do nosso campo (ver CAMPOS em
# vcard_util.py). Propriedades sem correspondência aqui são ignoradas.
MAPA_PROPRIEDADES = {
    "FirstName": "nome",
    "LastName": "sobrenome",
    "DisplayName": "nome_exibicao",
    "NickName": "apelido",
    "PrimaryEmail": "email1",
    "SecondEmail": "email2",
    "WorkPhone": "tel_comercial",
    "HomePhone": "tel_residencial",
    "CellularNumber": "tel_celular",
    "FaxNumber": "tel_fax",
    "Company": "empresa",
    "Department": "departamento",
    "JobTitle": "cargo",
    "WorkAddress": "endereco_com",
    "WorkCity": "cidade_com",
    "WorkState": "estado_com",
    "WorkZipCode": "cep_com",
    "WorkCountry": "pais_com",
    "HomeAddress": "endereco_res",
    "HomeCity": "cidade_res",
    "HomeState": "estado_res",
    "HomeZipCode": "cep_res",
    "HomeCountry": "pais_res",
    "WebPage1": "site_com",
    "WebPage2": "site_pessoal",
    "Notes": "notas",
}


def listar_catalogos(pasta_perfil: Path):
    """[(rótulo amigável, caminho)] de todo catálogo de endereços achado
    direto na raiz do perfil (abook.sqlite = catálogo padrão; qualquer
    abook-N.sqlite = catálogo extra criado pelo usuário)."""
    pasta_perfil = Path(pasta_perfil)
    catalogos = []
    padrao = pasta_perfil / "abook.sqlite"
    if padrao.exists():
        catalogos.append(("Catálogo Pessoal", padrao))
    for caminho in sorted(pasta_perfil.glob("abook-*.sqlite")):
        catalogos.append((caminho.stem, caminho))
    return catalogos


def _tabela_existe(conexao, nome_tabela) -> bool:
    linha = conexao.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (nome_tabela,)
    ).fetchone()
    return linha is not None


def _ler_contatos(caminho_sqlite: Path):
    """Lê um abook.sqlite e devolve uma lista de dicionários de contato
    (formato de vcard_util.CAMPOS). Devolve [] com aviso se o schema
    esperado (tabelas 'cards' e 'properties') não existir."""
    uri = f"file:{caminho_longo(caminho_sqlite).as_posix()}?mode=ro"
    try:
        conexao = sqlite3.connect(uri, uri=True)
    except sqlite3.Error as erro:
        print(f"  [AVISO] não foi possível abrir '{caminho_sqlite.name}': {erro}")
        return []

    try:
        if not _tabela_existe(conexao, "cards") or not _tabela_existe(conexao, "properties"):
            print(
                f"  [AVISO] '{caminho_sqlite.name}' não tem o formato esperado "
                "(tabelas 'cards'/'properties' não encontradas) - pulando."
            )
            return []

        uids_cards = [
            linha[0] for linha in conexao.execute(
                "SELECT uid FROM cards WHERE COALESCE(isMailList, 0) = 0"
            )
        ]
        if not uids_cards:
            return []

        propriedades_por_card = {}
        cursor = conexao.execute("SELECT card, name, value FROM properties")
        for uid_card, nome_prop, valor in cursor:
            if uid_card not in propriedades_por_card:
                propriedades_por_card[uid_card] = {}
            propriedades_por_card[uid_card][nome_prop] = valor

        contatos = []
        for uid_card in uids_cards:
            propriedades = propriedades_por_card.get(uid_card, {})
            if not propriedades:
                continue
            contato = contato_vazio()
            for nome_prop, valor in propriedades.items():
                campo = MAPA_PROPRIEDADES.get(nome_prop)
                if campo and valor:
                    contato[campo] = valor

            ano = propriedades.get("BirthYear")
            mes = propriedades.get("BirthMonth")
            dia = propriedades.get("BirthDay")
            if ano and mes and dia:
                try:
                    contato["aniversario"] = f"{int(ano):04d}-{int(mes):02d}-{int(dia):02d}"
                except ValueError:
                    pass

            if any(contato.values()):
                contatos.append(contato)
        return contatos
    finally:
        conexao.close()


def exportar_contatos_desta_maquina(pasta_saida: Path) -> int:
    """Procura o(s) catálogo(s) de endereços no(s) perfil(s) do Thunderbird
    desta máquina e exporta cada um pra um .vcf em pasta_saida. Devolve o
    total de contatos exportados."""
    from exportar_eml import listar_perfis_thunderbird

    perfis = listar_perfis_thunderbird()
    if not perfis:
        print("Nenhum perfil do Thunderbird encontrado nesta máquina.")
        return 0

    pasta_saida = Path(pasta_saida)
    total = 0
    for nome_perfil, caminho_perfil in perfis:
        catalogos = listar_catalogos(caminho_perfil)
        if not catalogos:
            continue
        print(f"Perfil '{nome_perfil}':")
        for rotulo, caminho_sqlite in catalogos:
            contatos = _ler_contatos(caminho_sqlite)
            if not contatos:
                print(f"  '{rotulo}': nenhum contato encontrado.")
                continue

            texto_vcards = "".join(contato_para_vcard(c) for c in contatos)
            pasta_saida.mkdir(parents=True, exist_ok=True)
            nome_arquivo = sanitizar_nome(f"{nome_perfil}_{rotulo}", limite=80) + ".vcf"
            caminho_destino = caminho_longo(pasta_saida / nome_arquivo)
            caminho_destino.write_text(texto_vcards, encoding="utf-8")

            print(f"  '{rotulo}': {len(contatos)} contatos -> {nome_arquivo}")
            total += len(contatos)

    print(f"\nExportação de contatos concluída: {total} contatos salvos em:\n  {pasta_saida}")
    return total


if __name__ == "__main__":
    pasta_saida_str = input(
        r"Pasta de saída para os contatos (.vcf) [C:\Temp\ContatosExportados]: "
    ).strip()
    pasta_saida = Path(pasta_saida_str) if pasta_saida_str else Path(r"C:\Temp\ContatosExportados")
    exportar_contatos_desta_maquina(pasta_saida)
