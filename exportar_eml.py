"""Exporta todos os e-mails de um perfil do Thunderbird para arquivos .eml,
preservando a estrutura de pastas (Inbox, Enviados, subpastas, etc).

Não depende de nenhuma biblioteca externa - só Python padrão.

Uso:
    python exportar_eml.py
"""
import configparser
import mailbox
import os
import re
import sys
from email.header import decode_header
from pathlib import Path

EXTENSOES_IGNORADAS = {
    ".msf", ".dat", ".sqlite", ".sqlite-journal", ".json",
    ".db", ".log", ".html", ".mab", ".bak",
}

CARACTERES_INVALIDOS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def listar_perfis_thunderbird():
    """Lê o profiles.ini do Thunderbird e devolve [(nome, caminho), ...]."""
    caminho_ini = Path(os.environ["APPDATA"]) / "Thunderbird" / "profiles.ini"
    perfis = []
    if not caminho_ini.exists():
        return perfis

    config = configparser.ConfigParser()
    config.read(caminho_ini, encoding="utf-8")
    for secao in config.sections():
        if not secao.startswith("Profile"):
            continue
        dados = config[secao]
        caminho = dados.get("Path")
        if not caminho:
            continue
        if dados.get("IsRelative", "1") == "1":
            caminho_completo = Path(os.environ["APPDATA"]) / "Thunderbird" / caminho
        else:
            caminho_completo = Path(caminho)
        perfis.append((dados.get("Name", secao), caminho_completo))
    return perfis


def sanitizar_nome(texto, limite=80):
    """Remove caracteres inválidos para nomes de arquivo/pasta no Windows."""
    if not texto:
        return "sem_nome"
    texto = CARACTERES_INVALIDOS.sub("_", texto).strip(" .")
    return texto[:limite] or "sem_nome"


def decodificar_cabecalho(valor):
    """Decodifica cabeçalhos MIME (ex: '=?UTF-8?Q?Relat=C3=B3rio?=' -> 'Relatório')."""
    if not valor:
        return ""
    partes = decode_header(valor)
    resultado = []
    for texto, codificacao in partes:
        if isinstance(texto, bytes):
            try:
                resultado.append(texto.decode(codificacao or "utf-8", errors="replace"))
            except (LookupError, TypeError):
                resultado.append(texto.decode("utf-8", errors="replace"))
        else:
            resultado.append(texto)
    return "".join(resultado)


def _desescapar_valor(texto):
    """Desfaz o escape de string do prefs.js (aspas e barras invertidas)."""
    return texto.replace('\\"', '"').replace("\\\\", "\\")


def carregar_mapa_contas(pasta_perfil: Path):
    """Lê o prefs.js e devolve {nome_da_pasta_da_conta: rótulo_amigável (e-mail)}.

    Necessário porque o nome de pasta que o Thunderbird usa no disco (baseado
    no hostname do servidor) não identifica a conta com clareza - contas
    diferentes no mesmo domínio podem gerar nomes como 'mail.dominio.com' e
    'mail.dominio.com-1', sem indicar qual é qual.
    """
    caminho_prefs = pasta_perfil / "prefs.js"
    mapa = {}
    if not caminho_prefs.exists():
        return mapa

    padrao = re.compile(r'^user_pref\("([^"]+)",\s*(.*)\);\s*$')
    servidores = {}
    identidades = {}
    contas = {}

    with open(caminho_prefs, "r", encoding="utf-8", errors="replace") as f:
        for linha in f:
            if not linha.startswith('user_pref("mail.'):
                continue
            correspondencia = padrao.match(linha.strip())
            if not correspondencia:
                continue
            chave, valor_bruto = correspondencia.group(1), correspondencia.group(2).strip()
            if valor_bruto.startswith('"') and valor_bruto.endswith('"'):
                valor = _desescapar_valor(valor_bruto[1:-1])
            else:
                valor = valor_bruto

            if chave.startswith("mail.server.server"):
                _, _, resto = chave.partition("mail.server.")
                servidor_id, _, campo = resto.partition(".")
                servidores.setdefault(servidor_id, {})[campo] = valor
            elif chave.startswith("mail.identity."):
                _, _, resto = chave.partition("mail.identity.")
                identidade_id, _, campo = resto.partition(".")
                if campo == "useremail":
                    identidades[identidade_id] = valor
            elif chave.startswith("mail.account."):
                _, _, resto = chave.partition("mail.account.")
                conta_id, _, campo = resto.partition(".")
                contas.setdefault(conta_id, {})[campo] = valor

    for conta in contas.values():
        servidor = servidores.get(conta.get("server", ""))
        if not servidor or "directory" not in servidor:
            continue
        nome_pasta_disco = Path(servidor["directory"]).name

        if servidor.get("type") == "none":
            rotulo = "Local Folders"
        else:
            rotulo = None
            for identidade_id in conta.get("identities", "").split(","):
                if identidade_id in identidades:
                    rotulo = identidades[identidade_id]
                    break
            rotulo = rotulo or servidor.get("userName") or servidor.get("hostname") or nome_pasta_disco

        mapa[nome_pasta_disco] = rotulo

    return mapa


def caminho_longo(caminho: Path) -> Path:
    """Contorna o limite de 260 caracteres de caminho do Windows (MAX_PATH),
    usando o prefixo \\\\?\\ suportado nativamente pela API do Windows.
    Sem isso, pastas do Thunderbird aninhadas (.sbd em vários níveis) com
    assuntos longos estouram o limite e o arquivo falha ao ser gravado."""
    if os.name != "nt":
        return caminho
    absoluto = str(caminho.resolve())
    if not absoluto.startswith("\\\\?\\"):
        absoluto = "\\\\?\\" + absoluto
    return Path(absoluto)


def parece_mbox(caminho: Path):
    """Filtra arquivos que claramente não são caixas de e-mail (índices, cache, etc)."""
    if caminho.suffix.lower() in EXTENSOES_IGNORADAS:
        return False
    if caminho.name.startswith("."):
        return False
    if not caminho.is_file() or caminho.stat().st_size == 0:
        return False
    return True


def converter_mbox_para_eml(caminho_mbox: Path, pasta_destino: Path):
    """Lê um arquivo mbox e grava cada mensagem como um .eml em pasta_destino."""
    try:
        caixa = mailbox.mbox(str(caminho_longo(caminho_mbox)))
    except Exception as erro:
        print(f"  [AVISO] não foi possível abrir '{caminho_mbox}': {erro}")
        return 0

    pasta_destino_segura = caminho_longo(pasta_destino)
    total = 0
    erros = 0
    try:
        for indice, mensagem in enumerate(caixa, start=1):
            assunto = decodificar_cabecalho(mensagem.get("subject")) or f"email_{indice}"
            nome_arquivo = f"{indice:04d}_{sanitizar_nome(assunto)}.eml"
            if total == 0:
                pasta_destino_segura.mkdir(parents=True, exist_ok=True)
            try:
                with open(pasta_destino_segura / nome_arquivo, "wb") as f:
                    f.write(mensagem.as_bytes())
                total += 1
            except OSError as erro:
                erros += 1
                print(f"  [AVISO] falhou ao gravar mensagem {indice} de '{caminho_mbox}': {erro}")
    finally:
        caixa.close()
    if erros:
        print(f"  [AVISO] {erros} mensagem(ns) não puderam ser gravadas nesta pasta")
    return total


def calcular_pasta_destino(pasta_saida: Path, raiz: Path, dirpath: Path, nome_arquivo: str, mapa_contas: dict):
    """Espelha a hierarquia de pastas do Thunderbird, trocando a pasta da conta
    (nível 1) pelo e-mail real da conta, e removendo o sufixo .sbd das demais."""
    caminho_relativo = dirpath.relative_to(raiz)
    partes = list(caminho_relativo.parts)

    if partes:
        rotulo_conta = mapa_contas.get(partes[0])
        if rotulo_conta:
            partes[0] = sanitizar_nome(rotulo_conta, limite=60)
        else:
            # conta não identificada no prefs.js: mantém rastreável (Mail/ImapMail + nome bruto)
            partes[0] = sanitizar_nome(f"{raiz.name}_{partes[0]}", limite=60)

    partes_limpas = [
        parte[:-4] if parte.lower().endswith(".sbd") else parte
        for parte in partes
    ]
    nome_pasta = sanitizar_nome(nome_arquivo, limite=100)
    return pasta_saida / Path(*partes_limpas, nome_pasta)


def _listar_arquivos_mbox(pasta_perfil: Path):
    """Pré-lista (raiz, dirpath, nome_arquivo, tamanho) de todo arquivo que
    parece mbox, sem abrir nenhum - só pra saber o volume total de antemão
    (usado pra calcular o percentual de progresso pelo tamanho em bytes)."""
    resultado = []
    for raiz in (pasta_perfil / "Mail", pasta_perfil / "ImapMail"):
        if not raiz.exists():
            continue
        for dirpath_str, _dirnames, filenames in os.walk(raiz):
            dirpath = Path(dirpath_str)
            for nome_arquivo in filenames:
                caminho = dirpath / nome_arquivo
                if parece_mbox(caminho):
                    resultado.append((raiz, dirpath, nome_arquivo, caminho.stat().st_size))
    return resultado


def percorrer_perfil(pasta_perfil: Path, pasta_saida: Path, progresso=None, contas_selecionadas=None):
    """progresso(feito_bytes, total_bytes), chamado a cada pasta concluída -
    o percentual aqui é por volume de dados, já que contar e-mails de
    antemão exigiria ler cada mbox duas vezes.

    contas_selecionadas: se não for None, é um conjunto de rótulos de conta
    (e-mails, ou "Local Folders") - só as contas dessa lista são exportadas,
    o resto é ignorado. None (padrão) exporta tudo, sem filtro."""
    mapa_contas = carregar_mapa_contas(pasta_perfil)

    if mapa_contas:
        print("Contas identificadas no perfil:")
        for rotulo in sorted(set(mapa_contas.values())):
            print(f"  - {rotulo}")
        print()
    else:
        print("[AVISO] Não foi possível ler as contas do prefs.js; pastas de conta")
        print("        desconhecidas serão nomeadas pelo diretório bruto.\n")

    arquivos = _listar_arquivos_mbox(pasta_perfil)

    if contas_selecionadas is not None:
        selecionadas_lower = {c.lower() for c in contas_selecionadas}
        arquivos_filtrados = []
        contas_puladas = set()
        for raiz, dirpath, nome_arquivo, tamanho in arquivos:
            partes = dirpath.relative_to(raiz).parts
            rotulo_bruto = partes[0] if partes else None
            rotulo_conta = (mapa_contas.get(rotulo_bruto) or rotulo_bruto) if rotulo_bruto else None
            if rotulo_conta and rotulo_conta.lower() in selecionadas_lower:
                arquivos_filtrados.append((raiz, dirpath, nome_arquivo, tamanho))
            elif rotulo_conta:
                contas_puladas.add(rotulo_conta)
        if contas_puladas:
            print("Contas não selecionadas, ignoradas nesta exportação: " + ", ".join(sorted(contas_puladas)))
            print()
        arquivos = arquivos_filtrados

    total_bytes = sum(tamanho for *_resto, tamanho in arquivos) or 1
    bytes_feitos = 0

    total_pastas = 0
    total_emails = 0

    for raiz, dirpath, nome_arquivo, tamanho in arquivos:
        caminho = dirpath / nome_arquivo
        pasta_destino = calcular_pasta_destino(pasta_saida, raiz, dirpath, nome_arquivo, mapa_contas)
        rotulo = str(pasta_destino.relative_to(pasta_saida))
        print(f"Convertendo '{rotulo}'...")

        qtd = converter_mbox_para_eml(caminho, pasta_destino)
        if qtd:
            total_pastas += 1
            total_emails += qtd
            print(f"  -> {qtd} e-mails salvos")

        bytes_feitos += tamanho
        if progresso:
            progresso(bytes_feitos, total_bytes)

    return total_pastas, total_emails


def escolher_perfil():
    perfis = listar_perfis_thunderbird()
    if perfis:
        print("Perfis do Thunderbird encontrados:")
        for i, (nome, caminho) in enumerate(perfis, start=1):
            print(f"  {i}) {nome}  ({caminho})")
        print("  0) Digitar um caminho manualmente")
        escolha = input("Escolha o número do perfil: ").strip()
        if escolha.isdigit() and 1 <= int(escolha) <= len(perfis):
            return perfis[int(escolha) - 1][1]

    caminho_digitado = input("Caminho completo da pasta do perfil do Thunderbird: ").strip()
    return Path(caminho_digitado)


def main():
    try:
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass

    print("=== EXPORTADOR THUNDERBIRD -> EML (para Outlook) ===\n")
    print("Dica: feche o Thunderbird antes de rodar, para evitar arquivos travados.\n")

    pasta_perfil = escolher_perfil()
    if not pasta_perfil.exists():
        print(f"Pasta de perfil não encontrada: {pasta_perfil}")
        return

    pasta_saida_str = input(
        r"Pasta de saída para os .eml [C:\Temp\EmailsConvertidosOutlook]: "
    ).strip()
    pasta_saida = Path(pasta_saida_str) if pasta_saida_str else Path(r"C:\Temp\EmailsConvertidosOutlook")
    pasta_saida.mkdir(parents=True, exist_ok=True)

    print(f"\nLendo perfil: {pasta_perfil}\n")
    total_pastas, total_emails = percorrer_perfil(pasta_perfil, pasta_saida)

    print(f"\nConcluído! {total_emails} e-mails em {total_pastas} pastas salvos em:\n  {pasta_saida}")
    print(
        "\nAgora você pode arrastar as subpastas para dentro do Outlook, "
        "ou rodar 'importar_outlook.py' para importar tudo automaticamente."
    )


if __name__ == "__main__":
    main()
