"""Utilitários compartilhados para ler/escrever contatos no formato vCard
3.0 (RFC 2426) - o formato universal que tanto Thunderbird quanto Outlook
sabem importar/exportar nativamente, usado aqui como "linguagem comum"
entre os dois, do mesmo jeito que .eml é a linguagem comum pros e-mails.

Um contato é representado como um dicionário simples com chaves fixas
(ver CAMPOS abaixo) - cada lado (Thunderbird, Outlook) só precisa saber
converter pra esse dicionário e de volta, sem precisar entender o formato
interno do outro.
"""
import re

CAMPOS = (
    "nome", "sobrenome", "nome_exibicao", "apelido",
    "email1", "email2",
    "tel_comercial", "tel_residencial", "tel_celular", "tel_fax",
    "empresa", "departamento", "cargo",
    "endereco_com", "cidade_com", "estado_com", "cep_com", "pais_com",
    "endereco_res", "cidade_res", "estado_res", "cep_res", "pais_res",
    "site_com", "site_pessoal",
    "aniversario",  # "AAAA-MM-DD" ou None
    "notas",
)


def _escapar(texto):
    if not texto:
        return ""
    return (
        str(texto)
        .replace("\\", "\\\\")
        .replace(",", "\\,")
        .replace(";", "\\;")
        .replace("\n", "\\n")
    )


def _desescapar(texto):
    resultado = []
    i = 0
    while i < len(texto):
        if texto[i] == "\\" and i + 1 < len(texto):
            proximo = texto[i + 1]
            if proximo == "n":
                resultado.append("\n")
            elif proximo in ("\\", ",", ";"):
                resultado.append(proximo)
            else:
                resultado.append(proximo)
            i += 2
        else:
            resultado.append(texto[i])
            i += 1
    return "".join(resultado)


def contato_vazio():
    return {campo: "" for campo in CAMPOS}


def contato_para_vcard(contato: dict) -> str:
    """Dicionário de contato -> texto vCard 3.0 (pronto pra gravar num .vcf)."""
    c = {**contato_vazio(), **contato}
    linhas = ["BEGIN:VCARD", "VERSION:3.0"]

    sobrenome, nome = _escapar(c["sobrenome"]), _escapar(c["nome"])
    linhas.append(f"N:{sobrenome};{nome};;;")

    nome_exibicao = c["nome_exibicao"] or f"{c['nome']} {c['sobrenome']}".strip() or c["email1"] or "Sem nome"
    linhas.append(f"FN:{_escapar(nome_exibicao)}")

    if c["apelido"]:
        linhas.append(f"NICKNAME:{_escapar(c['apelido'])}")
    if c["email1"]:
        linhas.append(f"EMAIL;TYPE=INTERNET,PREF:{_escapar(c['email1'])}")
    if c["email2"]:
        linhas.append(f"EMAIL;TYPE=INTERNET:{_escapar(c['email2'])}")
    if c["tel_comercial"]:
        linhas.append(f"TEL;TYPE=WORK,VOICE:{_escapar(c['tel_comercial'])}")
    if c["tel_residencial"]:
        linhas.append(f"TEL;TYPE=HOME,VOICE:{_escapar(c['tel_residencial'])}")
    if c["tel_celular"]:
        linhas.append(f"TEL;TYPE=CELL:{_escapar(c['tel_celular'])}")
    if c["tel_fax"]:
        linhas.append(f"TEL;TYPE=FAX:{_escapar(c['tel_fax'])}")
    if c["empresa"] or c["departamento"]:
        linhas.append(f"ORG:{_escapar(c['empresa'])};{_escapar(c['departamento'])}")
    if c["cargo"]:
        linhas.append(f"TITLE:{_escapar(c['cargo'])}")
    if any(c[k] for k in ("endereco_com", "cidade_com", "estado_com", "cep_com", "pais_com")):
        linhas.append(
            "ADR;TYPE=WORK:;;{};{};{};{};{}".format(
                _escapar(c["endereco_com"]), _escapar(c["cidade_com"]),
                _escapar(c["estado_com"]), _escapar(c["cep_com"]), _escapar(c["pais_com"]),
            )
        )
    if any(c[k] for k in ("endereco_res", "cidade_res", "estado_res", "cep_res", "pais_res")):
        linhas.append(
            "ADR;TYPE=HOME:;;{};{};{};{};{}".format(
                _escapar(c["endereco_res"]), _escapar(c["cidade_res"]),
                _escapar(c["estado_res"]), _escapar(c["cep_res"]), _escapar(c["pais_res"]),
            )
        )
    if c["site_com"]:
        linhas.append(f"URL;TYPE=WORK:{_escapar(c['site_com'])}")
    if c["site_pessoal"]:
        linhas.append(f"URL;TYPE=HOME:{_escapar(c['site_pessoal'])}")
    if c["aniversario"]:
        aniversario_limpo = c["aniversario"].replace("-", "")
        if aniversario_limpo:
            linhas.append(f"BDAY:{aniversario_limpo}")
    if c["notas"]:
        linhas.append(f"NOTE:{_escapar(c['notas'])}")

    linhas.append("END:VCARD")
    return "\r\n".join(linhas) + "\r\n"


_PADRAO_LINHA = re.compile(r"^([^:;]+)((?:;[^:;=]+=[^:;]+)*)(?:;([^:]+))?:(.*)$")


def _desdobrar_linhas_dobradas(texto: str):
    """vCard permite quebrar uma linha longa continuando na próxima com um
    espaço/tab no início (RFC 2426, 'folding') - desfaz isso antes de
    parsear linha por linha."""
    linhas_brutas = texto.replace("\r\n", "\n").split("\n")
    linhas = []
    for linha in linhas_brutas:
        if linha.startswith((" ", "\t")) and linhas:
            linhas[-1] += linha[1:]
        elif linha.strip():
            linhas.append(linha)
    return linhas


def vcard_para_contato(texto_vcard: str) -> dict:
    """Texto de UM vCard (entre BEGIN:VCARD e END:VCARD) -> dicionário de
    contato. Só entende o subconjunto de campos que contato_para_vcard()
    escreve - suficiente pro caminho Thunderbird -> vCard -> Outlook."""
    c = contato_vazio()
    for linha in _desdobrar_linhas_dobradas(texto_vcard):
        if ":" not in linha:
            continue
        chave_completa, valor = linha.split(":", 1)
        partes_chave = chave_completa.split(";")
        nome_campo = partes_chave[0].upper()
        parametros = ";".join(partes_chave[1:]).upper()
        valor = _desescapar(valor)

        if nome_campo == "N":
            campos_n = valor.split(";")
            if len(campos_n) >= 2:
                c["sobrenome"], c["nome"] = campos_n[0], campos_n[1]
        elif nome_campo == "FN":
            c["nome_exibicao"] = valor
        elif nome_campo == "NICKNAME":
            c["apelido"] = valor
        elif nome_campo == "EMAIL":
            if not c["email1"]:
                c["email1"] = valor
            elif not c["email2"]:
                c["email2"] = valor
        elif nome_campo == "TEL":
            if "FAX" in parametros:
                c["tel_fax"] = valor
            elif "CELL" in parametros:
                c["tel_celular"] = valor
            elif "HOME" in parametros:
                c["tel_residencial"] = valor
            else:
                c["tel_comercial"] = valor
        elif nome_campo == "ORG":
            campos_org = valor.split(";")
            c["empresa"] = campos_org[0] if campos_org else ""
            if len(campos_org) > 1:
                c["departamento"] = campos_org[1]
        elif nome_campo == "TITLE":
            c["cargo"] = valor
        elif nome_campo == "ADR":
            campos_adr = (valor.split(";") + [""] * 7)[:7]
            _, _, rua, cidade, estado, cep, pais = campos_adr
            if "HOME" in parametros:
                c["endereco_res"], c["cidade_res"], c["estado_res"], c["cep_res"], c["pais_res"] = (
                    rua, cidade, estado, cep, pais,
                )
            else:
                c["endereco_com"], c["cidade_com"], c["estado_com"], c["cep_com"], c["pais_com"] = (
                    rua, cidade, estado, cep, pais,
                )
        elif nome_campo == "URL":
            if "HOME" in parametros:
                c["site_pessoal"] = valor
            else:
                c["site_com"] = valor
        elif nome_campo == "BDAY":
            digitos = re.sub(r"\D", "", valor)
            if len(digitos) == 8:
                c["aniversario"] = f"{digitos[0:4]}-{digitos[4:6]}-{digitos[6:8]}"
        elif nome_campo == "NOTE":
            c["notas"] = valor
    return c


def separar_vcards(texto: str):
    """Um arquivo .vcf pode conter vários BEGIN:VCARD...END:VCARD seguidos -
    devolve a lista de blocos individuais (cada um ainda com BEGIN/END)."""
    blocos = []
    for bloco in re.split(r"(?=BEGIN:VCARD)", texto, flags=re.IGNORECASE):
        bloco = bloco.strip()
        if bloco.upper().startswith("BEGIN:VCARD"):
            blocos.append(bloco)
    return blocos
