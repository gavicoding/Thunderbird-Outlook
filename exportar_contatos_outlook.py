"""Exporta os contatos do Outlook clássico pra arquivos .vcf (vCard), um
por conta - pra depois importar no Thunderbird usando o "Importar"
nativo dele (Catálogo de Endereços -> Ferramentas -> Importar -> vCard),
que é mais seguro do que este programa tentar escrever direto no banco de
dados do Thunderbird.

Requisitos: Outlook CLÁSSICO instalado e configurado, pywin32.
"""
from pathlib import Path

try:
    import win32com.client
    PYWIN32_DISPONIVEL = True
except ImportError:
    PYWIN32_DISPONIVEL = False

from exportar_eml import sanitizar_nome, caminho_longo
from vcard_util import contato_vazio, contato_para_vcard

OL_CONTACT_ITEM = 2
# Data "sentinela" que o Outlook usa internamente pra Birthday/Anniversary
# quando o campo nunca foi preenchido - não é uma data real, então não deve
# virar BDAY no vCard.
_ANO_SEM_DATA = 4501


def _item_para_contato(item) -> dict:
    c = contato_vazio()
    c["nome"] = getattr(item, "FirstName", "") or ""
    c["sobrenome"] = getattr(item, "LastName", "") or ""
    c["nome_exibicao"] = getattr(item, "FileAs", "") or getattr(item, "FullName", "") or ""
    c["apelido"] = getattr(item, "NickName", "") or ""

    c["email1"] = getattr(item, "Email1Address", "") or ""
    c["email2"] = getattr(item, "Email2Address", "") or ""

    c["tel_comercial"] = getattr(item, "BusinessTelephoneNumber", "") or ""
    c["tel_residencial"] = getattr(item, "HomeTelephoneNumber", "") or ""
    c["tel_celular"] = getattr(item, "MobileTelephoneNumber", "") or ""
    c["tel_fax"] = getattr(item, "BusinessFaxNumber", "") or ""

    c["empresa"] = getattr(item, "CompanyName", "") or ""
    c["departamento"] = getattr(item, "Department", "") or ""
    c["cargo"] = getattr(item, "JobTitle", "") or ""

    c["endereco_com"] = getattr(item, "BusinessAddressStreet", "") or ""
    c["cidade_com"] = getattr(item, "BusinessAddressCity", "") or ""
    c["estado_com"] = getattr(item, "BusinessAddressState", "") or ""
    c["cep_com"] = getattr(item, "BusinessAddressPostalCode", "") or ""
    c["pais_com"] = getattr(item, "BusinessAddressCountry", "") or ""

    c["endereco_res"] = getattr(item, "HomeAddressStreet", "") or ""
    c["cidade_res"] = getattr(item, "HomeAddressCity", "") or ""
    c["estado_res"] = getattr(item, "HomeAddressState", "") or ""
    c["cep_res"] = getattr(item, "HomeAddressPostalCode", "") or ""
    c["pais_res"] = getattr(item, "HomeAddressCountry", "") or ""

    c["site_com"] = getattr(item, "BusinessHomePage", "") or ""
    c["site_pessoal"] = getattr(item, "PersonalHomePage", "") or ""

    try:
        aniversario = item.Birthday
        if aniversario and aniversario.year != _ANO_SEM_DATA:
            c["aniversario"] = f"{aniversario.year:04d}-{aniversario.month:02d}-{aniversario.day:02d}"
    except Exception:
        pass

    c["notas"] = getattr(item, "Body", "") or ""
    return c


def _contas_por_entryid(namespace):
    contas = {}
    try:
        for conta in namespace.Session.Accounts:
            email = (conta.SmtpAddress or "").strip().lower()
            if not email:
                continue
            try:
                raiz = conta.DeliveryStore.GetRootFolder()
                contas[raiz.EntryID] = email
            except Exception:
                pass
    except Exception:
        pass
    return contas


def exportar_contatos_desta_maquina(pasta_saida: Path, contas_selecionadas=None) -> int:
    """contas_selecionadas: None exporta contatos de todas as contas; senão
    é um conjunto de e-mails (minúsculo) a exportar."""
    if not PYWIN32_DISPONIVEL:
        print("Pacote 'pywin32' não disponível nesta instalação.")
        return 0

    print("Conectando ao Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    contas = _contas_por_entryid(namespace)

    pasta_saida = Path(pasta_saida)
    total = 0
    selecionadas_lower = {c.lower() for c in contas_selecionadas} if contas_selecionadas is not None else None

    for pasta_raiz in namespace.Folders:
        email_conta = contas.get(pasta_raiz.EntryID, pasta_raiz.Name)
        if selecionadas_lower is not None and email_conta.lower() not in selecionadas_lower:
            continue

        pasta_contatos = None
        for sub in pasta_raiz.Folders:
            if sub.Name in ("Contacts", "Contatos") and getattr(sub, "DefaultItemType", None) == OL_CONTACT_ITEM:
                pasta_contatos = sub
                break
        if pasta_contatos is None or pasta_contatos.Items.Count == 0:
            continue

        contatos = []
        for item in pasta_contatos.Items:
            try:
                if getattr(item, "Class", None) != 40:  # olContact
                    continue
                contatos.append(_item_para_contato(item))
            except Exception as erro:
                print(f"  [AVISO] falhou ao ler um contato de '{email_conta}': {erro}")

        if not contatos:
            continue

        texto_vcards = "".join(contato_para_vcard(c) for c in contatos)
        pasta_saida.mkdir(parents=True, exist_ok=True)
        nome_arquivo = sanitizar_nome(email_conta, limite=80) + ".vcf"
        caminho_destino = caminho_longo(pasta_saida / nome_arquivo)
        caminho_destino.write_text(texto_vcards, encoding="utf-8")

        print(f"'{email_conta}': {len(contatos)} contatos -> {nome_arquivo}")
        total += len(contatos)

    print(f"\nExportação de contatos concluída: {total} contatos salvos em:\n  {pasta_saida}")
    print(
        "\nPra importar no Thunderbird: abra o Catálogo de Endereços "
        "(Ferramentas -> Catálogo de Endereços), depois "
        "Ferramentas -> Importar -> Contatos, e escolha cada arquivo .vcf "
        "dessa pasta."
    )
    return total


if __name__ == "__main__":
    pasta_saida_str = input(
        r"Pasta de saída para os contatos (.vcf) [C:\Temp\ContatosExportadosDoOutlook]: "
    ).strip()
    pasta_saida = Path(pasta_saida_str) if pasta_saida_str else Path(r"C:\Temp\ContatosExportadosDoOutlook")
    exportar_contatos_desta_maquina(pasta_saida)
