"""Importa contatos de arquivos .vcf (gerados por
exportar_contatos_thunderbird.py) direto pro Outlook clássico, via
automação COM - cria os contatos dentro de uma pasta própria
("Contatos Migrados do Thunderbird"), sem mexer nos contatos que já
existem.

Requisitos: Outlook CLÁSSICO instalado e configurado, pywin32.
"""
import sys
import tempfile
from pathlib import Path

import pythoncom

try:
    import win32com.client
    import pywintypes
    PYWIN32_DISPONIVEL = True
except ImportError:
    PYWIN32_DISPONIVEL = False

from vcard_util import separar_vcards, vcard_para_contato

OL_FOLDER_CONTACTS = 10
OL_CONTACT_ITEM = 2

NOME_PASTA_RAIZ = "Contatos Migrados do Thunderbird"


def obter_ou_criar_subpasta(pasta_pai, nome):
    for subpasta in pasta_pai.Folders:
        if subpasta.Name == nome:
            return subpasta
    return pasta_pai.Folders.Add(nome, OL_FOLDER_CONTACTS)


def _aplicar_contato(item, c: dict):
    item.FirstName = c["nome"] or ""
    item.LastName = c["sobrenome"] or ""
    if not item.FirstName and not item.LastName:
        item.FirstName = c["nome_exibicao"] or c["email1"] or "Sem nome"
    if c["nome_exibicao"]:
        item.FileAs = c["nome_exibicao"]
    if c["apelido"]:
        item.NickName = c["apelido"]

    if c["email1"]:
        item.Email1Address = c["email1"]
    if c["email2"]:
        item.Email2Address = c["email2"]

    if c["tel_comercial"]:
        item.BusinessTelephoneNumber = c["tel_comercial"]
    if c["tel_residencial"]:
        item.HomeTelephoneNumber = c["tel_residencial"]
    if c["tel_celular"]:
        item.MobileTelephoneNumber = c["tel_celular"]
    if c["tel_fax"]:
        item.BusinessFaxNumber = c["tel_fax"]

    if c["empresa"]:
        item.CompanyName = c["empresa"]
    if c["departamento"]:
        item.Department = c["departamento"]
    if c["cargo"]:
        item.JobTitle = c["cargo"]

    if c["endereco_com"]:
        item.BusinessAddressStreet = c["endereco_com"]
    if c["cidade_com"]:
        item.BusinessAddressCity = c["cidade_com"]
    if c["estado_com"]:
        item.BusinessAddressState = c["estado_com"]
    if c["cep_com"]:
        item.BusinessAddressPostalCode = c["cep_com"]
    if c["pais_com"]:
        item.BusinessAddressCountry = c["pais_com"]

    if c["endereco_res"]:
        item.HomeAddressStreet = c["endereco_res"]
    if c["cidade_res"]:
        item.HomeAddressCity = c["cidade_res"]
    if c["estado_res"]:
        item.HomeAddressState = c["estado_res"]
    if c["cep_res"]:
        item.HomeAddressPostalCode = c["cep_res"]
    if c["pais_res"]:
        item.HomeAddressCountry = c["pais_res"]

    if c["site_com"]:
        item.BusinessHomePage = c["site_com"]
    if c["site_pessoal"]:
        item.PersonalHomePage = c["site_pessoal"]

    if c["aniversario"]:
        try:
            ano, mes, dia = (int(p) for p in c["aniversario"].split("-"))
            import datetime
            item.Birthday = pywintypes.Time(datetime.datetime(ano, mes, dia))
        except (ValueError, TypeError):
            pass

    if c["notas"]:
        item.Body = c["notas"]


def importar_pasta_vcf(pasta_origem: Path, pasta_contatos_destino, progresso=None) -> int:
    outlook = win32com.client.Dispatch("Outlook.Application")

    arquivos_vcf = sorted(Path(pasta_origem).glob("*.vcf"))
    total_contatos_arquivo = 0
    blocos_por_arquivo = {}
    for arquivo in arquivos_vcf:
        texto = arquivo.read_text(encoding="utf-8", errors="replace")
        blocos = separar_vcards(texto)
        blocos_por_arquivo[arquivo] = blocos
        total_contatos_arquivo += len(blocos)

    if total_contatos_arquivo == 0:
        print("Nenhum contato (.vcf) encontrado na pasta de origem.")
        return 0

    importados = 0
    contador = 0
    for arquivo, blocos in blocos_por_arquivo.items():
        print(f"Importando '{arquivo.name}' ({len(blocos)} contatos)...")
        for bloco in blocos:
            try:
                contato = vcard_para_contato(bloco)
                item = outlook.CreateItem(OL_CONTACT_ITEM)
                _aplicar_contato(item, contato)
                item.Save()
                item.Move(pasta_contatos_destino)
                importados += 1
            except Exception as erro:
                print(f"  [AVISO] falhou ao importar um contato: {erro}")
            finally:
                contador += 1
                if progresso:
                    progresso(contador, total_contatos_arquivo)

    print(f"\nImportação de contatos concluída: {importados} de {total_contatos_arquivo} importados.")
    return importados


def importar_contatos_desta_maquina(pasta_origem: Path, progresso=None) -> int:
    if not PYWIN32_DISPONIVEL:
        print("Pacote 'pywin32' não disponível nesta instalação.")
        return 0

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        pasta_contatos_padrao = namespace.GetDefaultFolder(OL_FOLDER_CONTACTS)
        pasta_destino = obter_ou_criar_subpasta(pasta_contatos_padrao, NOME_PASTA_RAIZ)
        return importar_pasta_vcf(pasta_origem, pasta_destino, progresso)
    finally:
        pythoncom.CoUninitialize()


def main():
    if not PYWIN32_DISPONIVEL:
        print("Este script precisa do pacote 'pywin32'. Instale com:\n  pip install pywin32")
        return

    print("=== IMPORTADOR DE CONTATOS (.vcf) -> OUTLOOK (automático via COM) ===\n")
    pasta_origem_str = input(
        r"Pasta com os .vcf exportados (ex: C:\Temp\ContatosExportados): "
    ).strip()
    pasta_origem = Path(pasta_origem_str)
    if not pasta_origem.exists():
        print(f"Pasta não encontrada: {pasta_origem}")
        return

    importar_contatos_desta_maquina(pasta_origem)


if __name__ == "__main__":
    main()
