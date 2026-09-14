"""Importa uma árvore de e-mails exportada do Outlook (gerada por
exportar_outlook.py, já no formato mbox/.sbd nativo do Thunderbird) direto
para dentro de Mail/Local Folders de um perfil do Thunderbird - sem precisar
de nenhum complemento do Thunderbird.

Tudo entra dentro de uma pasta própria (ex: "Importado do Outlook") em
Pastas Locais, pra nunca encostar nas pastas reais que já existem no perfil.

Por padrão os arquivos são MOVIDOS (não copiados) da pasta de exportação de
origem - depois de importados, essa pasta de origem deixa de ser necessária,
e mover no mesmo disco não duplica o espaço usado (diferente de copiar).
Isso é o ideal quando exportação e importação rodam em sequência (ex:
"Fazer tudo automaticamente"). Quando a importação é um passo separado
(pasta de origem pode ter sido guardada de propósito, levada de outra
máquina, ou reaproveitada depois), use mover=False pra copiar em vez de
mover, preservando a pasta de origem intacta.
"""
import os
import shutil
import subprocess
from pathlib import Path

NOME_PASTA_RAIZ_THUNDERBIRD = "Importado do Outlook"


def caminho_longo(caminho: Path) -> Path:
    """Mesma lógica usada em exportar_eml.py / exportar_outlook.py: contorna
    o limite de 260 caracteres do Windows com o prefixo \\\\?\\."""
    if os.name != "nt":
        return caminho
    absoluto = str(caminho.resolve())
    if not absoluto.startswith("\\\\?\\"):
        absoluto = "\\\\?\\" + absoluto
    return Path(absoluto)


def thunderbird_esta_rodando() -> bool:
    try:
        resultado = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq thunderbird.exe"],
            capture_output=True, text=True, timeout=10,
        )
        return "thunderbird.exe" in resultado.stdout.lower()
    except Exception:
        return False


def _nome_destino_disponivel(pasta_local_folders: Path, nome_base: str) -> str:
    """Evita sobrescrever uma importação anterior: se o nome já existir,
    acrescenta ' (2)', ' (3)', etc."""
    def existe(nome):
        return (pasta_local_folders / nome).exists() or (pasta_local_folders / f"{nome}.sbd").exists()

    if not existe(nome_base):
        return nome_base
    contador = 2
    while existe(f"{nome_base} ({contador})"):
        contador += 1
    return f"{nome_base} ({contador})"


def _contar_arquivos(pasta: Path) -> int:
    """Contagem rápida (só listagem de diretório) pro cálculo de progresso."""
    total = 0
    for _dirpath, _dirnames, filenames in os.walk(str(caminho_longo(pasta))):
        total += len(filenames)
    return total


def _copiar_arvore(origem: Path, destino: Path, progresso=None, total=1, contador=None, mover=True):
    """Por padrão move os arquivos (não copia) - depois de importado pro
    Thunderbird, a pasta de exportação de origem não precisa mais existir
    separada, e mover no mesmo disco não duplica espaço (é praticamente uma
    renomeação, ao contrário de copiar, que precisaria do dobro do espaço em
    disco). Se mover=False, copia em vez de mover (preserva a pasta de
    origem, útil quando a importação é feita como um passo separado)."""
    if contador is None:
        contador = [0]
    transferir = shutil.move if mover else shutil.copyfile
    origem_segura = caminho_longo(origem)
    for dirpath_str, _dirnames, filenames in os.walk(str(origem_segura)):
        dirpath = Path(dirpath_str)
        relativo = dirpath.relative_to(str(origem_segura))
        pasta_destino = caminho_longo(destino / relativo)
        pasta_destino.mkdir(parents=True, exist_ok=True)
        for nome_arquivo in filenames:
            transferir(str(dirpath / nome_arquivo), str(pasta_destino / nome_arquivo))
            contador[0] += 1
            if progresso:
                progresso(contador[0], total)


def importar_para_thunderbird(pasta_origem: Path, pasta_perfil: Path, progresso=None, mover=True) -> Path:
    if thunderbird_esta_rodando():
        print("[AVISO] O Thunderbird está aberto agora. Recomendado fechar antes de importar,")
        print("        pra evitar conflito de arquivos. Continuando mesmo assim...\n")

    pasta_local_folders = pasta_perfil / "Mail" / "Local Folders"
    caminho_longo(pasta_local_folders).mkdir(parents=True, exist_ok=True)

    nome_destino = _nome_destino_disponivel(pasta_local_folders, NOME_PASTA_RAIZ_THUNDERBIRD)
    caminho_destino = pasta_local_folders / nome_destino
    if mover:
        print(f"Movendo e-mails para: {caminho_destino}")
        print("(move em vez de copiar - não duplica o espaço em disco da pasta de origem)")
    else:
        print(f"Copiando e-mails para: {caminho_destino}")
        print("(copiando - a pasta de origem é preservada intacta)")

    caminho_longo(caminho_destino).touch(exist_ok=True)
    pasta_sbd_destino = pasta_local_folders / f"{nome_destino}.sbd"
    total_arquivos = _contar_arquivos(pasta_origem) or 1
    _copiar_arvore(pasta_origem, pasta_sbd_destino, progresso, total_arquivos, mover=mover)

    print(f"\nImportação concluída! Abra o Thunderbird (nesse mesmo perfil) e veja")
    print(f"'{nome_destino}' dentro de 'Pastas Locais'.")
    print("Se a pasta não aparecer de primeira, clique com o botão direito em 'Pastas")
    print("Locais' e escolha 'Gerenciar pastas e assinaturas' (ou reinicie o Thunderbird).")
    return caminho_destino
