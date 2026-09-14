"""Exporta e-mails do Outlook clássico direto para o formato mbox nativo do
Thunderbird (arquivo sem extensão = mensagens da pasta, pasta "Nome.sbd" =
subpastas) - o inverso de exportar_eml.py.

Cada mensagem é salva pelo Outlook como .msg (formato nativo dele) e depois
convertida para .eml de verdade com a biblioteca extract-msg, que sabe
reaproveitar o cabeçalho original de transporte quando ele existe (e-mails
recebidos) ou sintetizar um cabeçalho a partir das propriedades da mensagem
(e-mails só enviados/rascunhos).
"""
import mailbox
import os
import re
import tempfile
import time
from pathlib import Path

try:
    import win32com.client
    PYWIN32_DISPONIVEL = True
except ImportError:
    PYWIN32_DISPONIVEL = False

try:
    import extract_msg
    EXTRACT_MSG_DISPONIVEL = True
except ImportError:
    EXTRACT_MSG_DISPONIVEL = False

OL_MAIL_ITEM = 43

CARACTERES_INVALIDOS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# Mesma lista de importar_outlook.py: HRESULTs que indicam que a conexão com
# o Outlook caiu de vez (ex: o processo dele reiniciou/travou no meio da
# exportação) - nesses casos não adianta seguir tentando pasta por pasta,
# porque todas vão falhar; melhor parar na hora e deixar retomar depois.
_ERROS_CONEXAO_PERDIDA = {
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE
    -2147023170,  # RPC_S_CALL_FAILED
    -2147417848,  # RPC_E_DISCONNECTED
}


def _conexao_outlook_perdida(erro: Exception) -> bool:
    codigo = erro.args[0] if getattr(erro, "args", None) else None
    return codigo in _ERROS_CONEXAO_PERDIDA


def sanitizar_nome(texto, limite=80):
    if not texto:
        return "sem_nome"
    texto = CARACTERES_INVALIDOS.sub("_", texto).strip(" .")
    return texto[:limite] or "sem_nome"


def caminho_longo(caminho: Path) -> Path:
    """Contorna o limite de 260 caracteres do Windows (mesma lógica usada em
    exportar_eml.py e importar_outlook.py)."""
    if os.name != "nt":
        return caminho
    absoluto = str(caminho.resolve())
    if not absoluto.startswith("\\\\?\\"):
        absoluto = "\\\\?\\" + absoluto
    return Path(absoluto)


def _item_para_eml_bytes(item, caminho_temp_msg: Path) -> bytes:
    caminho_seguro = caminho_longo(caminho_temp_msg)
    item.SaveAs(str(caminho_seguro), 3)  # 3 = olMSG
    with extract_msg.openMsg(str(caminho_seguro)) as msg:
        email_msg = msg.asEmailMessage()
        # asEmailMessage() às vezes deriva os cabeçalhos do bloco de
        # transporte bruto (headerText) e, pra e-mails com cabeçalho fora do
        # padrão (ex: notificações automáticas de sistema), esse bloco pode
        # não ter uma linha "Subject:" reconhecível - mesmo o assunto
        # existindo de verdade na mensagem (msg.subject sempre funciona,
        # porque lê a propriedade MAPI direto, não depende do parsing do
        # cabeçalho). Confirmado ao vivo: sem isso, o assunto some.
        if not email_msg.get("Subject") and msg.subject:
            # Alguns assuntos têm quebra de linha embutida de verdade (dado
            # malformado/injeção de cabeçalho) - o email lib rejeita CR/LF
            # cru num valor de cabeçalho, então limpa antes de atribuir.
            assunto_limpo = msg.subject.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
            email_msg["Subject"] = assunto_limpo
        return email_msg.as_bytes()


def _contar_itens_outlook(pasta_outlook) -> int:
    """Conta rápida (via Items.Count, sem abrir nenhum item) pra saber de
    antemão o volume total e calcular percentual de progresso."""
    try:
        total = pasta_outlook.Items.Count
    except Exception:
        total = 0
    for subpasta in pasta_outlook.Folders:
        total += _contar_itens_outlook(subpasta)
    return total


def _contar_mensagens_mbox_existente(caminho: Path) -> int:
    """Quantas mensagens já existem num mbox já exportado antes - usado pra
    retomar sem duplicar quando uma pasta já tinha sido totalmente
    exportada numa execução anterior que parou no meio."""
    caminho_seguro = caminho_longo(caminho)
    if not caminho_seguro.exists() or caminho_seguro.stat().st_size == 0:
        return 0
    try:
        caixa = mailbox.mbox(str(caminho_seguro))
        try:
            return len(caixa)
        finally:
            caixa.close()
    except Exception:
        return 0


def _exportar_pasta_outlook(pasta_outlook, caminho_destino: Path, caminho_temp_msg: Path, contador, progresso, total):
    """Espelha uma pasta do Outlook (e suas subpastas) na convenção de
    arquivo/`.sbd` que o Thunderbird usa em Mail/Local Folders."""
    itens_mail = [item for item in pasta_outlook.Items if getattr(item, "Class", None) == OL_MAIL_ITEM]
    subpastas = list(pasta_outlook.Folders)

    if not itens_mail and not subpastas:
        return

    caminho_destino_seguro = caminho_longo(caminho_destino)
    caminho_destino_seguro.parent.mkdir(parents=True, exist_ok=True)

    if itens_mail:
        ja_existentes = _contar_mensagens_mbox_existente(caminho_destino)
        if ja_existentes >= len(itens_mail):
            print(f"Pulando '{pasta_outlook.Name}' - já exportada antes ({ja_existentes} e-mails).")
            contador[0] += len(itens_mail)
            if progresso:
                progresso(contador[0], total)
        else:
            if ja_existentes:
                print(
                    f"'{pasta_outlook.Name}' tinha exportação parcial "
                    f"({ja_existentes}/{len(itens_mail)}) - refazendo essa pasta."
                )
                try:
                    caminho_destino_seguro.unlink()
                except OSError:
                    pass

            print(f"Exportando '{pasta_outlook.Name}' ({len(itens_mail)} e-mails)...")
            caixa = mailbox.mbox(str(caminho_destino_seguro))
            try:
                for item in itens_mail:
                    try:
                        eml_bytes = _item_para_eml_bytes(item, caminho_temp_msg)
                        caixa.add(eml_bytes)
                        contador[0] += 1
                        if contador[0] % 50 == 0:
                            print(f"  {contador[0]} e-mails exportados...")
                        if contador[0] % 300 == 0:
                            time.sleep(2)  # respiro periódico pro Outlook não sobrecarregar
                    except Exception as erro:
                        if _conexao_outlook_perdida(erro):
                            raise
                        print(f"  [AVISO] falhou ao exportar um e-mail de '{pasta_outlook.Name}': {erro}")
                    finally:
                        if progresso:
                            progresso(contador[0], total)
            finally:
                caixa.close()

    if subpastas:
        pasta_sbd = caminho_destino.parent / f"{caminho_destino.name}.sbd"
        for subpasta in subpastas:
            caminho_sub = pasta_sbd / sanitizar_nome(subpasta.Name, limite=100)
            _exportar_pasta_outlook(subpasta, caminho_sub, caminho_temp_msg, contador, progresso, total)


def _contas_por_entryid(namespace):
    """Mapa {EntryID da pasta raiz da conta: e-mail}, pra nomear cada conta
    do Outlook pelo endereço real em vez do nome de exibição."""
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


def exportar_outlook_desta_maquina(pasta_saida: Path, progresso=None, contas_selecionadas=None) -> int:
    """progresso(feito, total) em número de e-mails - a contagem total vem
    de Items.Count por pasta (rápido, sem abrir nenhum item).

    contas_selecionadas: se não for None, é um conjunto de e-mails (minúsculo
    ou não) - só as contas do Outlook cujo e-mail estiver nessa lista são
    exportadas, o resto é ignorado. None (padrão) exporta tudo, sem filtro."""
    if not PYWIN32_DISPONIVEL:
        print("Pacote 'pywin32' não disponível nesta instalação.")
        return 0
    if not EXTRACT_MSG_DISPONIVEL:
        print("Pacote 'extract-msg' não disponível nesta instalação.")
        return 0

    print("Conectando ao Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    try:
        print(f"Versão do Outlook detectada: {outlook.Version}")
    except Exception:
        pass
    namespace = outlook.GetNamespace("MAPI")
    contas = _contas_por_entryid(namespace)

    pastas_raiz = list(namespace.Folders)
    if contas_selecionadas is not None:
        selecionadas_lower = {c.lower() for c in contas_selecionadas}
        pastas_filtradas = []
        contas_puladas = []
        for pasta_raiz in pastas_raiz:
            email_conta = contas.get(pasta_raiz.EntryID)
            if email_conta and email_conta.lower() in selecionadas_lower:
                pastas_filtradas.append(pasta_raiz)
            else:
                contas_puladas.append(email_conta or pasta_raiz.Name)
        if contas_puladas:
            print("Contas não selecionadas, ignoradas nesta exportação: " + ", ".join(contas_puladas))
        pastas_raiz = pastas_filtradas

    print("Contando e-mails (pra estimar o progresso)...")
    total_itens = sum(_contar_itens_outlook(pasta_raiz) for pasta_raiz in pastas_raiz) or 1

    contador = [0]
    with tempfile.TemporaryDirectory(prefix="msg_tmp_") as pasta_temp:
        caminho_temp_msg = Path(pasta_temp) / "temp.msg"

        for pasta_raiz in pastas_raiz:
            rotulo_conta = sanitizar_nome(
                contas.get(pasta_raiz.EntryID, pasta_raiz.Name), limite=60
            )
            print(f"\n--- Conta: {rotulo_conta} ---")

            caminho_conta = pasta_saida / rotulo_conta
            caminho_conta_segura = caminho_longo(caminho_conta)
            caminho_conta_segura.parent.mkdir(parents=True, exist_ok=True)
            if not caminho_conta_segura.exists():
                caminho_conta_segura.touch()

            pasta_sbd_conta = pasta_saida / f"{rotulo_conta}.sbd"
            for subpasta in pasta_raiz.Folders:
                caminho_destino = pasta_sbd_conta / sanitizar_nome(subpasta.Name, limite=100)
                _exportar_pasta_outlook(subpasta, caminho_destino, caminho_temp_msg, contador, progresso, total_itens)

    print(f"\nExportação concluída: {contador[0]} e-mails salvos em:\n  {pasta_saida}")
    return contador[0]
