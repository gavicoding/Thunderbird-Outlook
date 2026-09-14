"""Importa uma árvore de arquivos .eml (gerada por exportar_eml.py) direto para
dentro do Outlook clássico do Windows, recriando a estrutura de pastas
automaticamente via automação COM.

Não usa NameSpace.OpenSharedItem porque, nos testes, essa API abre .msg mas
recusa .eml ("caminho inválido") em builds recentes do Outlook clássico.
Em vez disso, cada e-mail é reconstruído como um MailItem novo com as
propriedades certas: corpo (texto/HTML), destinatários, anexos, e - o mais
importante - a data e o remetente originais, gravados via PropertyAccessor
direto nas propriedades MAPI (PR_MESSAGE_DELIVERY_TIME, PR_SENDER_*, etc) e
com a flag MSGFLAG_UNSENT desligada, porque sem isso o Outlook sempre grava
a data de agora, como se fosse um rascunho novo sendo salvo.

Requisitos:
    - Outlook CLÁSSICO para Windows, instalado e já configurado com a conta
      de destino (o "novo Outlook" não suporta essa automação COM).
    - pip install pywin32

Uso:
    python importar_outlook.py
"""
import email
import email.policy
import email.utils
import os
import tempfile
import time
import uuid
from pathlib import Path

try:
    import win32com.client
    import pywintypes
    PYWIN32_DISPONIVEL = True
except ImportError:
    PYWIN32_DISPONIVEL = False

from exportar_eml import sanitizar_nome

OL_FOLDER_INBOX = 6
OL_MAIL_ITEM = 0

# Propriedades MAPI (proptags) usadas pra preservar data e remetente originais.
PROPTAG_MESSAGE_DELIVERY_TIME = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
PROPTAG_CLIENT_SUBMIT_TIME = "http://schemas.microsoft.com/mapi/proptag/0x00390040"
PROPTAG_MESSAGE_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x0E070003"
PROPTAG_SENDER_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001E"
PROPTAG_SENDER_EMAIL = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001E"
PROPTAG_SENT_REPRESENTING_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0042001E"
PROPTAG_SENT_REPRESENTING_EMAIL = "http://schemas.microsoft.com/mapi/proptag/0x0065001E"
PROPTAG_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
PROPTAG_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
MSGFLAG_UNSENT = 0x08


def caminho_longo(caminho: Path) -> Path:
    """Mesma lógica de exportar_eml.caminho_longo: contorna o limite de 260
    caracteres do Windows usando o prefixo \\\\?\\."""
    if os.name != "nt":
        return caminho
    absoluto = str(caminho.resolve())
    if not absoluto.startswith("\\\\?\\"):
        absoluto = "\\\\?\\" + absoluto
    return Path(absoluto)


def obter_ou_criar_subpasta(pasta_pai, nome):
    for subpasta in pasta_pai.Folders:
        if subpasta.Name == nome:
            return subpasta
    return pasta_pai.Folders.Add(nome)


NOME_MANIFESTO = ".importados_outlook.txt"

# HRESULTs que indicam que a conexão com o Outlook caiu de vez (ex: o
# processo dele reiniciou/travou no meio da importação) - nesses casos não
# adianta pular e tentar a próxima pasta, porque TODAS vão falhar do mesmo
# jeito; melhor parar na hora e deixar o usuário rodar de novo (a retomada
# por pasta/manifesto garante que não duplica nada ao continuar).
_ERROS_CONEXAO_PERDIDA = {
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE - "O servidor RPC não está disponível"
    -2147023170,  # RPC_S_CALL_FAILED - "O processo remoto do RPC foi encerrado"
    -2147417848,  # RPC_E_DISCONNECTED - "O objeto chamado foi desconectado dos clientes"
}


def _conexao_outlook_perdida(erro: Exception) -> bool:
    codigo = erro.args[0] if getattr(erro, "args", None) else None
    return codigo in _ERROS_CONEXAO_PERDIDA


def carregar_manifesto(pasta_origem: Path) -> set:
    """Lista de e-mails (.eml) já importados com sucesso numa execução
    anterior, pra permitir retomar sem duplicar depois de uma falha (ex: o
    Outlook travar/ficar indisponível no meio de uma importação grande)."""
    caminho = pasta_origem / NOME_MANIFESTO
    if not caminho.exists():
        return set()
    try:
        return set(caminho.read_text(encoding="utf-8").splitlines())
    except OSError:
        return set()


def _registrar_no_manifesto(caminho_manifesto: Path, chave: str):
    try:
        with open(caminho_manifesto, "a", encoding="utf-8") as f:
            f.write(chave + "\n")
    except OSError:
        pass


NOME_ARQUIVO_ITENS_PROBLEMATICOS = ".itens_com_falha_outlook.txt"

# Se o MESMO e-mail derrubar o Outlook mais de uma vez, não adianta ficar
# tentando de novo pra sempre (o retry automático ia ficar preso batendo
# no mesmo item) - depois desse tanto de falhas nesse item específico, ele
# é pulado de vez e a importação segue pros próximos.
LIMITE_FALHAS_MESMO_ITEM = 2


def _caminho_itens_problematicos(pasta_origem: Path) -> Path:
    return pasta_origem / NOME_ARQUIVO_ITENS_PROBLEMATICOS


def _contar_falhas_item(pasta_origem: Path, chave: str) -> int:
    caminho = _caminho_itens_problematicos(pasta_origem)
    if not caminho.exists():
        return 0
    try:
        return caminho.read_text(encoding="utf-8").splitlines().count(chave)
    except OSError:
        return 0


def _registrar_falha_item(pasta_origem: Path, chave: str):
    try:
        with open(_caminho_itens_problematicos(pasta_origem), "a", encoding="utf-8") as f:
            f.write(chave + "\n")
    except OSError:
        pass


_CARACTERES_ARRISCADOS_NOME = set('"\';,<>')


def _lista_destinatarios_segura(valor_cabecalho: str) -> str:
    """Reformata uma lista de destinatários como endereços limpos separados
    por ';'. Necessário porque nomes de exibição com aspas/formatação
    incomum (comuns em listas de distribuição antigas, ex:
    '"\\'fulano@dominio.com\\'" <fulano@dominio.com>') fazem o Outlook travar
    de verdade ao tentar `.To = <esse texto>` com dezenas de destinatários -
    confirmado ao vivo: a mesma lista, só com endereços limpos, não trava."""
    if not valor_cabecalho:
        return ""
    partes = []
    for nome, endereco in email.utils.getaddresses([valor_cabecalho]):
        if not endereco:
            continue
        if nome and not (_CARACTERES_ARRISCADOS_NOME & set(nome)):
            partes.append(f"{nome} <{endereco}>")
        else:
            partes.append(endereco)
    return "; ".join(partes)


def _criar_item_a_partir_do_eml(outlook, caminho_eml: Path, marcar_como_lido=False):
    """Lê um .eml e devolve um MailItem do Outlook já salvo, com corpo,
    anexos, data e remetente originais - pronto só para dar .Move()."""
    with open(str(caminho_longo(caminho_eml)), "rb") as f:
        dados = f.read()
    mensagem = email.message_from_bytes(dados, policy=email.policy.default)

    nome_remetente, email_remetente = email.utils.parseaddr(mensagem.get("From", ""))
    nome_remetente = nome_remetente or email_remetente or "Desconhecido"

    data_original = None
    if mensagem.get("Date"):
        try:
            data_original = email.utils.parsedate_to_datetime(mensagem["Date"])
        except (TypeError, ValueError, OverflowError):
            data_original = None

    item = outlook.CreateItem(OL_MAIL_ITEM)
    item.Subject = mensagem.get("Subject") or "(sem assunto)"
    destinatarios_to = _lista_destinatarios_segura(mensagem.get("To", ""))
    if destinatarios_to:
        item.To = destinatarios_to
    destinatarios_cc = _lista_destinatarios_segura(mensagem.get("Cc", ""))
    if destinatarios_cc:
        item.CC = destinatarios_cc

    corpo_html = None
    corpo_texto = None
    parte_html = None
    parte_texto = None
    try:
        parte_html = mensagem.get_body(preferencelist=("html",))
        if parte_html is not None:
            corpo_html = parte_html.get_content()
    except Exception:
        pass
    try:
        parte_texto = mensagem.get_body(preferencelist=("plain",))
        if parte_texto is not None:
            corpo_texto = parte_texto.get_content()
    except Exception:
        pass

    if corpo_html:
        item.HTMLBody = corpo_html
    elif corpo_texto:
        item.Body = corpo_texto

    def _coletar_partes_extras(parte, saida):
        # iter_attachments() só olha os filhos diretos do multipart/mixed e
        # trata todo multipart/related como "o corpo", pulando-o inteiro -
        # por isso imagem embutida via cid: dentro de multipart/related
        # aninhado (a estrutura mais comum de HTML com imagem inline) nunca
        # era encontrada: o cid: ficava sem anexo correspondente e a
        # imagem aparecia removida/quebrada no Outlook, mesmo com os
        # anexos de arquivo normais migrando certinho. Em vez disso,
        # percorre a árvore inteira à mão, coletando toda parte que não
        # seja o corpo já usado acima - exceto e-mail encaminhado como
        # anexo (message/rfc822), que tem payload em lista (então também
        # "parece" multipart) mas precisa virar um anexo só, sem espalhar
        # as partes internas dele.
        if parte.get_content_maintype() == "message":
            if parte is not parte_html and parte is not parte_texto:
                saida.append(parte)
            return
        if parte.is_multipart():
            for sub in parte.get_payload():
                _coletar_partes_extras(sub, saida)
            return
        if parte is not parte_html and parte is not parte_texto:
            saida.append(parte)

    arquivos_temp = []
    partes_anexos = []
    try:
        _coletar_partes_extras(mensagem, partes_anexos)
    except Exception:
        partes_anexos = []

    for parte in partes_anexos:
        # Cada anexo é tratado isoladamente: um anexo problemático (ex: e-mail
        # encaminhado como anexo, tipo incomum) não pode derrubar os outros
        # anexos nem o e-mail inteiro.
        try:
            nome_anexo = parte.get_filename()
            content_id = (parte.get("Content-ID") or "").strip("<>") or None
            eh_inline = bool(content_id) or parte.get_content_disposition() == "inline"
            dados_anexo = parte.get_content()

            if isinstance(dados_anexo, str):
                dados_anexo = dados_anexo.encode("utf-8", errors="replace")
            elif isinstance(dados_anexo, email.message.Message):
                # "Encaminhar como anexo": o anexo é outro e-mail inteiro
                # (message/rfc822), não um arquivo comum - serializa como
                # .eml em vez de tentar gravar o objeto Python direto.
                dados_anexo = dados_anexo.as_bytes()
                nome_anexo = nome_anexo or "email_encaminhado.eml"
                if not nome_anexo.lower().endswith(".eml"):
                    nome_anexo += ".eml"
            elif not isinstance(dados_anexo, (bytes, bytearray)):
                print(f"  [AVISO] anexo de tipo não reconhecido ({type(dados_anexo).__name__}) ignorado")
                continue

            if len(dados_anexo) > 50 * 1024 * 1024:
                print(f"  [AVISO] anexo '{nome_anexo}' maior que 50 MB, pulado pra evitar instabilidade")
                continue

            nome_anexo = sanitizar_nome(nome_anexo or (content_id or "anexo"), limite=100)
            # Cada anexo em sua própria subpasta temporária com nome único
            # (UUID) - preserva o nome de arquivo original (que é o que
            # aparece pro destinatário via .FileName) e ainda assim garante
            # que dois anexos com o mesmo nome genérico (ex: "image001.png",
            # comum em milhares de e-mails diferentes) nunca concorram pelo
            # mesmo caminho em disco.
            pasta_temp_anexo = Path(tempfile.gettempdir()) / f"anexo_{uuid.uuid4().hex}"
            pasta_temp_anexo.mkdir(parents=True, exist_ok=True)
            arquivo_temp = pasta_temp_anexo / nome_anexo
            arquivo_temp.write_bytes(dados_anexo)
            arquivos_temp.append(arquivo_temp)
            anexo_outlook = item.Attachments.Add(str(arquivo_temp))

            # Imagem embutida no corpo (referenciada no HTML via cid:) - marca
            # como anexo "inline" de verdade (mesmo Content-ID + oculto),
            # senão vira um anexo solto E ainda deixa o cid: quebrado.
            if eh_inline and content_id:
                try:
                    pa_anexo = anexo_outlook.PropertyAccessor
                    pa_anexo.SetProperty(PROPTAG_ATTACH_CONTENT_ID, content_id)
                    pa_anexo.SetProperty(PROPTAG_ATTACHMENT_HIDDEN, True)
                except Exception:
                    pass
        except Exception as erro:
            print(f"  [AVISO] falhou ao anexar um arquivo desta mensagem, seguindo sem ele: {erro}")

    pa = item.PropertyAccessor
    if data_original is not None:
        try:
            data_com = pywintypes.Time(data_original)
            pa.SetProperty(PROPTAG_MESSAGE_DELIVERY_TIME, data_com)
            pa.SetProperty(PROPTAG_CLIENT_SUBMIT_TIME, data_com)
        except Exception:
            pass
    if email_remetente:
        try:
            pa.SetProperty(PROPTAG_SENDER_NAME, nome_remetente)
            pa.SetProperty(PROPTAG_SENDER_EMAIL, email_remetente)
            pa.SetProperty(PROPTAG_SENT_REPRESENTING_NAME, nome_remetente)
            pa.SetProperty(PROPTAG_SENT_REPRESENTING_EMAIL, email_remetente)
        except Exception:
            pass
    try:
        flags_atuais = pa.GetProperty(PROPTAG_MESSAGE_FLAGS)
        pa.SetProperty(PROPTAG_MESSAGE_FLAGS, flags_atuais & ~MSGFLAG_UNSENT)
    except Exception:
        pass

    if marcar_como_lido:
        try:
            item.UnRead = False
        except Exception:
            pass

    item.Save()

    for arquivo_temp in arquivos_temp:
        try:
            arquivo_temp.unlink()
            arquivo_temp.parent.rmdir()
        except OSError:
            pass

    return item


def contar_emls(pasta: Path) -> int:
    """Contagem rápida (só listagem de diretório) pra saber de antemão o
    total de .eml a importar e calcular percentual de progresso."""
    total = 0
    for _dirpath, _dirnames, filenames in os.walk(str(caminho_longo(pasta))):
        total += sum(1 for nome in filenames if nome.lower().endswith(".eml"))
    return total


def importar_pasta(
    namespace, caminho_local: Path, pasta_outlook, contador,
    progresso=None, total=1, ja_importados=None, caminho_manifesto=None,
    marcar_como_lido=False,
):
    """ja_importados/caminho_manifesto permitem retomar depois de uma falha
    no meio do caminho (ex: Outlook ficar momentaneamente indisponível) sem
    duplicar o que já tinha sido importado com sucesso.

    marcar_como_lido: se True, cada e-mail importado já entra marcado como
    lido (útil pra migração em massa não lotar a caixa de entrada nova de
    não-lidos) - se False (padrão), entra com o status original."""
    if ja_importados is None:
        ja_importados = set()

    outlook = namespace.Application

    # Retomada de uma importação anterior que não tinha o manifesto por
    # arquivo (ex: parou por uma falha antes dessa proteção existir): se a
    # pasta de destino já tem pelo menos a mesma quantidade de itens que a
    # pasta de origem tem de .eml, considera a pasta inteira como já feita.
    arquivos_eml_aqui = [
        p for p in caminho_local.iterdir() if p.is_file() and p.suffix.lower() == ".eml"
    ]
    pasta_ja_completa = False
    if arquivos_eml_aqui:
        try:
            if pasta_outlook.Items.Count >= len(arquivos_eml_aqui):
                pasta_ja_completa = True
        except Exception:
            pass

    for item_local in sorted(caminho_local.iterdir()):
        if item_local.is_dir():
            try:
                subpasta_outlook = obter_ou_criar_subpasta(pasta_outlook, item_local.name)
                importar_pasta(
                    namespace, item_local, subpasta_outlook, contador,
                    progresso, total, ja_importados, caminho_manifesto,
                    marcar_como_lido,
                )
            except Exception as erro:
                if _conexao_outlook_perdida(erro):
                    raise
                print(f"  [AVISO] falhou ao processar a pasta '{item_local}', pulando pra próxima: {erro}")
        elif item_local.suffix.lower() == ".eml":
            chave = str(item_local.resolve())
            pasta_origem_raiz = caminho_manifesto.parent if caminho_manifesto else None

            if chave in ja_importados or pasta_ja_completa:
                ja_importados.add(chave)
                if caminho_manifesto:
                    _registrar_no_manifesto(caminho_manifesto, chave)
                contador[0] += 1
                if progresso:
                    progresso(contador[0], total)
                continue

            if pasta_origem_raiz and _contar_falhas_item(pasta_origem_raiz, chave) >= LIMITE_FALHAS_MESMO_ITEM:
                print(f"  [AVISO] '{item_local.name}' já derrubou o Outlook antes - pulando de vez esse e-mail.")
                ja_importados.add(chave)
                _registrar_no_manifesto(caminho_manifesto, chave)
                contador[0] += 1
                if progresso:
                    progresso(contador[0], total)
                continue

            try:
                email_item = _criar_item_a_partir_do_eml(outlook, item_local, marcar_como_lido)
                email_item.Move(pasta_outlook)
                ja_importados.add(chave)
                if caminho_manifesto:
                    _registrar_no_manifesto(caminho_manifesto, chave)
                contador[0] += 1
                if contador[0] % 50 == 0:
                    print(f"  {contador[0]} e-mails importados...")
                if contador[0] % 300 == 0:
                    # Respiro periódico pro Outlook processar a fila de COM -
                    # ajuda a evitar que ele fique sobrecarregado e caia
                    # depois de milhares de itens seguidos.
                    time.sleep(2)
            except Exception as erro:
                if _conexao_outlook_perdida(erro):
                    if pasta_origem_raiz:
                        _registrar_falha_item(pasta_origem_raiz, chave)
                    raise
                print(f"  [AVISO] falhou ao importar '{item_local}': {erro}")
            finally:
                if progresso:
                    progresso(contador[0], total)


def main():
    if not PYWIN32_DISPONIVEL:
        print("Este script precisa do pacote 'pywin32'. Instale com:\n  pip install pywin32")
        return

    print("=== IMPORTADOR EML -> OUTLOOK (automático via COM) ===\n")
    print("IMPORTANTE:")
    print("  - Só funciona com o Outlook CLÁSSICO para Windows (não o novo Outlook).")
    print("  - O Outlook precisa estar instalado e configurado com a conta de destino.\n")

    pasta_origem_str = input(
        r"Pasta com os .eml exportados (ex: C:\Temp\EmailsConvertidosOutlook): "
    ).strip()
    pasta_origem = Path(pasta_origem_str)
    if not pasta_origem.exists():
        print(f"Pasta não encontrada: {pasta_origem}")
        return

    nome_pasta_raiz = input(
        "Nome da pasta raiz a criar no Outlook [Importado Thunderbird]: "
    ).strip() or "Importado Thunderbird"

    marcar_como_lido = input(
        "Importar já marcado como lido? (s/N): "
    ).strip().lower() == "s"

    print("\nConectando ao Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    caixa_padrao = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
    pasta_raiz_outlook = obter_ou_criar_subpasta(caixa_padrao.Parent, nome_pasta_raiz)

    contador = [0]
    importar_pasta(
        namespace, pasta_origem, pasta_raiz_outlook, contador,
        marcar_como_lido=marcar_como_lido,
    )

    print(f"\nConcluído! {contador[0]} e-mails importados para a pasta '{nome_pasta_raiz}' no Outlook.")


if __name__ == "__main__":
    main()
