"""Aplicativo com janela para migrar e-mails entre Thunderbird e Outlook,
nos dois sentidos (uma aba pra cada direção).

Cada direção roda em duas etapas independentes, porque exportar e importar
podem acontecer em máquinas diferentes (ex: exportar numa máquina sem
Outlook, copiar a pasta gerada pra um pendrive/rede, e importar em outra
máquina que já tem o programa de destino):

  Aba "Thunderbird → Outlook":
    1. Exportar: lê o Thunderbird desta máquina e gera uma pasta de e-mails
       em .eml, organizada por conta, ao lado do próprio programa.
    2. Importar: pega essa pasta (desta máquina ou trazida de outra) e
       importa automaticamente para dentro do Outlook clássico.

  Aba "Outlook → Thunderbird":
    1. Exportar: lê o Outlook clássico desta máquina e gera uma pasta de
       e-mails já no formato mbox nativo do Thunderbird.
    2. Importar: pega essa pasta e copia direto para Mail/Local Folders de
       um perfil do Thunderbird desta máquina.

Quando as duas etapas de uma aba estão disponíveis na mesma máquina, o botão
"Fazer tudo automaticamente" daquela aba faz as duas em sequência.
"""
import os
import queue
import sys
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path

import pythoncom

# Em builds .exe com --windowed, não existe console: sys.stdout/stderr vêm
# como None, e qualquer print() antes de serem redirecionados quebraria o
# programa. Substitui por um destino inofensivo até a interface assumir.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

import tkinter as tk
from tkinter import filedialog

import ttkbootstrap as tb
from PIL import Image, ImageTk
from ttkbootstrap.dialogs import Messagebox
from ttkbootstrap.widgets.scrolled import ScrolledText

from exportar_eml import listar_perfis_thunderbird, percorrer_perfil, carregar_mapa_contas
from importar_outlook import (
    NOME_MANIFESTO, carregar_manifesto, contar_emls, obter_ou_criar_subpasta, importar_pasta,
    _conexao_outlook_perdida,
)
from exportar_outlook import EXTRACT_MSG_DISPONIVEL, exportar_outlook_desta_maquina
from importar_thunderbird import importar_para_thunderbird

NOME_PASTA_RAIZ_OUTLOOK = "E-mails Migrados do Thunderbird"


def pasta_do_programa() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def caminho_recurso(nome: str) -> Path:
    """Acha um arquivo empacotado junto (funciona tanto rodando o .py direto
    quanto no .exe gerado pelo PyInstaller, que extrai pra uma pasta temporária)."""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base) / nome
    return Path(__file__).resolve().parent / nome


PASTA_EXPORTACAO_PADRAO = pasta_do_programa() / "EmailsExportados"
PASTA_EXPORTACAO_OUTLOOK_PADRAO = pasta_do_programa() / "EmailsExportadosDoOutlook"
PASTA_LOGS = pasta_do_programa() / "logs"


class EscritorFila:
    """Redireciona print() para a fila que a interface consome, e também
    grava uma cópia em arquivo de log (útil pra suporte, já que builds
    --windowed não têm console pra olhar depois)."""

    def __init__(self, fila: queue.Queue, arquivo_log: Path):
        self.fila = fila
        self.arquivo_log = arquivo_log

    def write(self, texto):
        if not texto:
            return
        self.fila.put(texto)
        try:
            with open(self.arquivo_log, "a", encoding="utf-8") as f:
                f.write(texto)
        except OSError:
            pass

    def flush(self):
        pass


def perfis_thunderbird_com_dados():
    return [
        (nome, caminho)
        for nome, caminho in listar_perfis_thunderbird()
        if (caminho / "Mail").exists() or (caminho / "ImapMail").exists()
    ]


def _outlook_registrado_como(bits: int) -> bool:
    """Consulta o registro do Windows (só leitura) pra ver se o COM
    'Outlook.Application' está registrado na visão de 32 ou 64 bits."""
    import winreg

    flag = winreg.KEY_WOW64_32KEY if bits == 32 else winreg.KEY_WOW64_64KEY
    try:
        with winreg.OpenKey(
            winreg.HKEY_CLASSES_ROOT, r"Outlook.Application\CLSID", 0, winreg.KEY_READ | flag
        ):
            return True
    except OSError:
        return False


def _pista_bitness_outlook():
    """Se o Outlook instalado for de 32 bits, um programa de 64 bits nunca
    consegue conectar nele via COM (limitação do Windows, não tem contorno
    no código) - detecta esse caso específico pra dar um diagnóstico claro."""
    tem_64 = _outlook_registrado_como(64)
    tem_32 = _outlook_registrado_como(32)
    if tem_32 and not tem_64:
        return (
            "O Outlook instalado aqui parece ser de 32 bits, mas este programa é de "
            "64 bits - o Windows não deixa um se conectar no outro. Seria preciso uma "
            "versão de 32 bits deste programa."
        )
    if not tem_32 and not tem_64:
        return "Não encontrei 'Outlook.Application' registrado no Windows (nem 32 nem 64 bits)."
    return None


def outlook_classico_disponivel():
    """Devolve (disponível, detalhe). O detalhe fica vazio quando dá certo,
    e traz o motivo real quando falha - pra não ter que adivinhar à distância
    por que o Outlook não foi detectado."""
    try:
        import win32com.client
    except ImportError as erro:
        return False, f"pacote 'pywin32' indisponível: {erro}"

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        outlook.GetNamespace("MAPI")
        return True, ""
    except Exception as erro:
        detalhe = str(erro).strip()
        pista = _pista_bitness_outlook()
        if pista:
            detalhe = f"{detalhe}\n{pista}" if detalhe else pista
        return False, detalhe or "motivo desconhecido"


def _aguardar_outlook_disponivel(tempo_maximo: float = 180) -> bool:
    """Espera o Outlook voltar a responder (ex: depois de reiniciar sozinho
    por causa de um crash), checando a cada poucos segundos."""
    inicio = time.monotonic()
    while time.monotonic() - inicio < tempo_maximo:
        disponivel, _ = outlook_classico_disponivel()
        if disponivel:
            return True
        time.sleep(5)
    return False


def com_retry_automatico(funcao, *args, max_tentativas: int = 20, **kwargs):
    """Roda funcao(*args, **kwargs); se a conexão com o Outlook cair no meio
    (RPC indisponível - normalmente o Outlook travando/reiniciando sozinho
    depois de muitas operações seguidas), espera ele voltar e tenta de novo
    automaticamente, até max_tentativas vezes - inclusive se o Outlook
    demorar mais que uma espera pra voltar (nesse caso só volta e espera de
    novo, em vez de desistir na hora). Como tanto a exportação quanto a
    importação são retomáveis (pulam o que já foi feito), cada nova
    tentativa só continua de onde parou, sem duplicar nem perder nada."""
    tentativa = 1
    while True:
        try:
            return funcao(*args, **kwargs)
        except Exception as erro:
            if not _conexao_outlook_perdida(erro) or tentativa >= max_tentativas:
                raise
            print(f"\n[AVISO] O Outlook ficou indisponível (tentativa {tentativa}/{max_tentativas}).")
            print("        Aguardando ele se estabilizar pra continuar automaticamente...\n")
            if _aguardar_outlook_disponivel():
                print("Outlook disponível de novo - retomando de onde parou...\n")
            else:
                print(f"Outlook ainda não voltou (tentativa {tentativa}/{max_tentativas}) - tentando de novo...\n")
            tentativa += 1


def exportar_desta_maquina(pasta_saida: Path, progresso=None, contas_selecionadas=None) -> int:
    """Exporta direto pra pasta_saida/<e-mail da conta>/... - sempre no
    mesmo nível, mesmo com vários perfis do Thunderbird na máquina. Nunca
    aninha por nome de perfil: isso quebraria o casamento de conta por
    e-mail na hora de importar pro Outlook (o importador espera achar as
    pastas de conta direto dentro da pasta de origem).

    contas_selecionadas: se não for None, filtra pra exportar só essas
    contas (rótulos vindos de escanear_contas_thunderbird). None exporta
    todas as contas encontradas, sem filtro."""
    perfis = perfis_thunderbird_com_dados()
    if not perfis:
        print("Nenhum perfil do Thunderbird com e-mails foi encontrado nesta máquina.")
        return 0

    total_emails = 0
    for nome_perfil, caminho_perfil in perfis:
        print(f"--- Perfil: {nome_perfil} ---")
        _, qtd = percorrer_perfil(caminho_perfil, pasta_saida, progresso, contas_selecionadas)
        total_emails += qtd
        print()

    print(f"Exportação concluída: {total_emails} e-mails salvos em:\n  {pasta_saida}")
    return total_emails


def escanear_contas_thunderbird():
    """Rótulos de conta (e-mails, ou "Local Folders") encontrados em todos
    os perfis do Thunderbird desta máquina que têm e-mails - usado pra
    montar a tela de seleção de contas antes de exportar."""
    rotulos = set()
    for _nome_perfil, caminho_perfil in perfis_thunderbird_com_dados():
        mapa = carregar_mapa_contas(caminho_perfil)
        rotulos.update(mapa.values())
    return sorted(rotulos, key=str.lower)


def escanear_contas_outlook():
    """Conecta no Outlook só pra listar os e-mails das contas configuradas
    (rápido, não lê nenhuma pasta/e-mail) - usado pra montar a tela de
    seleção de contas antes de exportar. Devolve (lista_emails, detalhe_erro)."""
    try:
        import win32com.client
    except ImportError as erro:
        return [], f"pacote 'pywin32' indisponível: {erro}"

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        emails = sorted(_contas_outlook_por_email(namespace).keys())
        return emails, ""
    except Exception as erro:
        return [], str(erro).strip() or "motivo desconhecido"


def _contas_outlook_por_email(namespace):
    """Mapa {e-mail em minúsculo: pasta raiz da conta no Outlook}, usado pra
    achar em qual conta já configurada cada backup do Thunderbird deve cair."""
    contas = {}
    try:
        for conta in namespace.Session.Accounts:
            email = (conta.SmtpAddress or "").strip().lower()
            if not email:
                continue
            try:
                contas[email] = conta.DeliveryStore.GetRootFolder()
            except Exception:
                pass
    except Exception:
        pass
    return contas


def importar_para_outlook(pasta_origem: Path, progresso=None, marcar_como_lido=False) -> int:
    import win32com.client

    print("Conectando ao Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    try:
        print(f"Versão do Outlook detectada: {outlook.Version}")
    except Exception:
        pass
    namespace = outlook.GetNamespace("MAPI")
    raiz_conta_padrao = namespace.GetDefaultFolder(6).Parent
    contas_outlook = _contas_outlook_por_email(namespace)

    subpastas_conta = sorted(p for p in pasta_origem.iterdir() if p.is_dir())
    if not subpastas_conta:
        print(f"Nenhuma pasta de conta encontrada dentro de: {pasta_origem}")
        return 0

    caminho_manifesto = pasta_origem / NOME_MANIFESTO
    ja_importados = carregar_manifesto(pasta_origem)
    if ja_importados:
        print(f"Retomando: {len(ja_importados)} e-mails já importados numa execução anterior serão pulados.\n")

    total_emls = sum(contar_emls(p) for p in subpastas_conta) or 1
    contador = [0]
    for subpasta_conta in subpastas_conta:
        email_conta = subpasta_conta.name.strip().lower()
        raiz_destino = contas_outlook.get(email_conta)
        if raiz_destino is not None:
            print(f"Conta '{subpasta_conta.name}' já configurada neste Outlook - importando direto nela.")
            pasta_raiz_outlook = obter_ou_criar_subpasta(raiz_destino, NOME_PASTA_RAIZ_OUTLOOK)
            importar_pasta(
                namespace, subpasta_conta, pasta_raiz_outlook, contador,
                progresso, total_emls, ja_importados, caminho_manifesto,
                marcar_como_lido,
            )
        else:
            print(f"Conta '{subpasta_conta.name}' não configurada neste Outlook - importando na conta padrão.")
            pasta_raiz_outlook = obter_ou_criar_subpasta(raiz_conta_padrao, NOME_PASTA_RAIZ_OUTLOOK)
            pasta_conta_outlook = obter_ou_criar_subpasta(pasta_raiz_outlook, subpasta_conta.name)
            importar_pasta(
                namespace, subpasta_conta, pasta_conta_outlook, contador,
                progresso, total_emls, ja_importados, caminho_manifesto,
                marcar_como_lido,
            )

    print(f"\nImportação concluída: {contador[0]} e-mails importados, dentro de pastas")
    print(f"'{NOME_PASTA_RAIZ_OUTLOOK}' em cada conta correspondente.")
    return contador[0]


def perfil_thunderbird_destino():
    """Escolhe em qual perfil do Thunderbird escrever ao importar do Outlook.
    Se houver mais de um perfil, usa o primeiro e avisa qual foi escolhido."""
    perfis = listar_perfis_thunderbird()
    if not perfis:
        return None
    if len(perfis) > 1:
        nomes = ", ".join(nome for nome, _ in perfis)
        print(f"Vários perfis do Thunderbird encontrados ({nomes}) - usando '{perfis[0][0]}'.")
    return perfis[0][1]


def importar_desta_maquina_para_thunderbird(pasta_origem: Path, progresso=None, mover=True):
    perfil = perfil_thunderbird_destino()
    if perfil is None:
        print("Nenhum perfil do Thunderbird encontrado nesta máquina.")
        return None

    return importar_para_thunderbird(pasta_origem, perfil, progresso, mover=mover)


def _formatar_numero(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def _formatar_duracao(segundos: float) -> str:
    segundos = max(0, int(segundos))
    if segundos < 60:
        return f"{segundos}s"
    minutos, segundos = divmod(segundos, 60)
    if minutos < 60:
        return f"{minutos}min"
    horas, minutos = divmod(minutos, 60)
    return f"{horas}h{minutos:02d}min"


class App:
    def __init__(self, root: tb.Window):
        self.root = root
        self.root.title("Thunderbird ⇄ Outlook")
        self.root.geometry("760x680")
        self.root.minsize(600, 480)
        try:
            self.root.iconbitmap(str(caminho_recurso("icone.ico")))
        except Exception:
            pass

        PASTA_LOGS.mkdir(parents=True, exist_ok=True)
        self.arquivo_log = PASTA_LOGS / f"log_{datetime.now():%Y%m%d_%H%M%S}.txt"

        self.fila: queue.Queue = queue.Queue()
        self.pasta_importar_outlook = tk.StringVar(value="")
        self.pasta_importar_thunderbird = tk.StringVar(value="")
        self.marcar_lido_tb_ol = tk.BooleanVar(value=True)
        self.em_execucao = False
        self.perfis_disponiveis = []
        self.outlook_disponivel = False
        self.contas_thunderbird_disponiveis = []
        self.contas_outlook_disponiveis = []
        self.contas_selecionadas_tb_ol = None  # None = todas as contas
        self.contas_selecionadas_ol_tb = None  # None = todas as contas
        self._logo_img = None
        self._fase_inicio = time.monotonic()

        self._montar_layout()
        self._detectar_ambiente()
        self.root.after(150, self._consumir_fila)

    def _montar_layout(self):
        cabecalho = tb.Frame(self.root, padding=(24, 20, 24, 12))
        cabecalho.pack(fill="x")

        try:
            imagem = Image.open(caminho_recurso("icone.png")).resize((46, 46), Image.LANCZOS)
            self._logo_img = ImageTk.PhotoImage(imagem)
            tb.Label(cabecalho, image=self._logo_img).pack(side="left", padx=(0, 14))
        except Exception:
            pass

        bloco_titulo = tb.Frame(cabecalho)
        bloco_titulo.pack(side="left", fill="x", expand=True)
        tb.Label(bloco_titulo, text="Thunderbird ⇄ Outlook", font=("Segoe UI", 17, "bold")).pack(anchor="w")
        tb.Label(
            bloco_titulo,
            text="Migre seus e-mails automaticamente, nos dois sentidos.",
            font=("Segoe UI", 9), bootstyle="secondary",
        ).pack(anchor="w")

        barra_status = tb.Frame(self.root, padding=(24, 0, 24, 14))
        barra_status.pack(fill="x")
        self.status_tb = tb.Label(barra_status, text="● Verificando o Thunderbird...", bootstyle="secondary")
        self.status_tb.pack(anchor="w")
        self.status_ol = tb.Label(barra_status, text="● Verificando o Outlook...", bootstyle="secondary")
        self.status_ol.pack(anchor="w", pady=(2, 0))

        abas = tb.Notebook(self.root, padding=(24, 0, 24, 0))
        abas.pack(fill="x")

        aba_para_outlook = tb.Frame(abas, padding=(0, 14, 0, 6))
        aba_para_thunderbird = tb.Frame(abas, padding=(0, 14, 0, 6))
        abas.add(aba_para_outlook, text="  Thunderbird → Outlook  ")
        abas.add(aba_para_thunderbird, text="  Outlook → Thunderbird  ")

        self._montar_aba_para_outlook(aba_para_outlook)
        self._montar_aba_para_thunderbird(aba_para_thunderbird)

        barra_progresso = tb.Frame(self.root, padding=(24, 6, 24, 0))
        barra_progresso.pack(fill="x")
        self.label_fase = tb.Label(barra_progresso, text="", font=("Segoe UI", 9, "bold"))
        self.label_fase.pack(anchor="w")
        self.label_progresso = tb.Label(barra_progresso, text="", bootstyle="secondary")
        self.label_progresso.pack(anchor="w")

        self.progresso = tb.Progressbar(self.root, mode="determinate", maximum=100, bootstyle="success")
        self.progresso.pack(fill="x", padx=24, pady=(4, 12))

        area_log = tb.Labelframe(self.root, text=" Andamento ", padding=6)
        area_log.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        self.log = ScrolledText(area_log, height=10, auto_hide=True, font=("Consolas", 9))
        self.log.pack(fill="both", expand=True)
        self.log.config(state="disabled")

    def _montar_aba_para_outlook(self, aba):
        aba.columnconfigure(0, weight=1, uniform="cards")
        aba.columnconfigure(1, weight=1, uniform="cards")

        linha_contas = tb.Frame(aba)
        linha_contas.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        self.label_contas_tb_ol = tb.Label(
            linha_contas, text="Contas a migrar: verificando...", bootstyle="secondary"
        )
        self.label_contas_tb_ol.pack(side="left")
        self.btn_selecionar_contas_tb_ol = tb.Button(
            linha_contas, text="Selecionar contas...", bootstyle="secondary-outline",
            command=lambda: self._abrir_selecao_contas("tb_ol"), state="disabled",
        )
        self.btn_selecionar_contas_tb_ol.pack(side="right")

        card_exportar = tb.Labelframe(aba, text=" 1. Exportar do Thunderbird ", padding=16, bootstyle="primary")
        card_exportar.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=6)
        tb.Label(
            card_exportar,
            text="Lê o Thunderbird desta máquina e gera uma pasta de e-mails (.eml) ao lado do programa.",
            justify="left", wraplength=270,
        ).pack(anchor="w", fill="x", pady=(0, 12))
        self.btn_exportar_tb = tb.Button(
            card_exportar, text="Exportar e-mails", bootstyle="primary",
            command=self._ao_clicar_exportar_thunderbird, state="disabled",
        )
        self.btn_exportar_tb.pack(fill="x")

        card_importar = tb.Labelframe(aba, text=" 2. Importar para o Outlook ", padding=16, bootstyle="primary")
        card_importar.grid(row=1, column=1, sticky="nsew", padx=(8, 0), pady=6)
        tb.Label(
            card_importar,
            text="Pega uma pasta de e-mails exportados (desta máquina ou copiada de outra) e importa.",
            justify="left", wraplength=270,
        ).pack(anchor="w", fill="x", pady=(0, 12))
        linha_pasta = tb.Frame(card_importar)
        linha_pasta.pack(fill="x", pady=(0, 12))
        tb.Entry(linha_pasta, textvariable=self.pasta_importar_outlook).pack(side="left", fill="x", expand=True)
        tb.Button(
            linha_pasta, text="...", width=3, bootstyle="secondary-outline",
            command=lambda: self._escolher_pasta(self.pasta_importar_outlook),
        ).pack(side="left", padx=(6, 0))
        tb.Checkbutton(
            card_importar, text="Importar já marcado como lido",
            variable=self.marcar_lido_tb_ol, bootstyle="round-toggle",
        ).pack(anchor="w", pady=(0, 12))
        self.btn_importar_ol = tb.Button(
            card_importar, text="Importar e-mails", bootstyle="primary",
            command=self._ao_clicar_importar_outlook, state="disabled",
        )
        self.btn_importar_ol.pack(fill="x")

        self.btn_tudo_tb_ol = tb.Button(
            aba, text="⚡  Fazer tudo automaticamente (nesta máquina)",
            bootstyle="success", command=self._ao_clicar_tudo_thunderbird_outlook, state="disabled",
        )
        self.btn_tudo_tb_ol.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))

    def _montar_aba_para_thunderbird(self, aba):
        aba.columnconfigure(0, weight=1, uniform="cards2")
        aba.columnconfigure(1, weight=1, uniform="cards2")

        linha_contas = tb.Frame(aba)
        linha_contas.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        self.label_contas_ol_tb = tb.Label(
            linha_contas, text="Contas a migrar: verificando...", bootstyle="secondary"
        )
        self.label_contas_ol_tb.pack(side="left")
        self.btn_selecionar_contas_ol_tb = tb.Button(
            linha_contas, text="Selecionar contas...", bootstyle="secondary-outline",
            command=lambda: self._abrir_selecao_contas("ol_tb"), state="disabled",
        )
        self.btn_selecionar_contas_ol_tb.pack(side="right")

        card_exportar = tb.Labelframe(aba, text=" 1. Exportar do Outlook ", padding=16, bootstyle="primary")
        card_exportar.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=6)
        tb.Label(
            card_exportar,
            text="Lê o Outlook clássico desta máquina e gera uma pasta de e-mails ao lado do programa.",
            justify="left", wraplength=270,
        ).pack(anchor="w", fill="x", pady=(0, 12))
        self.btn_exportar_ol = tb.Button(
            card_exportar, text="Exportar e-mails", bootstyle="primary",
            command=self._ao_clicar_exportar_outlook, state="disabled",
        )
        self.btn_exportar_ol.pack(fill="x")

        card_importar = tb.Labelframe(aba, text=" 2. Importar para o Thunderbird ", padding=16, bootstyle="primary")
        card_importar.grid(row=1, column=1, sticky="nsew", padx=(8, 0), pady=6)
        tb.Label(
            card_importar,
            text="Pega uma pasta de e-mails exportados do Outlook (desta máquina ou copiada de outra) e importa.",
            justify="left", wraplength=270,
        ).pack(anchor="w", fill="x", pady=(0, 12))
        linha_pasta = tb.Frame(card_importar)
        linha_pasta.pack(fill="x", pady=(0, 12))
        tb.Entry(linha_pasta, textvariable=self.pasta_importar_thunderbird).pack(side="left", fill="x", expand=True)
        tb.Button(
            linha_pasta, text="...", width=3, bootstyle="secondary-outline",
            command=lambda: self._escolher_pasta(self.pasta_importar_thunderbird),
        ).pack(side="left", padx=(6, 0))
        self.btn_importar_tb = tb.Button(
            card_importar, text="Importar e-mails", bootstyle="primary",
            command=self._ao_clicar_importar_thunderbird, state="disabled",
        )
        self.btn_importar_tb.pack(fill="x")

        self.btn_tudo_ol_tb = tb.Button(
            aba, text="⚡  Fazer tudo automaticamente (nesta máquina)",
            bootstyle="success", command=self._ao_clicar_tudo_outlook_thunderbird, state="disabled",
        )
        self.btn_tudo_ol_tb.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        if not EXTRACT_MSG_DISPONIVEL:
            tb.Label(
                aba, text="Pacote 'extract-msg' não encontrado - exportar do Outlook fica indisponível.",
                bootstyle="warning",
            ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

    def _detectar_ambiente(self):
        def worker():
            # COM exige inicialização própria por thread - sem isso, toda
            # chamada win32com nesta thread falha com "CoInitialize não foi
            # chamado", mesmo com o Outlook instalado e aberto.
            pythoncom.CoInitialize()
            try:
                perfis = perfis_thunderbird_com_dados()
                outlook_ok, detalhe_outlook = outlook_classico_disponivel()
                contas_tb = escanear_contas_thunderbird() if perfis else []
                contas_ol, _detalhe_contas_ol = escanear_contas_outlook() if outlook_ok else ([], "")
            finally:
                pythoncom.CoUninitialize()
            self.root.after(
                0, lambda: self._atualizar_status(perfis, outlook_ok, detalhe_outlook, contas_tb, contas_ol)
            )

        threading.Thread(target=worker, daemon=True).start()

    def _atualizar_status(self, perfis, outlook_ok, detalhe_outlook="", contas_tb=None, contas_ol=None):
        self.perfis_disponiveis = perfis
        self.outlook_disponivel = outlook_ok
        self.contas_thunderbird_disponiveis = contas_tb or []
        self.contas_outlook_disponiveis = contas_ol or []

        if perfis:
            nomes = ", ".join(nome for nome, _ in perfis)
            self.status_tb.config(text=f"●  Thunderbird encontrado ({nomes})", bootstyle="success")
        else:
            self.status_tb.config(text="●  Nenhum perfil do Thunderbird encontrado nesta máquina", bootstyle="danger")

        if outlook_ok:
            self.status_ol.config(text="●  Outlook clássico disponível", bootstyle="success")
        else:
            self.status_ol.config(text="●  Outlook clássico não encontrado nesta máquina", bootstyle="danger")
            if detalhe_outlook:
                self._logar(f"[Diagnóstico] Outlook não detectado - motivo: {detalhe_outlook}\n\n")

        if PASTA_EXPORTACAO_PADRAO.exists() and not self.pasta_importar_outlook.get():
            self.pasta_importar_outlook.set(str(PASTA_EXPORTACAO_PADRAO))
        if PASTA_EXPORTACAO_OUTLOOK_PADRAO.exists() and not self.pasta_importar_thunderbird.get():
            self.pasta_importar_thunderbird.set(str(PASTA_EXPORTACAO_OUTLOOK_PADRAO))

        self._restaurar_botoes()

    def _restaurar_botoes(self):
        tem_thunderbird = bool(self.perfis_disponiveis)
        tem_outlook = self.outlook_disponivel
        tem_outlook_export = tem_outlook and EXTRACT_MSG_DISPONIVEL

        self.btn_exportar_tb.config(state="normal" if tem_thunderbird else "disabled")
        self.btn_importar_ol.config(state="normal" if tem_outlook else "disabled")
        self.btn_tudo_tb_ol.config(state="normal" if (tem_thunderbird and tem_outlook) else "disabled")

        self.btn_exportar_ol.config(state="normal" if tem_outlook_export else "disabled")
        self.btn_importar_tb.config(state="normal" if tem_thunderbird else "disabled")
        self.btn_tudo_ol_tb.config(state="normal" if (tem_outlook_export and tem_thunderbird) else "disabled")

        self.btn_selecionar_contas_tb_ol.config(state="normal" if self.contas_thunderbird_disponiveis else "disabled")
        self.btn_selecionar_contas_ol_tb.config(state="normal" if self.contas_outlook_disponiveis else "disabled")
        self._atualizar_labels_contas()

    def _atualizar_labels_contas(self):
        def texto(contas_disponiveis, contas_selecionadas):
            if not contas_disponiveis:
                return "Contas a migrar: nenhuma encontrada"
            if contas_selecionadas is None:
                return f"Contas a migrar: todas ({len(contas_disponiveis)})"
            return f"Contas a migrar: {len(contas_selecionadas)} de {len(contas_disponiveis)}"

        self.label_contas_tb_ol.config(
            text=texto(self.contas_thunderbird_disponiveis, self.contas_selecionadas_tb_ol)
        )
        self.label_contas_ol_tb.config(
            text=texto(self.contas_outlook_disponiveis, self.contas_selecionadas_ol_tb)
        )

    def _abrir_selecao_contas(self, direcao: str):
        if direcao == "tb_ol":
            titulo = "Selecionar contas do Thunderbird a migrar"
            contas_disponiveis = self.contas_thunderbird_disponiveis
            selecao_atual = self.contas_selecionadas_tb_ol
        else:
            titulo = "Selecionar contas do Outlook a migrar"
            contas_disponiveis = self.contas_outlook_disponiveis
            selecao_atual = self.contas_selecionadas_ol_tb

        janela = tb.Toplevel(self.root)
        janela.title(titulo)
        janela.geometry("420x420")
        janela.transient(self.root)
        janela.grab_set()

        tb.Label(
            janela, text=titulo, font=("Segoe UI", 11, "bold"), padding=(16, 16, 16, 4),
        ).pack(anchor="w")
        tb.Label(
            janela,
            text="Desmarque as contas que não devem entrar nesta migração.",
            bootstyle="secondary", padding=(16, 0, 16, 10),
        ).pack(anchor="w")

        area_lista = tb.Frame(janela, padding=(16, 0))
        area_lista.pack(fill="both", expand=True)

        marcado_atual = (
            {c.lower() for c in selecao_atual} if selecao_atual is not None
            else {c.lower() for c in contas_disponiveis}
        )
        variaveis = {}
        for conta in contas_disponiveis:
            var = tk.BooleanVar(value=conta.lower() in marcado_atual)
            variaveis[conta] = var
            tb.Checkbutton(area_lista, text=conta, variable=var, bootstyle="round-toggle").pack(
                anchor="w", pady=3
            )

        linha_atalhos = tb.Frame(janela, padding=(16, 6))
        linha_atalhos.pack(fill="x")
        tb.Button(
            linha_atalhos, text="Marcar todas", bootstyle="link",
            command=lambda: [v.set(True) for v in variaveis.values()],
        ).pack(side="left")
        tb.Button(
            linha_atalhos, text="Desmarcar todas", bootstyle="link",
            command=lambda: [v.set(False) for v in variaveis.values()],
        ).pack(side="left")

        def confirmar():
            marcadas = [conta for conta, var in variaveis.items() if var.get()]
            nova_selecao = None if len(marcadas) == len(contas_disponiveis) else set(marcadas)
            if direcao == "tb_ol":
                self.contas_selecionadas_tb_ol = nova_selecao
            else:
                self.contas_selecionadas_ol_tb = nova_selecao
            self._atualizar_labels_contas()
            janela.destroy()

        linha_botoes = tb.Frame(janela, padding=(16, 8, 16, 16))
        linha_botoes.pack(fill="x")
        tb.Button(linha_botoes, text="Cancelar", bootstyle="secondary-outline", command=janela.destroy).pack(
            side="right", padx=(8, 0)
        )
        tb.Button(linha_botoes, text="Confirmar", bootstyle="success", command=confirmar).pack(side="right")

    def _escolher_pasta(self, variavel: tk.StringVar):
        inicial = variavel.get() or str(pasta_do_programa())
        escolhida = filedialog.askdirectory(title="Escolha a pasta com os e-mails exportados", initialdir=inicial)
        if escolhida:
            variavel.set(escolhida)

    def _logar(self, texto):
        self.log.config(state="normal")
        self.log.insert("end", texto)
        self.log.see("end")
        self.log.config(state="disabled")

    def _consumir_fila(self):
        try:
            while True:
                texto = self.fila.get_nowait()
                self._logar(texto)
        except queue.Empty:
            pass
        if self.em_execucao:
            self._garantir_janela_visivel()
        self.root.after(150, self._consumir_fila)

    def _garantir_janela_visivel(self):
        """Durante uma operação em andamento, traz a janela de volta se ela
        for minimizada/sumir de vista - pra nunca perder o rastro de uma
        migração longa rodando em segundo plano."""
        try:
            if self.root.state() == "iconic":
                self.root.deiconify()
                self.root.lift()
        except Exception:
            pass

    def _iniciar_fase(self, texto_fase: str):
        """Zera o cronômetro/barra pra uma nova fase (ex: ao passar de
        'exportando' pra 'importando' dentro de 'Fazer tudo automaticamente').
        Thread-safe: os widgets só são tocados via root.after."""
        self._fase_inicio = time.monotonic()

        def atualizar():
            self.label_fase.config(text=texto_fase)
            self.label_progresso.config(text="Calculando...")
            self.progresso.config(value=0)

        self.root.after(0, atualizar)

    def _reportar_progresso(self, feito: int, total: int):
        self.root.after(0, lambda: self._atualizar_progresso_ui(feito, total))

    def _atualizar_progresso_ui(self, feito: int, total: int):
        total = max(total, 1)
        percentual = min(100, int(feito * 100 / total))
        self.progresso.config(value=percentual)

        decorrido = time.monotonic() - self._fase_inicio
        texto_eta = ""
        if feito > 0 and decorrido > 3:
            taxa = feito / decorrido
            if taxa > 0:
                restante = (total - feito) / taxa
                texto_eta = f" — tempo restante estimado: ~{_formatar_duracao(restante)}"

        self.label_progresso.config(
            text=f"{_formatar_numero(feito)} / {_formatar_numero(total)} ({percentual}%){texto_eta}"
        )

    def _rodar_em_thread(self, funcao):
        if self.em_execucao:
            return
        self.em_execucao = True
        for botao in (
            self.btn_exportar_tb, self.btn_importar_ol, self.btn_tudo_tb_ol,
            self.btn_exportar_ol, self.btn_importar_tb, self.btn_tudo_ol_tb,
        ):
            botao.config(state="disabled")

        sys.stdout = EscritorFila(self.fila, self.arquivo_log)
        sys.stderr = EscritorFila(self.fila, self.arquivo_log)

        def alvo():
            pythoncom.CoInitialize()
            deu_erro = False
            outlook_caiu = False
            try:
                funcao()
            except Exception as erro:
                deu_erro = True
                outlook_caiu = _conexao_outlook_perdida(erro)
                if outlook_caiu:
                    print("\n[ERRO] O Outlook ficou indisponível no meio da operação (RPC caiu -")
                    print("       provavelmente o Outlook reiniciou ou travou). Parando aqui.\n")
                else:
                    print("\n[ERRO] Ocorreu um problema inesperado, a operação parou no meio:\n")
                print(traceback.format_exc())
            finally:
                pythoncom.CoUninitialize()
                self.root.after(0, lambda: self._finalizar_execucao(deu_erro, outlook_caiu))

        threading.Thread(target=alvo, daemon=True).start()

    def _finalizar_execucao(self, deu_erro: bool = False, outlook_caiu: bool = False):
        self.em_execucao = False
        self._restaurar_botoes()
        if outlook_caiu:
            self.label_fase.config(text="Parou: o Outlook ficou indisponível")
            Messagebox.show_error(
                "O Outlook ficou indisponível no meio da operação (provavelmente reiniciou ou "
                "travou) - isso pode acontecer depois de muitas operações seguidas. O que já tinha "
                "sido feito foi mantido. Verifique se o Outlook está aberto e normal, e clique em "
                "importar/exportar de novo - ele continua de onde parou, sem duplicar.",
                "O Outlook ficou indisponível",
            )
        elif deu_erro:
            self.label_fase.config(text="Parou por causa de um erro")
            Messagebox.show_error(
                "A operação parou no meio por causa de um erro - veja o final do andamento "
                "pra entender o que aconteceu. O que já tinha sido feito até ali foi mantido; "
                "clicar em importar/exportar de novo continua de onde parou, sem duplicar.",
                "Parou por causa de um erro",
            )
        else:
            self.label_fase.config(text="Concluído")
            Messagebox.show_info("A operação terminou. Veja os detalhes no andamento.", "Concluído")

    def _ao_clicar_exportar_thunderbird(self):
        PASTA_EXPORTACAO_PADRAO.mkdir(parents=True, exist_ok=True)
        self.pasta_importar_outlook.set(str(PASTA_EXPORTACAO_PADRAO))

        def tarefa():
            self._iniciar_fase("Exportando e-mails do Thunderbird...")
            print(f"=== EXPORTANDO DO THUNDERBIRD PARA: {PASTA_EXPORTACAO_PADRAO} ===\n")
            exportar_desta_maquina(
                PASTA_EXPORTACAO_PADRAO, progresso=self._reportar_progresso,
                contas_selecionadas=self.contas_selecionadas_tb_ol,
            )

        self._rodar_em_thread(tarefa)

    def _ao_clicar_importar_outlook(self):
        pasta = self.pasta_importar_outlook.get().strip()
        if not pasta or not Path(pasta).exists():
            Messagebox.show_warning("Escolha uma pasta válida com os e-mails exportados (.eml).", "Pasta inválida")
            return

        def tarefa():
            self._iniciar_fase("Importando e-mails para o Outlook...")
            print(f"=== IMPORTANDO PARA O OUTLOOK, DE: {pasta} ===\n")
            com_retry_automatico(
                importar_para_outlook, Path(pasta), progresso=self._reportar_progresso,
                marcar_como_lido=self.marcar_lido_tb_ol.get(),
            )

        self._rodar_em_thread(tarefa)

    def _ao_clicar_tudo_thunderbird_outlook(self):
        PASTA_EXPORTACAO_PADRAO.mkdir(parents=True, exist_ok=True)
        self.pasta_importar_outlook.set(str(PASTA_EXPORTACAO_PADRAO))

        def tarefa():
            self._iniciar_fase("Fase 1 de 2 — Exportando do Thunderbird...")
            print(f"=== EXPORTANDO DO THUNDERBIRD PARA: {PASTA_EXPORTACAO_PADRAO} ===\n")
            qtd = exportar_desta_maquina(
                PASTA_EXPORTACAO_PADRAO, progresso=self._reportar_progresso,
                contas_selecionadas=self.contas_selecionadas_tb_ol,
            )
            if qtd:
                self._iniciar_fase("Fase 2 de 2 — Importando para o Outlook...")
                print("\n=== IMPORTANDO PARA O OUTLOOK ===\n")
                com_retry_automatico(
                    importar_para_outlook, PASTA_EXPORTACAO_PADRAO, progresso=self._reportar_progresso,
                    marcar_como_lido=self.marcar_lido_tb_ol.get(),
                )

        self._rodar_em_thread(tarefa)

    def _ao_clicar_exportar_outlook(self):
        PASTA_EXPORTACAO_OUTLOOK_PADRAO.mkdir(parents=True, exist_ok=True)
        self.pasta_importar_thunderbird.set(str(PASTA_EXPORTACAO_OUTLOOK_PADRAO))

        def tarefa():
            self._iniciar_fase("Exportando e-mails do Outlook...")
            print(f"=== EXPORTANDO DO OUTLOOK PARA: {PASTA_EXPORTACAO_OUTLOOK_PADRAO} ===\n")
            com_retry_automatico(
                exportar_outlook_desta_maquina, PASTA_EXPORTACAO_OUTLOOK_PADRAO,
                progresso=self._reportar_progresso, contas_selecionadas=self.contas_selecionadas_ol_tb,
            )

        self._rodar_em_thread(tarefa)

    def _ao_clicar_importar_thunderbird(self):
        pasta = self.pasta_importar_thunderbird.get().strip()
        if not pasta or not Path(pasta).exists():
            Messagebox.show_warning("Escolha uma pasta válida com os e-mails exportados do Outlook.", "Pasta inválida")
            return

        def tarefa():
            self._iniciar_fase("Importando e-mails para o Thunderbird...")
            print(f"=== IMPORTANDO PARA O THUNDERBIRD, DE: {pasta} ===\n")
            print("(passo separado: copiando os arquivos, sem apagar a pasta de origem)\n")
            importar_desta_maquina_para_thunderbird(Path(pasta), progresso=self._reportar_progresso, mover=False)

        self._rodar_em_thread(tarefa)

    def _ao_clicar_tudo_outlook_thunderbird(self):
        PASTA_EXPORTACAO_OUTLOOK_PADRAO.mkdir(parents=True, exist_ok=True)
        self.pasta_importar_thunderbird.set(str(PASTA_EXPORTACAO_OUTLOOK_PADRAO))

        def tarefa():
            self._iniciar_fase("Fase 1 de 2 — Exportando do Outlook...")
            print(f"=== EXPORTANDO DO OUTLOOK PARA: {PASTA_EXPORTACAO_OUTLOOK_PADRAO} ===\n")
            qtd = com_retry_automatico(
                exportar_outlook_desta_maquina, PASTA_EXPORTACAO_OUTLOOK_PADRAO,
                progresso=self._reportar_progresso, contas_selecionadas=self.contas_selecionadas_ol_tb,
            )
            if qtd:
                self._iniciar_fase("Fase 2 de 2 — Importando para o Thunderbird...")
                print("\n=== IMPORTANDO PARA O THUNDERBIRD ===\n")
                print("(fluxo automático: movendo os arquivos, sem duplicar espaço em disco)\n")
                importar_desta_maquina_para_thunderbird(
                    PASTA_EXPORTACAO_OUTLOOK_PADRAO, progresso=self._reportar_progresso, mover=True
                )

        self._rodar_em_thread(tarefa)


def main():
    root = tb.Window(themename="flatly")
    App(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        erro = traceback.format_exc()
        try:
            PASTA_LOGS.mkdir(parents=True, exist_ok=True)
            with open(PASTA_LOGS / "erro_fatal.txt", "a", encoding="utf-8") as f:
                f.write(f"\n--- {datetime.now()} ---\n{erro}\n")
        except OSError:
            pass
        try:
            Messagebox.show_error(erro, "Erro inesperado")
        except Exception:
            pass
