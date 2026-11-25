[🇺🇸 English](README.md) | **🇧🇷 Português**

# Migrador Thunderbird ⇄ Outlook

[![CI](https://github.com/gavicoding/Thunderbird-Outlook/actions/workflows/ci.yml/badge.svg)](https://github.com/gavicoding/Thunderbird-Outlook/actions/workflows/ci.yml)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![Plataforma: Windows](https://img.shields.io/badge/plataforma-Windows-0078D6.svg)](https://www.microsoft.com/windows)
[![Licença: Apache-2.0](https://img.shields.io/badge/licen%C3%A7a-Apache_2.0-green.svg)](LICENSE)

Leva **e-mails e contatos** do **Mozilla Thunderbird** para o **Outlook clássico** e vice-versa, no Windows. Mantém as pastas, as **datas originais**, os remetentes, os anexos e as imagens dentro das mensagens. Tem uma janela simples para quem não quer usar linha de comando.

---

## Para que serve

Trocar de programa de e-mail costuma dar errado de dois jeitos: as pastas se perdem ou todas as mensagens importadas aparecem com a data de hoje, como se tivessem acabado de chegar (ou como rascunhos não enviados). Este migrador foi feito justamente para evitar isso.

## O que você ganha

- **Histórico intacto:** cada mensagem chega com a data em que foi recebida e com o remetente certo.
- **Nada sobrescrito:** e-mails e contatos migrados vão para pastas próprias, sem mexer na caixa que já existe.
- **Funciona entre computadores:** exporte numa máquina, copie a pasta gerada (pendrive ou rede) e importe na outra.
- **Contatos também:** a agenda vai junto, no formato vCard 3.0.

---

## Jeito mais fácil: `MigrarThunderbirdOutlook.exe`

Programa com janela (`migrar_app.py`, empacotado como `.exe` via PyInstaller).
Detecta sozinho o Thunderbird e o Outlook clássico desta máquina e tem duas
abas, uma pra cada sentido:

- **Thunderbird → Outlook**: exporta o Thunderbird pra uma pasta `.eml` e
  importa pro Outlook automaticamente.
- **Outlook → Thunderbird**: exporta o Outlook pra uma pasta já no formato
  mbox nativo do Thunderbird e copia pra dentro de `Mail/Local Folders` de um
  perfil.

Em cada aba, o botão **"Fazer tudo automaticamente"** roda as duas etapas
seguidas quando ambos os programas estão na mesma máquina. As etapas também
funcionam separadas - exporte numa máquina, copie a pasta gerada (fica do
lado do `.exe`) pra um pendrive/rede, e importe em outra máquina que tenha o
programa de destino.

Cada aba também tem a opção **"Também migrar o catálogo de endereços
(contatos)"** - desmarcada por padrão, marque se quiser levar os contatos
junto com os e-mails.

Não precisa instalar Python nem nada - só rodar o `.exe`.

## Rodando os scripts em vez do .exe

```bash
pip install -r requirements.txt
python migrar_app.py
```

### Módulos individuais (sem interface gráfica)

- **`exportar_eml.py`** - Thunderbird → `.eml`. Varre o perfil inteiro
  (Inbox, Enviados, pastas locais, contas IMAP, subpastas), nomeia cada
  conta pelo e-mail real (lido do `prefs.js`, evita ambiguidade tipo
  `mail.dominio.com` vs `mail.dominio.com-1`), e usa o prefixo `\\?\` do
  Windows pra não esbarrar no limite de 260 caracteres em pastas aninhadas.
  Não precisa de nenhuma dependência externa.

- **`importar_outlook.py`** - `.eml` → Outlook. Importa automaticamente via
  automação COM, identificando pelo e-mail em qual conta já configurada no
  Outlook cada backup deve cair (cai na conta padrão como fallback se aquela
  conta não estiver configurada ali). Requer Outlook **clássico** (o "novo
  Outlook" não tem essa automação) e `pywin32`.

  Cada `.eml` é reconstruído como um `MailItem` novo (não usa
  `NameSpace.OpenSharedItem`, que nos testes abre `.msg` mas recusa `.eml`
  com erro de "caminho inválido" em builds recentes do Outlook). Isso exige
  um cuidado extra: por padrão, todo `MailItem` criado via automação grava a
  data de **hoje** como "recebido em", mesmo setando a propriedade de data
  manualmente - a menos que a flag `MSGFLAG_UNSENT` seja explicitamente
  desligada via `PropertyAccessor` antes de salvar, dizendo ao Outlook "isso
  não é rascunho novo, é e-mail migrado". Com isso (mais `PR_SENDER_*` pro
  remetente original), data e remetente originais ficam intactos - validado
  com e-mails reais.

- **`exportar_outlook.py`** - Outlook → mbox nativo do Thunderbird. Cada
  e-mail é salvo pelo Outlook como `.msg` e convertido pra `.eml` de verdade
  com a biblioteca `extract-msg` (reaproveita o cabeçalho original quando
  existe, ou sintetiza um a partir das propriedades da mensagem). Requer
  Outlook clássico, `pywin32` e `extract-msg`.

- **`importar_thunderbird.py`** - mbox → Thunderbird. Como a árvore já sai no
  formato nativo (arquivo sem extensão = pasta, `Nome.sbd/` = subpastas), a
  importação é só copiar pra dentro de `Mail/Local Folders` do perfil - sem
  precisar de complemento nenhum do Thunderbird. Tudo entra dentro de uma
  pasta própria ("Importado do Outlook"), nunca mexendo nas pastas reais que
  já existem.

### Catálogo de endereços (contatos) - opcional

Além de e-mail, o programa também sabe migrar contatos, usando vCard
(`.vcf`) como formato comum entre os dois lados:

- **`exportar_contatos_thunderbird.py`** - lê direto o(s) banco(s) SQLite do
  catálogo de endereços do Thunderbird (`abook.sqlite` e qualquer
  `abook-N.sqlite` extra), sem nenhuma dependência externa (só `sqlite3` da
  biblioteca padrão). Abre em modo só-leitura e confere se o schema esperado
  existe antes de consultar, pra não quebrar em versões muito diferentes do
  Thunderbird.

- **`importar_contatos_outlook.py`** - `.vcf` → Outlook, via automação COM,
  criando os contatos numa pasta própria ("Contatos Migrados do
  Thunderbird") dentro do Catálogo de Endereços padrão, sem mexer nos
  contatos que já existem.

- **`exportar_contatos_outlook.py`** - Outlook → `.vcf`, um arquivo por
  conta. Como escrever direto no banco de dados do Thunderbird é arriscado
  (schema pode variar, banco pode estar em uso), esse lado **não** importa
  automaticamente - a própria função já imprime o caminho de cada `.vcf` e a
  instrução pra importar pelo Catálogo de Endereços do Thunderbird
  (Ferramentas → Importar → Contatos), que é o jeito seguro e já testado
  pela Mozilla.

- **`vcard_util.py`** - conversão contato ↔ vCard 3.0 compartilhada pelos
  três módulos acima, pra escrever e ler sempre no mesmo "dialeto".

Compatibilidade do lado do Outlook: a automação funciona com qualquer
Outlook clássico (2010, 2013, 2016, 2019, 2021, atual) - é a mesma API COM
estável há anos. Só não funciona com o **novo Outlook** (o app reformulado
tipo web/UWP), que não expõe automação nenhuma.

## Se preferir fazer à mão

Depois de exportar (qualquer uma das direções), dá pra abrir a pasta de
saída e arrastar as subpastas direto para dentro do programa de destino.

## ⚠️ Aviso

Esta ferramenta lê e escreve e-mails e contatos direto nos perfis do
Thunderbird e do Outlook. O uso é por sua conta e risco - faça backup do
seu perfil/pastas antes de usar, especialmente na primeira vez. O software
é distribuído "no estado em que se encontra" (AS IS), sem garantias de
nenhum tipo, e o autor não se responsabiliza por eventuais perdas de dados
decorrentes do uso. Veja a licença ([LICENSE](LICENSE)) para os termos
completos.
