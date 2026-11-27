import contextlib
import io
import tempfile
import unittest
from email.message import Message
from pathlib import Path

import exportar_eml

MBOX = (
    "From - Mon Nov 23 10:00:00 2026\n"
    "X-Mozilla-Status: 0001\n"
    "From: Empresa Exemplo Ltda <contato@exemplo.test>\n"
    "Subject: Mensagem lida\n"
    "\n"
    "corpo 1\n"
    "\n"
    "From - Mon Nov 23 10:05:00 2026\n"
    "X-Mozilla-Status: 0009\n"
    "From: Empresa Exemplo Ltda <contato@exemplo.test>\n"
    "Subject: Mensagem apagada sem compactar\n"
    "\n"
    "corpo 2\n"
    "\n"
    "From - Mon Nov 23 10:10:00 2026\n"
    "From: Empresa Exemplo Ltda <contato@exemplo.test>\n"
    "Subject: Sem cabecalho de status\n"
    "\n"
    "corpo 3\n"
)


def _msg(status):
    m = Message()
    if status is not None:
        m["X-Mozilla-Status"] = status
    return m


class TestMensagemExcluida(unittest.TestCase):
    def test_flag_expunged(self):
        self.assertTrue(exportar_eml.mensagem_excluida(_msg("0008")))
        self.assertTrue(exportar_eml.mensagem_excluida(_msg("0009")))
        self.assertFalse(exportar_eml.mensagem_excluida(_msg("0001")))
        self.assertFalse(exportar_eml.mensagem_excluida(_msg("0000")))

    def test_status_ausente_ou_invalido_nao_exclui(self):
        self.assertFalse(exportar_eml.mensagem_excluida(_msg(None)))
        self.assertFalse(exportar_eml.mensagem_excluida(_msg("xyz")))


class TestConverterMbox(unittest.TestCase):
    def test_ignora_mensagens_excluidas(self):
        with tempfile.TemporaryDirectory() as tmp:
            caixa = Path(tmp) / "Inbox"
            caixa.write_text(MBOX, encoding="utf-8")
            destino = Path(tmp) / "saida"
            with contextlib.redirect_stdout(io.StringIO()) as saida:
                total = exportar_eml.converter_mbox_para_eml(caixa, destino)
            nomes = sorted(p.name for p in destino.iterdir())
            self.assertEqual(total, 2)
            self.assertEqual(len(nomes), 2)
            self.assertFalse(any("apagada" in n for n in nomes))
            self.assertIn("1 mensagem(ns) excluída(s)", saida.getvalue())


if __name__ == "__main__":
    unittest.main()
