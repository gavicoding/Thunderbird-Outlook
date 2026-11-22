import unittest

import vcard_util


class TestVCardUtil(unittest.TestCase):
    def test_contato_vazio_has_all_campos(self):
        c = vcard_util.contato_vazio()
        self.assertEqual(set(c.keys()), set(vcard_util.CAMPOS))
        self.assertTrue(all(v == "" for v in c.values()))

    def test_roundtrip_complete_contact(self):
        original = {
            "nome": "Gabriel",
            "sobrenome": "Souza",
            "nome_exibicao": "Gabriel Souza",
            "apelido": "Gabi",
            "email1": "gabriel@example.com",
            "email2": "backup@example.org",
            "tel_comercial": "+55 11 4000-1000",
            "tel_residencial": "+55 11 3000-2000",
            "tel_celular": "+55 11 99999-8888",
            "tel_fax": "+55 11 4000-1099",
            "empresa": "Tech Ops",
            "departamento": "Infrastructure",
            "cargo": "Automation Engineer",
            "endereco_com": "Av. Paulista, 1000",
            "cidade_com": "São Paulo",
            "estado_com": "SP",
            "cep_com": "01310-100",
            "pais_com": "Brasil",
            "endereco_res": "Rua das Flores, 42",
            "cidade_res": "Campinas",
            "estado_res": "SP",
            "cep_res": "13000-000",
            "pais_res": "Brasil",
            "site_com": "https://work.example.com",
            "site_pessoal": "https://blog.example.com",
            "aniversario": "1995-07-21",
            "notas": "Linha 1, com vírgula\nLinha 2 \\ barra",
        }

        vcf = vcard_util.contato_para_vcard(original)
        self.assertIn("BEGIN:VCARD\r\nVERSION:3.0\r\n", vcf)
        self.assertIn("BDAY:19950721", vcf)
        self.assertTrue(vcf.endswith("END:VCARD\r\n"))

        parsed = vcard_util.vcard_para_contato(vcf)
        self.assertEqual(parsed, original)

    def test_display_name_fallbacks(self):
        vcf_name = vcard_util.contato_para_vcard({"nome": "Ana", "sobrenome": "Lima"})
        self.assertIn("FN:Ana Lima\r\n", vcf_name)

        vcf_email = vcard_util.contato_para_vcard({"email1": "ana@example.com"})
        self.assertIn("FN:ana@example.com\r\n", vcf_email)

        vcf_empty = vcard_util.contato_para_vcard({})
        self.assertIn("FN:Sem nome\r\n", vcf_empty)

    def test_rfc2426_line_unfolding(self):
        folded_vcard = (
            "BEGIN:VCARD\r\n"
            "VERSION:3.0\r\n"
            "N:Oliveira;Joao;;;\r\n"
            "FN:Joao Oliveira\r\n"
            "NOTE:Primeira parte da nota \r\n"
            " continuacao dobrada\r\n"
            "\tcom tab\r\n"
            "END:VCARD\r\n"
        )
        parsed = vcard_util.vcard_para_contato(folded_vcard)
        self.assertEqual(parsed["nome"], "Joao")
        self.assertEqual(parsed["sobrenome"], "Oliveira")
        self.assertEqual(parsed["notas"], "Primeira parte da nota continuacao dobradacom tab")

    def test_separar_vcards_splits_multiple_contacts(self):
        v1 = vcard_util.contato_para_vcard({"nome": "Alice", "email1": "alice@example.com"})
        v2 = vcard_util.contato_para_vcard({"nome": "Bob", "email1": "bob@example.com"})
        combined = f"{v1}\r\n{v2}"

        blocks = vcard_util.separar_vcards(combined)
        self.assertEqual(len(blocks), 2)
        self.assertEqual(vcard_util.vcard_para_contato(blocks[0])["nome"], "Alice")
        self.assertEqual(vcard_util.vcard_para_contato(blocks[1])["nome"], "Bob")


if __name__ == "__main__":
    unittest.main()
