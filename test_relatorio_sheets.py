import asyncio
import io
import unittest

from openpyxl import load_workbook

from relatorio_sheets import EPOCA_SHEETS, gerar_excel, montar_blocos
from datetime import datetime
from servidor import processar_excel


class RelatorioSheetsTests(unittest.TestCase):
    def setUp(self):
        self.dia = (datetime(2026, 9, 7) - EPOCA_SHEETS).days
        self.abas = [
            {"titulo": "BASE_ANDRE_LUCAS", "linhas": [
                [self.dia, "Nome Cliente Dois", 0.5, 47.38, 2, None, "Suporte Online"],
                [self.dia, "Nome Cliente Um", 0.25, 10, 0, None, "Suporte Online"],
                [self.dia + 1, "Outro dia", 0.5, 99, 0, None, "Suporte Online"],
            ]},
            {"titulo": "BASE_VITÓRIA_RÉGIA", "linhas": [
                [self.dia, "Nome Cliente Agência", 0.5, 20, 1, None, "Vale Viagens"],
            ]},
        ]

    def test_geral_filtra_data_ordena_hora_e_soma_taxa(self):
        blocos = montar_blocos(self.abas, "2026-09-07")
        self.assertEqual(len(blocos), 2)
        self.assertEqual(blocos[0]["linhas"][0], ["Cliente Um", "06:00", 10, 0, 10])
        self.assertEqual(blocos[0]["linhas"][1][-1], 49.38)

    def test_agencia_normaliza_acentos(self):
        blocos = montar_blocos(self.abas, "2026-09-07", "agencia")
        self.assertEqual([b["agente"] for b in blocos], ["VITÓRIA RÉGIA"])

    def test_excel_compativel_com_checker_e_setor_mesclado(self):
        arquivo = gerar_excel(montar_blocos(self.abas, "2026-09-07"), "2026-09-07")
        tabela = asyncio.run(processar_excel(arquivo))["tabela"]
        self.assertEqual(len(tabela), 3)
        self.assertEqual(tabela[0]["agente"], "ANDRE LUCAS - SUPORTE ONLINE")
        self.assertEqual(tabela[1]["valor"], 49.38)

    def test_texto_nao_vira_formula(self):
        blocos = montar_blocos(self.abas, "2026-09-07")
        blocos[0]["linhas"][0][0] = '=HYPERLINK("https://example.com")'
        sheet = load_workbook(io.BytesIO(gerar_excel(blocos, "2026-09-07"))).active
        self.assertEqual(sheet["A5"].data_type, "s")

    def test_rejeita_data_e_tipo_invalidos(self):
        for data, tipo in [("2026-02-30", "geral"), ("2026-09-07", "outro")]:
            with self.assertRaises(ValueError):
                montar_blocos(self.abas, data, tipo)


if __name__ == "__main__":
    unittest.main()
