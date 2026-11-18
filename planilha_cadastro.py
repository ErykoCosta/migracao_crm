"""Manipulação da planilha PLANILHA DE CADASTRO.xlsx e das colunas de controle."""
import shutil
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

PLANILHA_PADRAO = Path("PLANILHA DE CADASTRO.xlsx")
ABAS_PADRAO = ("Clientes 2025", "Clientes 2026")
COLUNAS_CONTROLE = (
    ("Etapa", 4, 15),
    ("Documentos origem", 5, 22),
    ("Documentos destino", 6, 22),
    ("Obs", 7, 38),
)


def garantir_backup(caminho_planilha: Path | str = PLANILHA_PADRAO) -> Path:
    """Cria uma cópia de backup caso ainda não exista."""
    caminho = Path(caminho_planilha)
    backup = caminho.with_name(f"{caminho.stem}.backup{caminho.suffix}")
    if caminho.exists() and not backup.exists():
        shutil.copy2(caminho, backup)
    return backup


def _copiar_estilo(origem, destino, incluir_fill: bool = True) -> None:
    if origem.has_style:
        destino.font = copy(origem.font)
        if incluir_fill:
            destino.fill = copy(origem.fill)
        destino.border = copy(origem.border)
        destino.alignment = copy(origem.alignment)
        destino.number_format = origem.number_format


def _encontrar_coluna(ws, nome_coluna: str) -> int | None:
    alvo = nome_coluna.strip().casefold()
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val is not None and str(val).strip().casefold() == alvo:
            return col
    return None


def garantir_colunas_controle(
    caminho_planilha: Path | str = PLANILHA_PADRAO,
    abas: tuple[str, ...] = ABAS_PADRAO,
) -> dict[str, dict[str, int]]:
    """Garante as colunas Etapa, Documentos origem, Documentos destino e Obs nas abas."""
    caminho = Path(caminho_planilha)
    garantir_backup(caminho)
    wb = openpyxl.load_workbook(caminho)
    mapa: dict[str, dict[str, int]] = {}
    alterado = False

    try:
        for aba in abas:
            if aba not in wb.sheetnames:
                continue
            ws = wb[aba]
            cols_aba: dict[str, int] = {}
            max_col = 3
            for nome_col, pos_padrao, largura in COLUNAS_CONTROLE:
                col = _encontrar_coluna(ws, nome_col)
                if col is None:
                    col = pos_padrao if ws.cell(row=1, column=pos_padrao).value is None else ws.max_column + 1
                    header = ws.cell(row=1, column=col, value=nome_col)
                    _copiar_estilo(ws.cell(row=1, column=3), header, incluir_fill=True)
                    ws.column_dimensions[get_column_letter(col)].width = largura
                    alterado = True
                cols_aba[nome_col] = col
                if col > max_col:
                    max_col = col

            letra_max = get_column_letter(max_col)
            filtro_esperado = f"A1:{letra_max}{ws.max_row}"
            if ws.auto_filter and ws.auto_filter.ref and ws.auto_filter.ref != filtro_esperado:
                ws.auto_filter.ref = filtro_esperado
                alterado = True

            mapa[aba] = cols_aba

        if alterado:
            wb.save(caminho)
    finally:
        wb.close()
    return mapa


def garantir_coluna_etapa(
    caminho_planilha: Path | str = PLANILHA_PADRAO,
    abas: tuple[str, ...] = ABAS_PADRAO,
) -> dict[str, int]:
    """Compatibilidade: garante todas as colunas de controle e retorna os índices de Etapa."""
    mapa = garantir_colunas_controle(caminho_planilha, abas=abas)
    return {aba: cols["Etapa"] for aba, cols in mapa.items()}


def atualizar_etapa_cliente(
    caminho_planilha: Path | str,
    aba: str,
    linha_excel: int,
    status_etapa: str,
    docs_origem: str | None = None,
    docs_destino: str | None = None,
    obs: str | None = None,
) -> None:
    """Grava Etapa, Documentos origem, Documentos destino e Obs na linha do cliente."""
    if status_etapa not in ("concluido", "pendente"):
        raise ValueError(f"Status de etapa inválido: {status_etapa!r}")

    caminho = Path(caminho_planilha)
    garantir_backup(caminho)
    wb = openpyxl.load_workbook(caminho)
    try:
        ws = wb[aba]
        cols = {nome: _encontrar_coluna(ws, nome) for nome, _, _ in COLUNAS_CONTROLE}
        if any(v is None for v in cols.values()):
            wb.close()
            mapa = garantir_colunas_controle(caminho, abas=(aba,))
            cols = mapa[aba]
            wb = openpyxl.load_workbook(caminho)
            ws = wb[aba]

        cel_ref = ws.cell(row=linha_excel, column=3)
        valores = [("Etapa", status_etapa)]
        if docs_origem is not None:
            valores.append(("Documentos origem", docs_origem))
        if docs_destino is not None:
            valores.append(("Documentos destino", docs_destino))
        if obs is not None:
            valores.append(("Obs", obs))

        for nome_col, valor in valores:
            cel = ws.cell(row=linha_excel, column=cols[nome_col], value=valor)
            _copiar_estilo(cel_ref, cel, incluir_fill=False)

        wb.save(caminho)
    finally:
        wb.close()


if __name__ == "__main__":
    mapa = garantir_colunas_controle()
    print("Colunas de controle garantidas:", mapa)
