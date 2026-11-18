"""Manipulação da planilha PLANILHA DE CADASTRO.xlsx e da coluna Etapa."""
import shutil
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

PLANILHA_PADRAO = Path("PLANILHA DE CADASTRO.xlsx")
ABAS_PADRAO = ("Clientes 2025", "Clientes 2026")


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


def _encontrar_coluna_etapa(ws) -> int | None:
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val is not None and str(val).strip().casefold() == "etapa":
            return col
    return None


def garantir_coluna_etapa(
    caminho_planilha: Path | str = PLANILHA_PADRAO,
    abas: tuple[str, ...] = ABAS_PADRAO,
) -> dict[str, int]:
    """Garante a existência da coluna 'Etapa' nas abas informadas e retorna seus índices."""
    caminho = Path(caminho_planilha)
    garantir_backup(caminho)
    wb = openpyxl.load_workbook(caminho)
    colunas: dict[str, int] = {}
    alterado = False

    try:
        for aba in abas:
            if aba not in wb.sheetnames:
                continue
            ws = wb[aba]
            col = _encontrar_coluna_etapa(ws)
            if col is None:
                col = 4 if ws.cell(row=1, column=4).value is None else ws.max_column + 1
                header = ws.cell(row=1, column=col, value="Etapa")
                ref_header = ws.cell(row=1, column=3)
                _copiar_estilo(ref_header, header, incluir_fill=True)
                letra = get_column_letter(col)
                ws.column_dimensions[letra].width = 15
                if ws.auto_filter and ws.auto_filter.ref:
                    ws.auto_filter.ref = f"A1:{letra}{ws.max_row}"
                alterado = True
            colunas[aba] = col

        if alterado:
            wb.save(caminho)
    finally:
        wb.close()
    return colunas


def atualizar_etapa_cliente(
    caminho_planilha: Path | str,
    aba: str,
    linha_excel: int,
    status_etapa: str,
) -> None:
    """Grava 'concluido' ou 'pendente' na coluna Etapa da linha do cliente e salva a planilha."""
    if status_etapa not in ("concluido", "pendente"):
        raise ValueError(f"Status de etapa inválido: {status_etapa!r}")

    caminho = Path(caminho_planilha)
    garantir_backup(caminho)
    wb = openpyxl.load_workbook(caminho)
    try:
        ws = wb[aba]
        col = _encontrar_coluna_etapa(ws)
        if col is None:
            wb.close()
            colunas = garantir_coluna_etapa(caminho, abas=(aba,))
            col = colunas[aba]
            wb = openpyxl.load_workbook(caminho)
            ws = wb[aba]

        cel_destino = ws.cell(row=linha_excel, column=col, value=status_etapa)
        cel_ref = ws.cell(row=linha_excel, column=3)
        _copiar_estilo(cel_ref, cel_destino, incluir_fill=False)
        wb.save(caminho)
    finally:
        wb.close()


if __name__ == "__main__":
    cols = garantir_coluna_etapa()
    print("Coluna Etapa garantida:", cols)
