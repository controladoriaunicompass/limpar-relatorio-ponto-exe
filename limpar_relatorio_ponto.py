# -*- coding: utf-8 -*-
"""
limpar_relatorio_ponto.py

Objetivo:
    Ler um CSV/TXT bruto exportado do relatório de frequência detalhado por funcionário
    e transformar em uma base limpa, uma linha por funcionário + dia.

O que o script faz:
    1. Corrige textos com codificação quebrada, quando aparecerem como:
       FuncionÃ¡rio, SaÃ­da, AusÃªncias, ObservaÃ§Ãµes etc.
    2. Identifica os blocos de cada funcionário.
    3. Identifica o cabeçalho de cada bloco, inclusive quando há variações de colunas.
    4. Consolida marcações extras que vierem em linha de continuação.
       Exemplo:
           05/05 com 4 marcações na linha principal
           e mais 2 marcações na linha de baixo
       vira:
           Marcação 1 ... Marcação 6 na mesma linha.
    5. Mantém campos principais:
       empresa, período, funcionário, matrícula, departamento, cargo,
       data, turno previsto, marcações, normais, extras, ausências,
       banco de horas e observações.

Como usar no Windows:
    1. Salve este arquivo como: limpar_relatorio_ponto.py
    2. Coloque o CSV/TXT bruto na mesma pasta.
    3. Abra o CMD/Prompt na pasta.
    4. Rode:

       python limpar_relatorio_ponto.py "relatorio_bruto.csv" "relatorio_limpo.csv"

    Se o arquivo estiver como TXT:

       python limpar_relatorio_ponto.py "Texto colado(23).txt" "relatorio_limpo.csv"

Observação:
    O arquivo de saída é gerado em UTF-8 com BOM, ideal para abrir direto no Excel.
"""

import argparse
import csv
import re
from pathlib import Path


TIME_RE = re.compile(r"^(?:\d{1,3}:\d{2}|--:--|\(\d{1,3}:\d{2}\))$")
DATE_ROW_RE = re.compile(r"^\s*(\d{2}/\d{2})\s+([A-Za-zÀ-ÿ]{3})\.?\s*$", re.IGNORECASE)


def fix_text(value: str) -> str:
    """Corrige mojibake comum de UTF-8 lido/exportado errado."""
    if value is None:
        return ""

    text = str(value).strip()

    # Tenta corrigir casos como FuncionÃ¡rio -> Funcionário
    # Mantém o original se a conversão não fizer sentido.
    if "Ã" in text or "Â" in text:
        try:
            fixed = text.encode("latin1").decode("utf-8")
            text = fixed
        except Exception:
            pass

    return text.strip()


def is_time(value: str) -> bool:
    value = fix_text(value)
    return bool(TIME_RE.match(value))


def normalize_empty_time(value: str) -> str:
    value = fix_text(value)
    if value == "--:--":
        return ""
    return value


def read_csv_any_encoding(path: Path):
    """Lê o arquivo tentando as codificações mais comuns."""
    encodings = ["utf-8-sig", "utf-8", "latin1", "cp1252"]

    last_error = None
    for enc in encodings:
        try:
            with path.open("r", encoding=enc, newline="") as f:
                rows = list(csv.reader(f, delimiter=","))
            return [[fix_text(cell) for cell in row] for row in rows]
        except UnicodeDecodeError as e:
            last_error = e

    raise RuntimeError(f"Não foi possível ler o arquivo. Último erro: {last_error}")


def non_empty_cells(row):
    return [fix_text(c) for c in row if fix_text(c) != ""]


def find_cell_index(row, label):
    label_norm = fix_text(label).lower().replace(":", "")
    for i, cell in enumerate(row):
        c = fix_text(cell).lower().replace(":", "")
        if c == label_norm:
            return i
    return None


def value_after_label(row, label):
    idx = find_cell_index(row, label)
    if idx is None:
        return ""

    for cell in row[idx + 1:]:
        cell = fix_text(cell)
        if cell:
            return cell
    return ""


def values_between(row, start, end):
    if start is None:
        return []

    if end is None:
        end = len(row)

    vals = []
    for c in row[start:end]:
        c = fix_text(c)
        if c:
            vals.append(c)
    return vals


def times_between(row, start, end, keep_dash=False):
    vals = []
    for c in values_between(row, start, end):
        if is_time(c):
            c = fix_text(c)
            if keep_dash or c != "--:--":
                vals.append(c)
    return vals


def first_real_time_between(row, start, end):
    for c in times_between(row, start, end, keep_dash=False):
        return c
    return ""


def extract_header_positions(header_top, header_bottom):
    """
    Monta um mapa de posições a partir das duas linhas de cabeçalho.

    Exemplo:
        linha topo:    Dia ... Turno de Trabalho ... Turno Total ... Jornada realizada ... Normais ... Extras ...
        linha baixo:       Entrada Saída Entrada Saída ... Entrada Saída Entrada Saída ... Diurnas Noturna ...
    """
    def idx_top(name):
        name = name.lower()
        for i, c in enumerate(header_top):
            c = fix_text(c).lower()
            if c == name:
                return i
        return None

    pos = {
        "turno_trabalho": idx_top("turno de trabalho"),
        "turno_total": idx_top("turno total"),
        "jornada_realizada": idx_top("jornada realizada"),
        "normais": idx_top("normais"),
        "extras": idx_top("extras"),
        "ausencias": idx_top("ausências"),
    }

    # Observações costuma estar apenas na linha inferior.
    obs_idx = None
    bh_idx = None
    intrajor_idx = None
    interjor_idx = None

    for i, c in enumerate(header_bottom):
        c_norm = fix_text(c).lower()
        if c_norm.startswith("observa"):
            obs_idx = i
        elif c_norm == "b.h." or c_norm == "b.h":
            bh_idx = i
        elif c_norm.startswith("intrajor"):
            intrajor_idx = i
        elif c_norm.startswith("interjor"):
            interjor_idx = i

    pos["observacoes"] = obs_idx
    pos["bh"] = bh_idx
    pos["intrajor"] = intrajor_idx
    pos["interjor"] = interjor_idx

    # Índices das marcações reais: Entrada/Saída entre Jornada realizada e Normais.
    jornada_start = pos["jornada_realizada"]
    normais_start = pos["normais"]

    marcacao_indices = []
    if jornada_start is not None:
        end = normais_start if normais_start is not None else len(header_bottom)
        for i in range(jornada_start, end):
            label = fix_text(header_bottom[i]).lower()
            if label in {"entrada", "saída", "saida"}:
                marcacao_indices.append(i)

    # Índices do turno previsto: Entrada/Saída entre Turno de Trabalho e Turno Total.
    turno_indices = []
    if pos["turno_trabalho"] is not None:
        end = pos["turno_total"] if pos["turno_total"] is not None else len(header_bottom)
        for i in range(pos["turno_trabalho"], end):
            label = fix_text(header_bottom[i]).lower()
            if label in {"entrada", "saída", "saida"}:
                turno_indices.append(i)

    pos["marcacao_indices"] = marcacao_indices
    pos["turno_indices"] = turno_indices

    return pos


def split_date_weekday(value):
    value = fix_text(value)
    m = DATE_ROW_RE.match(value)
    if not m:
        return "", ""
    return m.group(1), m.group(2)


def is_day_row(row):
    return bool(row and DATE_ROW_RE.match(fix_text(row[0])))


def is_continuation_row(row, header_pos):
    """
    Linha de continuação:
        - primeira coluna vazia
        - possui horários nas posições de Jornada realizada
        - não é cabeçalho, totalizador ou assinatura
    """
    if not row or fix_text(row[0]) != "":
        return False

    contents = non_empty_cells(row)
    if not contents:
        return False

    blockers = [
        "totais", "extras", "assinatura", "trabalhados",
        "saldo do banco", "detalhamento", "acerto",
        "normal", "limite", "página", "pagina"
    ]

    joined = " ".join(contents).lower()
    if any(b in joined for b in blockers):
        return False

    marc_indices = header_pos.get("marcacao_indices", []) if header_pos else []
    for idx in marc_indices:
        if idx < len(row):
            near = get_time_near(row, idx, window=2)
            if near and near != "--:--":
                return True

    return False


def append_markings(record, row, header_pos, max_marcacoes):
    """Adiciona marcações reais da linha atual nas colunas Marcação 1...N."""
    for idx in header_pos.get("marcacao_indices", []):
        if idx < len(row):
            value = get_time_near(row, idx, window=2)
            if value and is_time(value):
                # Evita repetir a mesma marcação imediatamente se vier duplicada.
                marcacoes = record.setdefault("_marcacoes", [])
                if len(marcacoes) < max_marcacoes:
                    marcacoes.append(value)


def extract_observacoes(row, header_pos):
    obs_idx = header_pos.get("observacoes")
    if obs_idx is None or obs_idx >= len(row):
        return ""

    vals = []
    for c in row[obs_idx:]:
        c = fix_text(c)
        if c and not is_time(c):
            vals.append(c)

    return " | ".join(vals)


def extract_marcador(row, header_pos):
    """
    Marcador fica entre a coluna Dia e o início do Turno de Trabalho.
    Exemplos: F, Ab, #, A.
    """
    start_turno = header_pos.get("turno_trabalho")
    if start_turno is None:
        start_turno = 6

    vals = []
    for c in row[1:start_turno]:
        c = fix_text(c)
        if c and not is_time(c):
            vals.append(c)

    return " | ".join(vals)


def get_index_value(row, idx):
    if idx is None or idx >= len(row):
        return ""
    return normalize_empty_time(row[idx])


def get_time_near(row, idx, window=1):
    """
    Busca horário na coluna do cabeçalho e nas colunas vizinhas.

    No relatório bruto, em alguns blocos o texto Entrada/Saída aparece em uma
    coluna, mas o horário correspondente vem 1 coluna ao lado.
    A prioridade é: coluna seguinte, própria coluna, coluna anterior.
    """
    if idx is None:
        return ""

    candidates = []
    for delta in [1, 0, -1, 2, -2]:
        if abs(delta) <= window or delta in [1, 0, -1]:
            candidates.append(idx + delta)

    seen = set()
    for pos in candidates:
        if pos in seen:
            continue
        seen.add(pos)
        if 0 <= pos < len(row):
            value = normalize_empty_time(row[pos])
            if value and is_time(value):
                return value
    return ""


def finalize_record(record, output_rows, max_marcacoes):
    if not record:
        return

    marcacoes = record.pop("_marcacoes", [])
    for i in range(max_marcacoes):
        record[f"Marcação {i + 1}"] = marcacoes[i] if i < len(marcacoes) else ""

    output_rows.append(record)


def limpar_relatorio(input_path, output_path, max_marcacoes=12):
    input_path = Path(input_path)
    output_path = Path(output_path)

    rows = read_csv_any_encoding(input_path)

    empresa = ""
    periodo_inicio = ""
    periodo_fim = ""
    data_emissao = ""

    funcionario = ""
    matricula = ""
    pis = ""
    data_admissao = ""
    departamento = ""
    cargo = ""
    cpf = ""

    header_pos = {}
    current_record = None
    output_rows = []

    for line_number, row in enumerate(rows, start=1):
        # Garante tamanho mínimo para evitar erro de índice.
        if len(row) < 5:
            row = row + [""] * (5 - len(row))

        first = fix_text(row[0])
        contents = non_empty_cells(row)

        if not contents:
            continue

        # Empresa geralmente aparece na primeira célula da página/bloco.
        if first and "LIDERKRAFT" in first.upper():
            empresa = first

        # Linha de período e data de emissão.
        if any("Relatório de frequência" in c for c in contents):
            for i, c in enumerate(row):
                c = fix_text(c).lower().replace(":", "")
                if c == "período":
                    # próximos dois campos com formato de data
                    dates = [fix_text(x) for x in row[i + 1:] if re.match(r"^\d{2}/\d{2}/\d{4}$", fix_text(x))]
                    if dates:
                        periodo_inicio = dates[0]
                    if len(dates) > 1:
                        periodo_fim = dates[1]
                if c == "data de emissão":
                    dates = [fix_text(x) for x in row[i + 1:] if re.match(r"^\d{2}/\d{2}/\d{4}$", fix_text(x))]
                    if dates:
                        data_emissao = dates[0]

        # Linha de funcionário.
        if any(fix_text(c).lower().replace(":", "") == "funcionário" for c in row):
            finalize_record(current_record, output_rows, max_marcacoes)
            current_record = None

            funcionario = value_after_label(row, "Funcionário:")
            matricula = value_after_label(row, "Matrícula:")
            pis = value_after_label(row, "PIS:")
            data_admissao = value_after_label(row, "Data de admissão:")
            continue

        # Linha de departamento/cargo.
        if any(fix_text(c).lower().replace(":", "") == "departamento" for c in row):
            departamento = value_after_label(row, "Departamento:")
            cargo = value_after_label(row, "Cargo:")
            cpf = value_after_label(row, "CPF:")
            continue

        # Cabeçalho superior.
        if first.lower() == "dia":
            # A próxima linha normalmente é o cabeçalho inferior.
            idx = line_number  # rows é zero-based; line_number é 1-based
            if idx < len(rows):
                header_bottom = rows[idx]
                header_pos = extract_header_positions(row, header_bottom)
            continue

        # Nova linha de dia.
        if is_day_row(row):
            finalize_record(current_record, output_rows, max_marcacoes)

            data, dia_semana = split_date_weekday(row[0])

            turno_indices = header_pos.get("turno_indices", [])
            turno_previsto = []
            for idx in turno_indices:
                turno_previsto.append(get_time_near(row, idx, window=2))

            while len(turno_previsto) < 4:
                turno_previsto.append("")

            current_record = {
                "Empresa": empresa,
                "Período Início": periodo_inicio,
                "Período Fim": periodo_fim,
                "Data Emissão": data_emissao,
                "Funcionário": funcionario,
                "Matrícula": matricula,
                "PIS": pis,
                "Data de Admissão": data_admissao,
                "Departamento": departamento,
                "Cargo": cargo,
                "CPF": cpf,
                "Data": data,
                "Dia Semana": dia_semana,
                "Marcador": extract_marcador(row, header_pos),
                "Turno Previsto 1": turno_previsto[0],
                "Turno Previsto 2": turno_previsto[1],
                "Turno Previsto 3": turno_previsto[2],
                "Turno Previsto 4": turno_previsto[3],
                "Turno Total": first_real_time_between(
                    row,
                    header_pos.get("turno_total", 0),
                    header_pos.get("jornada_realizada", len(row))
                ),
                "Normais Diurnas": first_real_time_between(
                    row,
                    max((header_pos.get("normais") or 0) - 1, 0),
                    header_pos.get("extras", len(row))
                ),
                "Normais Noturna": "",
                "Extras Diurnas": first_real_time_between(
                    row,
                    max((header_pos.get("extras") or 0) - 1, 0),
                    header_pos.get("ausencias", len(row))
                ),
                "Extras Noturnas": "",
                "Ausências Diurnas": first_real_time_between(
                    row,
                    max((header_pos.get("ausencias") or 0) - 1, 0),
                    header_pos.get("intrajor", len(row))
                ),
                "Ausências Noturnas": "",
                "Intrajornada": get_index_value(row, header_pos.get("intrajor")),
                "Interjornada": get_index_value(row, header_pos.get("interjor")),
                "Banco de Horas": get_index_value(row, header_pos.get("bh")),
                "Observações": extract_observacoes(row, header_pos),
                "Linha Original": line_number,
            }

            append_markings(current_record, row, header_pos, max_marcacoes)
            continue

        # Linha de continuação de marcações do mesmo dia.
        if current_record and is_continuation_row(row, header_pos):
            append_markings(current_record, row, header_pos, max_marcacoes)

            obs = extract_observacoes(row, header_pos)
            if obs:
                if current_record.get("Observações"):
                    current_record["Observações"] += " | " + obs
                else:
                    current_record["Observações"] = obs
            continue

    finalize_record(current_record, output_rows, max_marcacoes)

    base_columns = [
        "Empresa",
        "Período Início",
        "Período Fim",
        "Data Emissão",
        "Funcionário",
        "Matrícula",
        "PIS",
        "Data de Admissão",
        "Departamento",
        "Cargo",
        "CPF",
        "Data",
        "Dia Semana",
        "Marcador",
        "Turno Previsto 1",
        "Turno Previsto 2",
        "Turno Previsto 3",
        "Turno Previsto 4",
        "Turno Total",
    ]

    marcacao_columns = [f"Marcação {i + 1}" for i in range(max_marcacoes)]

    end_columns = [
        "Normais Diurnas",
        "Normais Noturna",
        "Extras Diurnas",
        "Extras Noturnas",
        "Ausências Diurnas",
        "Ausências Noturnas",
        "Intrajornada",
        "Interjornada",
        "Banco de Horas",
        "Observações",
        "Linha Original",
    ]

    columns = base_columns + marcacao_columns + end_columns

    with output_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=columns, delimiter=",", extrasaction="ignore")
        writer.writeheader()
        writer.writerows(output_rows)

    return len(output_rows), output_path


def main():
    parser = argparse.ArgumentParser(
        description="Limpa e padroniza relatório de frequência em CSV, consolidando marcações extras na mesma linha."
    )
    parser.add_argument("entrada", help="Caminho do CSV/TXT bruto")
    parser.add_argument("saida", help="Caminho do CSV limpo que será gerado")
    parser.add_argument(
        "--max-marcacoes",
        type=int,
        default=12,
        help="Quantidade máxima de marcações por dia. Padrão: 12"
    )

    args = parser.parse_args()

    total, saida = limpar_relatorio(args.entrada, args.saida, args.max_marcacoes)

    print(f"Arquivo gerado com sucesso: {saida}")
    print(f"Linhas de ponto exportadas: {total}")


if __name__ == "__main__":
    main()
