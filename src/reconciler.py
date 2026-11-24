from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class AgendaItem:
    unit: str
    scheduled_at: str
    reference_code: str
    event_type: str
    customer_name: str
    counterparty: str


@dataclass(frozen=True)
class CRMRecord:
    card_id: str
    title: str
    owner: str
    stage: str
    scheduled_in_crm: str
    url: str


@dataclass
class ReconciliationResult:
    pending: list[dict[str, Any]]
    scheduled: list[dict[str, Any]]
    partners: list[dict[str, Any]]
    missing_in_crm: list[dict[str, Any]]


def reconcile_agenda(
    agenda: list[AgendaItem],
    crm_index: dict[str, CRMRecord],
    partner_pattern: str = r"\bPRT\s*\d+",
) -> ReconciliationResult:
    """Classifies portal appointments against CRM records into 4 actionable buckets."""
    pending, scheduled, partners, missing = [], [], [], []
    rx = re.compile(partner_pattern, re.IGNORECASE)

    for item in agenda:
        base = {
            "Unidade": item.unit,
            "Data/Hora Portal": item.scheduled_at,
            "Código Referência": item.reference_code,
            "Tipo": item.event_type,
            "Cliente": item.customer_name,
            "Contraparte": item.counterparty,
        }
        card = crm_index.get(item.reference_code)
        if not card:
            missing.append({**base, "Status": "NÃO LOCALIZADO NO CRM (Cadastrar)"})
            continue

        if rx.search(card.title):
            partners.append(
                {
                    **base,
                    "Título CRM": card.title,
                    "Responsável": card.owner,
                    "Etapa": card.stage,
                    "Link": card.url,
                }
            )
            continue

        is_booked = bool(card.scheduled_in_crm) or "agendad" in card.stage.lower()
        if is_booked:
            scheduled.append(
                {
                    **base,
                    "Responsável": card.owner,
                    "Etapa": card.stage,
                    "Data no CRM": card.scheduled_in_crm or "Confirmado pela etapa",
                    "Link": card.url,
                }
            )
        else:
            pending.append(
                {
                    **base,
                    "Responsável": card.owner,
                    "Etapa": card.stage,
                    "Ação": f"Agendar para {item.scheduled_at} e atualizar etapa",
                    "Link": card.url,
                }
            )

    return ReconciliationResult(pending=pending, scheduled=scheduled, partners=partners, missing_in_crm=missing)


def export_excel_report(result: ReconciliationResult, destination: Path) -> Path:
    """Exports the 4 reconciliation buckets to a styled multi-sheet Excel workbook."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    destination.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()

    sheets_spec = [
        ("Pendente Agendar", "E26B00", result.pending, ["Unidade", "Data/Hora Portal", "Código Referência", "Tipo", "Cliente", "Contraparte", "Responsável", "Etapa", "Ação", "Link"]),
        ("Agendados no CRM", "1F4E79", result.scheduled, ["Unidade", "Data/Hora Portal", "Código Referência", "Tipo", "Cliente", "Contraparte", "Responsável", "Etapa", "Data no CRM", "Link"]),
        ("Parceiros Externos", "5B2C6F", result.partners, ["Unidade", "Data/Hora Portal", "Código Referência", "Tipo", "Cliente", "Contraparte", "Título CRM", "Responsável", "Etapa", "Link"]),
        ("Fora do CRM", "C00000", result.missing_in_crm, ["Unidade", "Data/Hora Portal", "Código Referência", "Tipo", "Cliente", "Contraparte", "Status"]),
    ]

    for idx, (title, color, rows, columns) in enumerate(sheets_spec):
        ws = wb.active if idx == 0 else wb.create_sheet(title)
        ws.title = title
        ws.append(columns)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill("solid", fgColor=color)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in rows:
            ws.append([row.get(col, "") for col in columns])
        for col_idx, col_name in enumerate(columns, start=1):
            col_letter = chr(64 + col_idx)
            ws.column_dimensions[col_letter].width = max(18, len(col_name) + 6)

    wb.save(destination)
    return destination
