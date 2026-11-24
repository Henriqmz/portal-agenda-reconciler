from pathlib import Path

from src.reconciler import AgendaItem, CRMRecord, export_excel_report, reconcile_agenda
from src.totp_sync import generate_totp_compensated


def test_totp_clock_compensation():
    secret = "JBSWY3DPEHPK3PXP"
    base_ts = 1700000000.0
    # Simulated local clock 75 seconds ahead; offset=-75 restores exact reference code
    expected, _ = generate_totp_compensated(secret, clock_offset_sec=0.0, now_ts=base_ts)
    compensated, _ = generate_totp_compensated(secret, clock_offset_sec=-75.0, now_ts=base_ts + 75.0)
    assert compensated == expected


def test_reconciler_four_buckets_and_excel_export(tmp_path: Path):
    agenda = [
        AgendaItem("SP", "2026-11-25 10:00", "R-1", "Visita", "Cli A", "Emp A"),
        AgendaItem("SP", "2026-11-25 11:00", "R-2", "Visita", "Cli B", "Emp B"),
        AgendaItem("RJ", "2026-11-26 15:00", "R-3", "Visita", "Cli C", "Emp C"),
        AgendaItem("MG", "2026-11-27 09:00", "R-4", "Visita", "Cli D", "Emp D"),
    ]
    crm = {
        "R-1": CRMRecord("1", "Projeto Cli A", "Ana", "Triagem", "", "https://crm.example/1"),
        "R-2": CRMRecord("2", "Projeto Cli B", "Bruno", "Agendada", "2026-11-25 11:00", "https://crm.example/2"),
        "R-3": CRMRecord("3", "[EXT] - PRT9021 - Cli C", "Carla", "Triagem", "", "https://crm.example/3"),
    }
    res = reconcile_agenda(agenda, crm)
    assert len(res.pending) == 1
    assert len(res.scheduled) == 1
    assert len(res.partners) == 1
    assert len(res.missing_in_crm) == 1

    out = export_excel_report(res, tmp_path / "report.xlsx")
    assert out.exists() and out.stat().st_size > 1000
