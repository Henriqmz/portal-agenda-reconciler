from __future__ import annotations

import argparse
from pathlib import Path

from .reconciler import AgendaItem, CRMRecord, export_excel_report, reconcile_agenda
from .server import run_studio
from .totp_sync import generate_totp_compensated


def run_demo(output_path: Path = Path("output/reconciliacao_agenda_crm.xlsx")) -> None:
    agenda = [
        AgendaItem("Regional SP", "2026-11-25 09:30", "REF-2026-001", "Vistoria Técnica", "Alfa Indústria", "Unidade Sul"),
        AgendaItem("Regional SP", "2026-11-25 14:00", "REF-2026-002", "Reunião de Alinhamento", "Beta Serviços", "Unidade Centro"),
        AgendaItem("Regional RJ", "2026-11-26 11:00", "REF-2026-003", "Entrega Técnica", "Gama Comércio", "Parceiro Externo"),
        AgendaItem("Regional MG", "2026-11-27 16:00", "REF-2026-004", "Auditoria Presencial", "Delta Logística", "Unidade BH"),
    ]
    crm = {
        "REF-2026-001": CRMRecord("501", "Alfa Indústria — Contrato Anual", "Ana Costa", "Em Análise", "", "https://crm.example.com/deals/501"),
        "REF-2026-002": CRMRecord("502", "Beta Serviços — Expansão", "Bruno Lima", "Visita Agendada", "2026-11-25 14:00", "https://crm.example.com/deals/502"),
        "REF-2026-003": CRMRecord("503", "[PARCEIRO] - PRT4090 - Gama Comércio", "Carla Souza", "Em Andamento", "", "https://crm.example.com/deals/503"),
    }

    code, rem = generate_totp_compensated("JBSWY3DPEHPK3PXP", clock_offset_sec=0.0)
    res = reconcile_agenda(agenda, crm)
    saved = export_excel_report(res, output_path)
    print(f"2FA Demo TOTP: {code} (restam {rem}s)")
    print(
        f"Reconciliação concluída -> Pendentes: {len(res.pending)} | Agendados: {len(res.scheduled)} | "
        f"Parceiros: {len(res.partners)} | Fora do CRM: {len(res.missing_in_crm)}"
    )
    print(f"Planilha salva em: {saved}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Portal Agenda Reconciler")
    parser.add_argument("--demo", action="store_true", help="Gera planilha demo com dados fictícios")
    parser.add_argument("--serve", action="store_true", help="Inicia o painel web interativo de reconciliação")
    parser.add_argument("--port", type=int, default=8022, help="Porta do painel web (padrão: 8022)")
    args = parser.parse_args()
    if args.serve:
        run_studio(port=args.port)
    elif args.demo:
        run_demo()
    else:
        print("Use --demo para gerar planilha ou --serve para abrir o painel web.")


if __name__ == "__main__":
    main()
