from __future__ import annotations

import json
import tempfile
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

from .reconciler import AgendaItem, CRMRecord, export_excel_report, reconcile_agenda
from .totp_sync import generate_totp_compensated

STUDIO_HTML = """<!doctype html>
<html lang="pt-BR" data-theme="dark">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Portal Agenda Reconciler — 2FA Clock-Drift &amp; 4-Sheet Excel</title>
  <link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600&family=Plus+Jakarta+Sans:wght@400;600;700;800&display=swap" rel="stylesheet" />
  <style>
    :root, [data-theme="dark"] {
      --bg: #050a0f; --surface: #0b1219; --surface-2: #111b24; --surface-hover: #172430;
      --text: #f0f7fc; --muted: #9ab3c8; --accent: #d4f1ff; --accent-ink: #050a0f;
      --accent-soft: rgba(212, 241, 255, 0.11); --accent-border: rgba(212, 241, 255, 0.34);
      --border: rgba(212, 241, 255, 0.14);
      --warn: #fbbf24; --ok: #34d399; --purple: #c084fc; --err: #f87171;
    }
    [data-theme="light"] {
      --bg: #f0f7fc; --surface: #ffffff; --surface-2: #e3eef6; --surface-hover: #d5e5f1;
      --text: #050a0f; --muted: #395064; --accent: #050a0f; --accent-ink: #f0f7fc;
      --accent-soft: rgba(5, 10, 15, 0.08); --accent-border: rgba(5, 10, 15, 0.3);
      --border: rgba(5, 10, 15, 0.14);
      --warn: #b45309; --ok: #047857; --purple: #6b21a8; --err: #b91c1c;
    }
    * { box-sizing: border-box; }
    body { margin: 0; font-family: "Plus Jakarta Sans", system-ui, sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; }
    .shell-header {
      padding: 18px max(24px, calc((100% - 1180px) / 2)); border-bottom: 1px solid var(--border);
      background: var(--surface); display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 1rem;
    }
    .brand { display: flex; align-items: center; gap: 0.85rem; }
    .logo { width: 38px; height: 38px; border-radius: 10px; background: var(--accent); color: var(--accent-ink); display: grid; place-items: center; font-weight: 900; font-size: 0.85rem; }
    .brand-title { font-weight: 800; font-size: 1.08rem; }
    .brand-sub { font-size: 0.78rem; color: var(--muted); }
    .btn {
      display: inline-flex; align-items: center; justify-content: center; gap: 0.4rem;
      padding: 0.54rem 0.96rem; border-radius: 9px; border: 1px solid var(--border);
      background: var(--surface-2); color: var(--text); font: inherit; font-size: 0.82rem; font-weight: 700; cursor: pointer; text-decoration: none;
    }
    .btn:hover { background: var(--surface-hover); border-color: var(--accent-border); }
    .btn-primary, .tab-btn.active { background: var(--accent); color: var(--accent-ink); border-color: var(--accent); }
    .shell-body { width: min(1180px, calc(100% - 48px)); margin: 0 auto; padding: 24px 0 44px; }
    .totp-strip {
      display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 1rem;
      padding: 14px 18px; border-radius: 13px; background: var(--surface-2); border: 1px solid var(--border);
      border-left: 4px solid var(--accent); margin-bottom: 20px;
    }
    .totp-code { font-family: "JetBrains Mono", monospace; font-size: 1.4rem; font-weight: 800; color: var(--accent); letter-spacing: 0.1em; }
    .kpi-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 0.9rem; margin-bottom: 20px; }
    .kpi { background: var(--surface-2); border: 1px solid var(--border); border-radius: 12px; padding: 0.95rem 1.1rem; }
    .kpi-label { font-size: 0.72rem; color: var(--muted); font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; }
    .kpi-val { font-size: 1.65rem; font-weight: 800; margin-top: 0.2rem; }
    .tabs { display: flex; flex-wrap: wrap; gap: 0.5rem; margin-bottom: 14px; }
    .card { background: var(--surface-2); border: 1px solid var(--border); border-radius: 14px; padding: 1.25rem; }
    table { width: 100%; border-collapse: collapse; font-size: 0.84rem; }
    th, td { text-align: left; padding: 0.75rem 0.6rem; border-bottom: 1px solid var(--border); }
    th { font-size: 0.71rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.06em; }
    .mono { font-family: "JetBrains Mono", monospace; font-size: 0.76rem; color: var(--muted); }
    @media (max-width: 900px) { .kpi-grid { grid-template-columns: repeat(2, 1fr); } }
  </style>
</head>
<body>
  <header class="shell-header">
    <div class="brand">
      <span class="logo">REC</span>
      <div>
        <div class="brand-title">Portal Agenda × CRM Reconciler</div>
        <div class="brand-sub">2FA TOTP Compensado por Clock Drift · Classificação 4 Abas · Exportação Excel</div>
      </div>
    </div>
    <div style="display:flex;gap:0.55rem;align-items:center">
      <a href="/api/export.xlsx" class="btn btn-primary">⬇ Baixar Relatório Excel (4 Abas)</a>
      <button id="themeToggle" type="button" class="btn" style="border-radius:999px">
        <span id="themeIcon">☀</span> <span id="themeLabel">Modo Claro</span>
      </button>
    </div>
  </header>

  <main class="shell-body">
    <section class="totp-strip">
      <div>
        <div style="font-size:0.76rem;color:var(--muted);font-weight:700;text-transform:uppercase">Token 2FA Portal Regional (Compensado por Offset HTTP Date)</div>
        <div style="display:flex;align-items:baseline;gap:0.8rem;margin-top:3px">
          <span class="totp-code" id="totpCode">------</span>
          <span class="mono" id="totpRem">expira em --s</span>
        </div>
      </div>
      <div style="display:flex;align-items:center;gap:0.65rem">
        <label class="mono">Simular Desvio de Relógio (s):</label>
        <input id="driftInput" type="number" value="34" style="width:76px;padding:0.45rem;border-radius:8px;border:1px solid var(--border);background:var(--bg);color:var(--text);font-family:'JetBrains Mono',monospace" />
        <button id="refreshTotpBtn" type="button" class="btn">Recalcular TOTP</button>
      </div>
    </section>

    <section class="kpi-grid">
      <div class="kpi" style="border-left:4px solid var(--warn)"><div class="kpi-label">Pendente Agendar</div><div class="kpi-val" style="color:var(--warn)" id="kPending">--</div></div>
      <div class="kpi" style="border-left:4px solid var(--ok)"><div class="kpi-label">Agendados no CRM</div><div class="kpi-val" style="color:var(--ok)" id="kScheduled">--</div></div>
      <div class="kpi" style="border-left:4px solid var(--purple)"><div class="kpi-label">Parceiros Externos (PRT)</div><div class="kpi-val" style="color:var(--purple)" id="kPartners">--</div></div>
      <div class="kpi" style="border-left:4px solid var(--err)"><div class="kpi-label">Fora do CRM (Cadastrar)</div><div class="kpi-val" style="color:var(--err)" id="kMissing">--</div></div>
    </section>

    <section class="card">
      <div class="tabs">
        <button type="button" class="btn tab-btn active" data-tab="pending">1 · Pendente Agendar</button>
        <button type="button" class="btn tab-btn" data-tab="scheduled">2 · Agendados no CRM</button>
        <button type="button" class="btn tab-btn" data-tab="partners">3 · Parceiros Externos</button>
        <button type="button" class="btn tab-btn" data-tab="missing_in_crm">4 · Fora do CRM</button>
      </div>
      <div style="overflow-x:auto">
        <table>
          <thead id="tableHead"></thead>
          <tbody id="tableBody"></tbody>
        </table>
      </div>
    </section>
  </main>

  <script>
    const root = document.documentElement;
    function applyTheme(t) {
      root.setAttribute("data-theme", t);
      localStorage.setItem("rec-theme", t);
      document.getElementById("themeIcon").textContent = t === "light" ? "☾" : "☀";
      document.getElementById("themeLabel").textContent = t === "light" ? "Modo Escuro" : "Modo Claro";
    }
    applyTheme(localStorage.getItem("rec-theme") || "dark");
    document.getElementById("themeToggle").addEventListener("click", () => {
      applyTheme(root.getAttribute("data-theme") === "dark" ? "light" : "dark");
    });

    let state = null;
    let currentTab = "pending";

    async function loadData() {
      const drift = document.getElementById("driftInput").value || "0";
      state = await (await fetch(`/api/reconcile?drift=${encodeURIComponent(drift)}`)).json();
      document.getElementById("totpCode").textContent = state.totp.code;
      document.getElementById("totpRem").textContent = `expira em ${state.totp.remaining}s (offset ${drift}s aplicado)`;
      document.getElementById("kPending").textContent = state.buckets.pending.length;
      document.getElementById("kScheduled").textContent = state.buckets.scheduled.length;
      document.getElementById("kPartners").textContent = state.buckets.partners.length;
      document.getElementById("kMissing").textContent = state.buckets.missing_in_crm.length;
      renderTab();
    }

    function renderTab() {
      const rows = (state && state.buckets[currentTab]) || [];
      const thead = document.getElementById("tableHead");
      const tbody = document.getElementById("tableBody");
      if (!rows.length) {
        thead.innerHTML = "";
        tbody.innerHTML = '<tr><td class="mono">Nenhum registro nesta aba.</td></tr>';
        return;
      }
      const cols = Object.keys(rows[0]).filter(c => c !== "Link");
      thead.innerHTML = "<tr>" + cols.map(c => `<th>${c}</th>`).join("") + "</tr>";
      tbody.innerHTML = rows.map(r =>
        "<tr>" + cols.map(c => `<td>${r[c]}</td>`).join("") + "</tr>"
      ).join("");
    }

    document.querySelectorAll("[data-tab]").forEach(btn => {
      btn.addEventListener("click", () => {
        document.querySelectorAll("[data-tab]").forEach(b => b.classList.remove("active"));
        btn.classList.add("active");
        currentTab = btn.dataset.tab;
        renderTab();
      });
    });
    document.getElementById("refreshTotpBtn").addEventListener("click", loadData);
    loadData();
  </script>
</body>
</html>
"""


def get_demo_dataset() -> tuple[list[AgendaItem], dict[str, CRMRecord]]:
    agenda = [
        AgendaItem("Regional SP", "2026-11-27 09:30", "REF-2026-001", "Vistoria Técnica", "Alfa Indústria S.A.", "Unidade Sul"),
        AgendaItem("Regional PR", "2026-11-27 11:00", "REF-2026-005", "Inspeção de Qualidade", "Épsilon Energia", "Unidade Curitiba"),
        AgendaItem("Regional SP", "2026-11-27 14:00", "REF-2026-002", "Reunião de Alinhamento", "Beta Serviços Ltda", "Unidade Centro"),
        AgendaItem("Regional RJ", "2026-11-28 10:30", "REF-2026-003", "Entrega Técnica", "Gama Comércio", "Parceiro Externo"),
        AgendaItem("Regional MG", "2026-11-28 16:00", "REF-2026-004", "Auditoria Presencial", "Delta Logística", "Unidade BH"),
        AgendaItem("Regional SC", "2026-11-29 09:00", "REF-2026-006", "Homologação Operacional", "Zeta Transportes", "Unidade Joinville"),
    ]
    crm = {
        "REF-2026-001": CRMRecord("501", "Alfa Indústria — Contrato Anual", "Ana Costa", "Em Análise", "", "https://crm.example.com/deals/501"),
        "REF-2026-005": CRMRecord("505", "Épsilon Energia — Implantação", "Diego Rocha", "Documentação OK", "", "https://crm.example.com/deals/505"),
        "REF-2026-002": CRMRecord("502", "Beta Serviços — Expansão", "Bruno Lima", "Visita Agendada", "2026-11-27 14:00", "https://crm.example.com/deals/502"),
        "REF-2026-003": CRMRecord("503", "[PARCEIRO] - PRT4090 - Gama Comércio", "Carla Souza", "Em Andamento", "", "https://crm.example.com/deals/503"),
    }
    return agenda, crm


def run_studio(port: int = 8022, host: str = "127.0.0.1") -> None:
    agenda, crm = get_demo_dataset()
    result = reconcile_agenda(agenda, crm)

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, fmt: str, *args) -> None:
            pass

        def do_GET(self) -> None:
            parsed = urlparse(self.path)
            qs = parse_qs(parsed.query)
            if parsed.path == "/api/reconcile":
                drift = float((qs.get("drift") or ["0"])[0])
                code, rem = generate_totp_compensated("JBSWY3DPEHPK3PXP", clock_offset_sec=drift)
                payload: dict[str, Any] = {
                    "totp": {"code": code, "remaining": rem},
                    "buckets": {
                        "pending": result.pending,
                        "scheduled": result.scheduled,
                        "partners": result.partners,
                        "missing_in_crm": result.missing_in_crm,
                    },
                }
                raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(raw)))
                self.end_headers()
                self.wfile.write(raw)
                return
            if parsed.path == "/api/export.xlsx":
                try:
                    with tempfile.TemporaryDirectory() as tmp:
                        out = export_excel_report(result, Path(tmp) / "reconciliacao.xlsx")
                        xlsx_bytes = out.read_bytes()
                    self.send_response(200)
                    self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    self.send_header("Content-Disposition", 'attachment; filename="reconciliacao_agenda_crm.xlsx"')
                    self.send_header("Content-Length", str(len(xlsx_bytes)))
                    self.end_headers()
                    self.wfile.write(xlsx_bytes)
                except ImportError:
                    self.send_error(501, "Instale openpyxl para baixar o arquivo .xlsx")
                return

            raw_html = STUDIO_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(raw_html)))
            self.end_headers()
            self.wfile.write(raw_html)

    httpd = HTTPServer((host, port), Handler)
    print(f"Portal Agenda Reconciler Studio em http://{host}:{port}")
    httpd.serve_forever()
