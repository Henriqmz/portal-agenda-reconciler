# Portal Agenda Reconciler (2FA + CRM + Excel)

Ferramenta em Python para coleta consolidada de agendas em múltiplos portais regionais (com geração de código **2FA TOTP com compensação automática de desvio de relógio**), cruzamento automático com um CRM e exportação em **planilha Excel multiabas estilizada (`openpyxl`)**.

> **100% demonstrável localmente:** execute `python -m src.cli --demo` para gerar uma planilha Excel pronta em `output/reconciliacao_agenda_crm.xlsx` usando dados sintéticos de clínicas/unidades regionais e parceiros externos.

## Recursos

- **TOTP com Compensação de Clock Drift:** calcula automaticamente a diferença entre o relógio local e o horário UTC de referência, evitando falhas de 2FA em máquinas dessincronizadas.
- **Cruzamento Inteligente com CRM:** identifica registros pendentes de agendamento, já agendados, pertencentes a parceiros externos (por regra de prefixo/tag no título) e ausentes do CRM.
- **Relatório Excel Profissional (4 Abas):**
  - `Pendente Agendar` (laranja — ação imediata)
  - `Agendados no CRM` (azul — conferidos)
  - `Parceiros Externos` (roxo — separados para não poluir pendências)
  - `Fora do CRM` (vermelho — precisam ser cadastrados)

## Como Executar

```bash
pip install -r requirements.txt

# Gera a planilha de reconciliação em modo demonstração
python -m src.cli --demo

# Roda os testes automatizados
pytest -v
```
