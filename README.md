# 🛡️ FPS Dashboard 2026

Dashboard KPI automatica per Family Protection Specialist.  
Si aggiorna automaticamente ogni volta che carichi un nuovo file Excel.

---

## ⚡ Setup iniziale (una volta sola — 5 minuti)

### 1. Crea il repository su GitHub
1. Vai su [github.com](https://github.com) → **New repository**
2. Nome: `fps-dashboard` (o quello che vuoi)
3. Visibilità: **Private** (consigliato) o Public
4. Clicca **Create repository**

### 2. Carica tutti i file
Trascina nella pagina del repository questi file:
- `genera_dashboard.py`
- `requirements.txt`
- `docs/index.html`
- `.github/workflows/genera_dashboard.yml`
- Il tuo file Excel `Report_TOP_FPS_2026_...xlsx`

### 3. Attiva GitHub Pages
1. Nel repository: **Settings** → **Pages**
2. Source: **Deploy from a branch**
3. Branch: `main` — Folder: `/docs`
4. Clicca **Save**

### 4. Prima generazione
Vai su **Actions** → **Genera Dashboard FPS** → **Run workflow**

Dopo 1-2 minuti la dashboard è online all'indirizzo:  
`https://TUO-USERNAME.github.io/fps-dashboard/`

---

## 🔄 Aggiornamento (ogni volta che aggiorni i dati)

1. Apri il repository su GitHub nel browser
2. Clicca sul file Excel → **Delete** → Commit
3. Clicca **Add file** → **Upload files**
4. Trascina il nuovo Excel aggiornato
5. Clicca **Commit changes**

**GitHub Actions parte automaticamente** e in ~1 minuto la dashboard è aggiornata.

---

## 📁 Struttura del repository

```
fps-dashboard/
├── .github/
│   └── workflows/
│       └── genera_dashboard.yml   ← Automazione GitHub Actions
├── docs/
│   └── index.html                 ← Dashboard HTML (generata automaticamente)
├── genera_dashboard.py            ← Script Python che legge l'Excel
├── requirements.txt               ← Dipendenze Python
├── README.md                      ← Questo file
└── Report_TOP_FPS_2026_...xlsx    ← Il tuo file Excel (caricalo qui)
```

---

## ❓ Domande frequenti

**Il file Excel è pubblico?**  
Se il repository è **Private**, no — solo tu (e chi inviti) può vederlo.  
Se è **Public**, sì — usa Private per dati aziendali.

**Quanto costa?**  
GitHub Actions è **gratuito** per repository privati (fino a 2.000 minuti/mese).  
GitHub Pages è gratuito.

**Posso condividere la dashboard con i colleghi?**  
Sì — basta mandargli il link `https://username.github.io/fps-dashboard/`

**Il link è sempre lo stesso anche dopo gli aggiornamenti?**  
Sì, il link non cambia mai.
