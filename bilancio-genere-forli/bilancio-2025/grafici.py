# Grafici del Bilancio di genere 2025: legge dati_2025.json e scrive PNG in grafici/.
# Palette validata (slot 1-3 della palette di riferimento): donne blu, uomini arancio.
import json, os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

D = json.load(open(os.path.join(os.path.dirname(__file__), 'dati_2025.json')))
OUT = os.path.join(os.path.dirname(__file__), 'grafici')
F, M = '#2a78d6', '#eb6834'
PREV, CUR = '#a8a69e', '#4a3aa7'  # anno precedente recessivo, anno corrente viola (non riusa i colori di genere)
INK, MUTED, GRID, SURF = '#1F2A37', '#52514e', '#e5e4df', '#fcfcfb'
plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 11, 'text.color': INK,
                     'axes.labelcolor': MUTED, 'xtick.color': MUTED, 'ytick.color': MUTED})

def base(ax):
    ax.set_facecolor(SURF)
    for s in ('top', 'right', 'left'): ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_color(GRID)
    ax.yaxis.grid(True, color=GRID, linewidth=0.8); ax.set_axisbelow(True)
    ax.tick_params(length=0)

def barre_gruppi(cats, serie, nomi, colori, titolo, file, fmt='{:.0f}', ylabel=None):
    fig, ax = plt.subplots(figsize=(8, 4.2), dpi=200); fig.patch.set_facecolor(SURF); base(ax)
    n = len(serie); w = 0.8 / n
    for i, (vals, nome, col) in enumerate(zip(serie, nomi, colori)):
        xs = [x + (i - (n - 1) / 2) * w for x in range(len(cats))]
        bars = ax.bar(xs, vals, w * 0.92, color=col, label=nome, edgecolor=SURF, linewidth=1)
        for b, v in zip(bars, vals):
            ax.text(b.get_x() + b.get_width() / 2, b.get_height(), fmt.format(v).replace('.', ','), ha='center', va='bottom', fontsize=9, color=INK)
    ax.set_xticks(range(len(cats))); ax.set_xticklabels(cats)
    if ylabel: ax.set_ylabel(ylabel)
    ax.set_title(titolo, loc='left', fontsize=13, fontweight='bold', color=INK, pad=12)
    ax.legend(frameon=False, loc='upper left', ncol=n, bbox_to_anchor=(0, 1.0))
    ax.margins(y=0.15); fig.tight_layout()
    fig.savefig(os.path.join(OUT, file), facecolor=SURF); plt.close(fig)

I = D['iscritti']
ce = I['classi_eta_2025']
barre_gruppi(list(ce), [[v[0] for v in ce.values()], [v[1] for v in ce.values()]], ['Donne', 'Uomini'], [F, M],
             'Iscritti per classe d’età e genere – 31/12/2025', 'eta_genere.png')
C = D['cassa_cnpadc']['divario_reddito']
barre_gruppi(list(C), [[v[0] for v in C.values()], [v[1] for v in C.values()]], ['Redditi 2023', 'Redditi 2024'], [PREV, CUR],
             'Divario di reddito donne-uomini (%)', 'divario.png', fmt='{:.1f}%')
print('grafici ok')
