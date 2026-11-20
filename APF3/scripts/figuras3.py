# Gráficos del APF3: mapa estratégico y organigramas.
import sys
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

PPT = len(sys.argv) > 1 and sys.argv[1] == 'ppt'
plt.rcParams['font.family'] = 'DejaVu Sans' if PPT else 'DejaVu Serif'
SUF = '_ppt' if PPT else ''
OSC, CLA, ACE, BORDE = ('#071D36', '#DFF3FA', '#2E86C1', '#102A43') if PPT else ('#1f3b73', '#EEF3F9', '#c8102e', '#333333')


def caja(ax, x, y, w, h, txt, fill='white', color='black', bold=False, size=9, edge=BORDE):
    ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h, boxstyle='round,pad=0.01,rounding_size=0.06', fc=fill, ec=edge, lw=1))
    ax.text(x, y, txt, ha='center', va='center', fontsize=size, color=color, weight='bold' if bold else 'normal', wrap=True)


def flecha(ax, a, b):
    ax.annotate('', xy=b, xytext=a, arrowprops=dict(arrowstyle='-|>', color=ACE, lw=1.3, shrinkA=2, shrinkB=2))


# ---------- Mapa estratégico ----------
fig, ax = plt.subplots(figsize=(10, 7))
ax.set_xlim(0, 10); ax.set_ylim(0, 8); ax.axis('off')
filas = [('Financiera', 7), ('Clientes', 5.15), ('Procesos\ninternos', 3.3), ('Aprendizaje y\ncrecimiento', 1.45)]
for n, y in filas:
    ax.add_patch(FancyBboxPatch((0.05, y - 0.8), 9.9, 1.6, boxstyle='round,pad=0.01,rounding_size=0.08', fc=CLA, ec='none'))
    ax.text(0.85, y, n, ha='center', va='center', fontsize=10, weight='bold', color=OSC)
W, H = 2.25, 0.95
P = {
    'FI1': (5.55, 7, 'FI1. Aumentar los\ningresos al mes'), 'FI2': (8.4, 7, 'FI2. Ganar con\nnuevos programas'),
    'CL1': (5.55, 5.15, 'CL1. Más alumnos en\noctubre y noviembre'), 'CL2': (3.0, 5.15, 'CL2. Más interesados\nque compran'), 'CL3': (8.4, 5.15, 'CL3. Opciones por grado\ny en grupo'),
    'PI1': (5.55, 3.3, 'PI1. Clases sin faltas,\ntardanzas ni cámara\napagada'), 'PI2': (8.4, 3.3, 'PI2. Programas de\nnivelación listos'), 'PI3': (3.0, 3.3, 'PI3. Campañas con\ndiagnóstico gratuito'),
    'AC1': (5.55, 1.45, 'AC1. Tablero mensual\npor profesor'), 'AC2': (3.0, 1.45, 'AC2. Retener a los\nbuenos profesores'), 'AC3': (8.4, 1.45, 'AC3. Profesores formados\nen grupo y nivelación'),
}
for k, (x, y, t) in P.items():
    caja(ax, x, y, W, H, t, fill=OSC, color='white', size=8.5, edge=OSC)
rel = [('AC1', 'PI1'), ('AC2', 'PI1'), ('AC3', 'PI2'), ('PI1', 'CL1'), ('PI2', 'CL3'), ('PI3', 'CL2'), ('CL2', 'CL1'), ('CL3', 'CL1'), ('CL1', 'FI1'), ('CL3', 'FI2'), ('FI2', 'FI1')]
for a, b in rel:
    xa, ya, _ = P[a]; xb, yb, _ = P[b]
    if ya == yb:
        s = 1 if xb > xa else -1
        flecha(ax, (xa + s * W / 2, ya), (xb - s * W / 2, yb))
    else:
        flecha(ax, (xa, ya + H / 2), (xb, yb - H / 2))
fig.tight_layout(); fig.savefig(f'figuras/mapa_bsc{SUF}.png', dpi=200)


# ---------- Organigramas ----------
def organigrama(nombre, nodos, lineas, alto=6):
    fig, ax = plt.subplots(figsize=(10, alto))
    ax.set_xlim(0, 10); ax.set_ylim(0, alto); ax.axis('off')
    for a, b in lineas:
        xa, ya = nodos[a][0], nodos[a][1]; xb, yb = nodos[b][0], nodos[b][1]
        ym = (ya + yb) / 2
        ax.plot([xa, xa, xb, xb], [ya, ym, ym, yb], color=BORDE, lw=1, zorder=0)
    for k, (x, y, w, h, t, estilo) in nodos.items():
        fill, color, edge = {'osc': (OSC, 'white', OSC), 'cla': (CLA, OSC, CLA), 'nuevo': ('white', OSC, ACE), 'base': ('white', OSC, BORDE)}[estilo]
        caja(ax, x, y, w, h, t, fill=fill, color=color, edge=edge, size=8.5, bold=estilo == 'osc')
    fig.tight_layout(); fig.savefig(f'figuras/{nombre}{SUF}.png', dpi=200)


organigrama('organigrama_actual', {
    'G': (5, 5.2, 2.8, 0.8, 'Gerencia\n2 personas', 'osc'),
    'M': (1.7, 3.5, 2.6, 0.9, 'Marketing\nlo dirige el fundador', 'cla'),
    'V': (5, 3.5, 2.6, 0.9, 'Comercial y ventas\nlíder de ventas, 6 personas', 'cla'),
    'A': (8.3, 3.5, 2.6, 0.9, 'Académica\ngrupo de dirección, 5 personas\nsin líder fijo', 'cla'),
    'P': (8.3, 1.6, 2.6, 0.9, '27 profesores', 'base'),
}, [('G', 'M'), ('G', 'V'), ('G', 'A'), ('A', 'P')], alto=6)

organigrama('organigrama_nuevo', {
    'G': (5, 7.2, 2.8, 0.8, 'Gerencia general', 'osc'),
    'M': (1.45, 5.5, 2.5, 1.0, 'Marketing y datos\nE2 · campañas y\ndiagnóstico', 'nuevo'),
    'V': (4.1, 5.5, 2.3, 1.0, 'Ventas\nE2 · seguimiento\na interesados', 'cla'),
    'A': (7.0, 5.5, 2.8, 1.0, 'Coordinación académica\nlíder fijo', 'nuevo'),
    'C': (4.6, 3.3, 2.1, 1.05, 'Calidad docente\nE5 · tablero,\nbono y rotación', 'nuevo'),
    'D': (7.0, 3.3, 2.0, 1.05, 'Diseño de\nprogramas\nE1 · E10', 'nuevo'),
    'S': (9.1, 3.3, 1.7, 1.05, 'Asesores\nacadémicos', 'cla'),
    'P': (4.6, 1.4, 2.1, 0.9, '27 profesores', 'base'),
}, [('G', 'M'), ('G', 'V'), ('G', 'A'), ('A', 'C'), ('A', 'D'), ('A', 'S'), ('C', 'P')], alto=8)
