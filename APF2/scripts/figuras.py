import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
plt.rcParams['font.family'] = 'DejaVu Serif'

# PEYEA
fig, ax = plt.subplots(figsize=(6, 6))
ax.axhline(0, color='black', lw=1); ax.axvline(0, color='black', lw=1)
ax.set_xlim(-6, 6); ax.set_ylim(-6, 6)
ax.set_xticks(range(-6, 7)); ax.set_yticks(range(-6, 7))
ax.grid(color='#dddddd', lw=0.5)
for x, y, t in [(-3, 4.5, 'Conservador'), (3, 4.5, 'Agresivo'), (-3, -4.5, 'Defensivo'), (3, -4.5, 'Competitivo')]:
    ax.text(x, y, t, ha='center', fontsize=12, weight='bold', color='#444')
ax.fill([0, 3.5, 0, -3.33], [4.40, 0, -3.33, 0], color='#c8102e', alpha=0.08)
ax.plot([0, 3.5, 0, -3.33, 0], [4.40, 0, -3.33, 0, 4.40], color='#c8102e', lw=1)
ax.annotate('', xy=(0.17, 1.07), xytext=(0, 0), arrowprops=dict(arrowstyle='-|>', color='#1f3b73', lw=2.2))
ax.plot([0.17, 0.17], [0, 1.07], ls='--', color='#c8102e', lw=1)
ax.plot([0, 0.17], [1.07, 1.07], ls='--', color='#c8102e', lw=1)
ax.text(0.35, 1.15, '(0.17, 1.07)', fontsize=10, color='#1f3b73')
ax.text(0.15, 5.45, 'Fortaleza financiera FF\nAlto', fontsize=9)
ax.text(0.15, -5.85, 'Estabilidad del entorno EE\nBajo', fontsize=9)
ax.text(5.9, 0.25, 'Fortaleza de la\nindustria FI · Alto', fontsize=9, ha='right')
ax.text(-5.9, 0.25, 'Ventaja\ncompetitiva VC · Bajo', fontsize=9)
fig.tight_layout(); fig.savefig('figuras/peyea.png', dpi=200)

# BCG
fig, ax = plt.subplots(figsize=(6, 5.2))
ax.set_xlim(0, 2); ax.set_ylim(0, 2)
ax.axhline(1, color='black', lw=1); ax.axvline(1, color='black', lw=1)
ax.set_xticks([0.5, 1.5]); ax.set_xticklabels(['Alta', 'Baja'])
ax.set_yticks([0.5, 1.5]); ax.set_yticklabels(['Bajo', 'Alto'])
ax.set_xlabel('Cuota de mercado relativa'); ax.set_ylabel('Tasa de crecimiento del mercado')
for x, y, t in [(0.5, 1.95, 'Estrella'), (1.5, 1.95, 'Interrogante'), (0.5, 0.95, 'Vaca'), (1.5, 0.95, 'Perro')]:
    ax.text(x, y, t, ha='center', va='top', fontsize=12, weight='bold', color='#444')
for x, y, s, t in [(0.33, 0.45, 60, 'Clases individuales\n60 %'), (0.72, 0.22, 30, 'Programas de verano\n30 %'), (1.5, 1.4, 10, 'Talleres y empresas\n10 %')]:
    ax.scatter(x, y, s=s * 40, color='#c8102e', alpha=0.3, edgecolor='#c8102e')
    ax.text(x, y, t, ha='center', va='center', fontsize=9)
fig.tight_layout(); fig.savefig('figuras/bcg.png', dpi=200)
