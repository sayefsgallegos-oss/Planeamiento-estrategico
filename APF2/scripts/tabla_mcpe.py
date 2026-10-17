import sys
import os; sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mcpe import F, E
N = {
 'O1':'Alta conectividad de los hogares','O2':'Bajo logro de los escolares en matemática',
 'O3':'Posible mejora de la economía de las familias','O4':'Empresas que buscan educación financiera y apoyo educativo',
 'O5':'Tecnologías digitales e inteligencia artificial','A1':'Posible alza de precios y de costos',
 'A2':'Riesgo de fenómenos climáticos','A3':'Cambios de gobierno y nuevas normas',
 'A4':'Normas de protección de datos de menores','A5':'Competencia de bajo costo y sustitutos gratuitos',
 'A6':'Disputa por docentes calificados','F1':'Enseñanza personalizada 1 a 1',
 'F2':'Trayectoria y marca conocida','F3':'Experiencia en clases 100 % virtuales',
 'F4':'Seguimiento y comunicación con los padres','F5':'Oferta variada de servicios',
 'F6':'Cultura de innovación y mejora de procesos','D1':'Dependencia de los profesores',
 'D2':'Faltas, tardanzas y cámara apagada','D3':'Malas prácticas de algunos docentes',
 'D4':'Desorden en el área comercial','D5':'Falta de inscripciones','D6':'Datos e inteligencia artificial en proceso',
}
G = {'O':'Oportunidades','A':'Amenazas','F':'Fortalezas','D':'Debilidades'}
h = '| Factor clave | Peso | ' + ' | '.join(f'{e} PA | {e} TPA' for e in E) + ' |'
print(h); print('|' + '---|' * (2 + 2 * len(E)))
tot = [0] * len(E); last = None
for c, w, pa in F:
    if c[0] != last:
        print(f'| **{G[c[0]]}** |' + ' |' * (1 + 2 * len(E))); last = c[0]
    print(f'| {c}. {N[c]} | {w:.2f} | ' + ' | '.join(f'{p} | {w*p:.2f}' for p in pa) + ' |')
    for i, p in enumerate(pa): tot[i] += w * p
print('| **Total** | **2.00** | ' + ' | '.join(f' | **{t:.2f}**' for t in tot) + ' |')
