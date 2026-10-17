// Presentación del APF2, máximo 10 láminas, con el estilo del PPT del APF1.
const path = require('path');
const fs = require('fs');
const pptxgen = require('pptxgenjs');
const { applyTheme } = require(process.env.PPTX_SKILL + '/scripts/apply_theme.js');

const DIR = path.join(__dirname, '..');
const IMG = (f) => path.join(DIR, 'figuras', f);
const OUT = path.join(DIR, 'APF2_Presentacion_MatLab_Peru.pptx');

const THEME = {
  name: 'MatLab APF2',
  headFontFace: 'Arial',
  bodyFontFace: 'Arial',
  colors: {
    dk1: '102A43', lt1: 'FFFFFF', dk2: '071D36', lt2: 'F4F9FC',
    accent1: '65C7E8', accent2: '2E86C1', accent3: 'DFF3FA', accent4: '5B7083',
    accent5: 'C9D6DF', accent6: 'E8F1F6', hlink: '2E86C1', folHlink: '5B7083',
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
pres.author = 'Pierre Gallegos Mallqui';
pres.title = 'Diseño del plan estratégico de MatLab Perú';
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const LOGO_W = 2.0;
const LOGO_R = (() => { const b = fs.readFileSync(IMG('logo_utp_banda.png')); return b.readUInt32BE(20) / b.readUInt32BE(16); })();

const FOOT = 'PLANEAMIENTO ESTRATÉGICO  ·  MATLAB PERÚ';
const layout = (name, bg, titleColor, footColor) => pres.defineSlideMaster({
  title: name,
  background: { color: bg },
  objects: [
    { text: { text: FOOT, options: { x: 0.6, y: 6.95, w: 6, h: 0.3, fontSize: 9, color: footColor, charSpacing: 1, margin: 0 } } },
    { placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.72, w: 12.1, h: 0.75, fontSize: 30, bold: true, color: titleColor, margin: 0, valign: 'top', align: 'left' }, text: '' } },
  ],
  slideNumber: { x: 12.2, y: 6.95, w: 0.5, h: 0.3, fontSize: 9, color: footColor, align: 'right' },
});
layout('CLARO', C.background2, C.text1, C.accent4);
layout('OSCURO', C.text2, C.background1, C.accent1);

const kicker = (s, t, dark) => s.addText(t, { x: 0.6, y: 0.4, w: 8, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: dark ? C.accent1 : C.accent2, margin: 0, isTextBox: true, objectName: 'kicker' });
const card = (s, x, y, w, h, fill, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill }, objectName: name });
const txt = (s, t, o) => s.addText(t, { margin: 0, isTextBox: true, valign: 'top', ...o });
const slide = (master, section) => pres.addSlide({ masterName: master, sectionTitle: section });

// ---------- 1. Portada ----------
pres.addSection({ title: 'Inicio' });
{
  const s = slide('OSCURO', 'Inicio');
  s.addImage({ path: IMG('portada.jpg'), x: 7.6, y: 0, w: 5.733, h: 7.5, sizing: { type: 'cover', w: 5.733, h: 7.5 }, objectName: 'foto' });
  txt(s, 'PLANEAMIENTO ESTRATÉGICO  ·  SECCIÓN 20726', { x: 0.7, y: 1.0, w: 6.6, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: C.accent1 });
  txt(s, 'Avance de Proyecto Final 2', { x: 0.7, y: 1.55, w: 6.6, h: 0.45, fontSize: 18, color: C.background1 });
  txt(s, 'Diseño del plan estratégico de MatLab Perú', { x: 0.7, y: 2.05, w: 6.6, h: 2.0, fontSize: 36, bold: true, color: C.background1 });
  s.addImage({ path: IMG('logo_utp_banda.png'), x: 0.7, y: 4.35, w: LOGO_W, h: LOGO_W * LOGO_R, objectName: 'logo' });
  txt(s, 'Pierre Gallegos Mallqui  ·  Código U23211135', { x: 0.7, y: 5.45, w: 6.6, h: 0.35, fontSize: 14, color: C.background1 });
  txt(s, 'Docente: Zeyla Casaverde Warthon', { x: 0.7, y: 5.85, w: 6.6, h: 0.35, fontSize: 14, color: C.accent5 });
  s.addNotes('Hoy presento el segundo avance de mi proyecto. En el primero hice el diagnóstico de MatLab Perú. En este uso ese diagnóstico para elegir las estrategias con más impacto para la empresa.');
}

// ---------- 2. Punto de partida y objetivo ----------
pres.addSection({ title: 'Punto de partida' });
{
  const s = slide('CLARO', 'Punto de partida');
  kicker(s, '01 / PUNTO DE PARTIDA');
  s.addText('Partimos de un diagnóstico débil y un objetivo claro', { placeholder: 'title' });
  const stats = [
    ['2.34', 'MEFE', 'No aprovecha del todo el entorno'],
    ['2.40', 'MEFI', 'Pesan más las debilidades'],
    ['2.68', 'MPC', 'Segunda detrás de las plataformas'],
    ['S/230 000', 'Ingresos al mes', 'Antes llegaron a unos S/400 000'],
  ];
  stats.forEach(([n, l, d], i) => {
    const x = 0.6 + (i % 2) * 3.55, y = 1.85 + Math.floor(i / 2) * 2.35;
    card(s, x, y, 3.3, 2.1, C.background1, `stat-${i}`);
    txt(s, n, { x: x + 0.3, y: y + 0.25, w: 2.8, h: 0.75, fontSize: i === 3 ? 30 : 40, bold: true, color: C.text1 });
    txt(s, l.toUpperCase(), { x: x + 0.3, y: y + 1.05, w: 2.8, h: 0.3, fontSize: 11, bold: true, charSpacing: 1, color: C.accent2 });
    txt(s, d, { x: x + 0.3, y: y + 1.38, w: 2.8, h: 0.55, fontSize: 13, color: C.accent4 });
  });
  card(s, 8.0, 1.85, 4.7, 4.45, C.text2, 'objetivo');
  txt(s, 'OBJETIVO PLANTEADO', { x: 8.35, y: 2.15, w: 4, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: C.accent1 });
  txt(s, '500 alumnos', { x: 8.35, y: 2.6, w: 4, h: 0.8, fontSize: 40, bold: true, color: C.background1 });
  txt(s, 'en octubre y noviembre de 2028, mejorando el cumplimiento de los profesores.', { x: 8.35, y: 3.45, w: 4.0, h: 1.2, fontSize: 18, color: C.background1 });
  txt(s, 'Son los meses de más demanda, cuando los padres buscan que sus hijos no desaprueben.', { x: 8.35, y: 4.85, w: 4.0, h: 1.0, fontSize: 13, color: C.accent5 });
  s.addNotes('El diagnóstico dio MEFE 2.34 y MEFI 2.40, los dos debajo de 2.5. En la MPC quedamos segundos. Los ingresos bajaron de unos 400 mil a unos 230 mil soles al mes. Por eso el objetivo es volver a unos 500 alumnos en octubre y noviembre de 2028, que son los meses de más demanda, y lograr que los profesores cumplan.');
}

// ---------- 3. FODA cruzado ----------
pres.addSection({ title: 'Generación de estrategias' });
{
  const s = slide('CLARO', 'Generación de estrategias');
  kicker(s, '02 / FODA CRUZADO');
  s.addText('El FODA cruzado dio diez estrategias', { placeholder: 'title' });
  const Q = [
    ['FO · EXPLOTAR', [['E1', 'Nivelación por grado escolar', 'F1, F4, O2'], ['E2', 'Campañas y diagnóstico gratuito con IA', 'F3, F6, O1, O5'], ['E3', 'Familias de otras ciudades', 'F2, F3, O1, O2']]],
    ['DO · BUSCAR', [['E4', 'Orden comercial en el CRM', 'D4, D5, O5'], ['E5', 'Medir y premiar el cumplimiento docente', 'D2, D3, D6, O5'], ['E6', 'Alianzas con colegios y empresas', 'D5, O2, O4']]],
    ['FA · CONFRONTAR', [['E7', 'Destacar profesor fijo, asesor y reporte', 'F1, F4, A5'], ['E8', 'Línea de carrera docente', 'F2, F6, A6']]],
    ['DA · EVITAR', [['E9', 'Profesores suplentes', 'D1, D2, A6'], ['E10', 'Grupos de tres o cuatro alumnos', 'D5, A1, A5']]],
  ];
  Q.forEach(([t, items], i) => {
    const x = 0.6 + (i % 2) * 6.15, y = 1.75 + Math.floor(i / 2) * 2.6;
    card(s, x, y, 5.95, 2.4, i % 3 === 0 ? C.accent3 : C.background1, `cuadrante-${i}`);
    txt(s, t, { x: x + 0.3, y: y + 0.2, w: 5.4, h: 0.3, fontSize: 12, bold: true, charSpacing: 2, color: C.accent2 });
    s.addText(items.flatMap(([c, n, f], k) => [
      { text: `${c}  `, options: { bold: true, color: C.text1 } },
      { text: n, options: { color: C.text1 } },
      { text: `   ${f}`, options: { color: C.accent4, fontSize: 11, breakLine: k < items.length - 1 } },
    ]), { x: x + 0.3, y: y + 0.6, w: 5.45, h: 1.65, fontSize: 15, valign: 'top', margin: 0, paraSpaceAfter: 6, isTextBox: true });
  });
  s.addNotes('Crucé fortalezas, debilidades, oportunidades y amenazas. Salieron diez estrategias, cada una con los factores que la sostienen. Las FO aprovechan lo que hacemos bien, las DO corrigen debilidades, las FA enfrentan amenazas y las DA nos protegen.');
}

// ---------- 4. MIE ----------
{
  const s = slide('OSCURO', 'Generación de estrategias');
  kicker(s, '03 / MATRIZ INTERNA EXTERNA', true);
  s.addText('La MIE ubica a MatLab en la celda V: retener y mantener', { placeholder: 'title' });
  const gx = 1.9, gy = 2.05, cw = 1.55, ch = 1.25;
  const cells = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX'];
  const zone = { I: 'g', II: 'g', IV: 'g', III: 'y', V: 'y', VII: 'y', VI: 'r', VIII: 'r', IX: 'r' };
  cells.forEach((c, i) => {
    const x = gx + (i % 3) * (cw + 0.08), y = gy + Math.floor(i / 3) * (ch + 0.08);
    const on = c === 'V';
    const fill = on ? C.accent1 : zone[c] === 'g' ? C.accent2 : zone[c] === 'y' ? C.accent4 : C.text1;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: ch, fill: { color: fill }, line: { color: fill }, objectName: `celda-${c}` });
    txt(s, on ? [{ text: 'V', options: { bold: true, fontSize: 22, breakLine: true } }, { text: 'MatLab', options: { fontSize: 13, bold: true } }] : c,
      { x, y, w: cw, h: ch, align: 'center', valign: 'middle', fontSize: 18, color: on ? C.text2 : C.background1 });
  });
  ['Fuerte', 'Promedio', 'Débil'].forEach((t, i) => txt(s, t, { x: gx + i * (cw + 0.08), y: gy - 0.4, w: cw, h: 0.3, align: 'center', fontSize: 12, color: C.accent5 }));
  ['Alto', 'Medio', 'Bajo'].forEach((t, i) => txt(s, t, { x: gx - 0.95, y: gy + i * (ch + 0.08) + 0.45, w: 0.85, h: 0.3, align: 'right', fontSize: 12, color: C.accent5 }));
  txt(s, 'MEFI 2.40 en el eje horizontal  ·  MEFE 2.34 en el vertical', { x: 1.0, y: 6.2, w: 5.8, h: 0.3, fontSize: 12, color: C.accent5 });
  txt(s, 'Qué recomienda', { x: 7.6, y: 2.0, w: 5, h: 0.4, fontSize: 20, bold: true, color: C.accent1 });
  txt(s, 'Penetración de mercado y desarrollo de productos. La empresa no debe arriesgarse con negocios nuevos, sino vender más de lo que ya hace bien.', { x: 7.6, y: 2.5, w: 5.0, h: 1.5, fontSize: 16, color: C.background1 });
  txt(s, 'Estrategias del FODA que coinciden', { x: 7.6, y: 4.2, w: 5, h: 0.4, fontSize: 20, bold: true, color: C.accent1 });
  ['E1', 'E2', 'E4', 'E7', 'E10'].forEach((c, i) => {
    s.addShape(pres.shapes.OVAL, { x: 7.6 + i * 0.95, y: 4.8, w: 0.8, h: 0.8, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: `chip-${c}` });
    txt(s, c, { x: 7.6 + i * 0.95, y: 4.8, w: 0.8, h: 0.8, align: 'center', valign: 'middle', fontSize: 14, bold: true, color: C.text2 });
  });
  s.addNotes('Con MEFI 2.40 y MEFE 2.34, MatLab cae en la celda V, que es de retener y mantener. Ahí se recomienda penetración de mercado y desarrollo de productos. Eso coincide con cinco estrategias del FODA: E1, E2, E4, E7 y E10.');
}

// ---------- 5. PEYEA ----------
{
  const s = slide('CLARO', 'Generación de estrategias');
  kicker(s, '04 / MATRIZ PEYEA');
  s.addText('PEYEA: agresiva, pero frenada por el cumplimiento', { placeholder: 'title' });
  card(s, 0.6, 1.75, 4.95, 4.95, C.background1, 'marco-peyea');
  s.addImage({ path: IMG('peyea_ppt.png'), x: 0.75, y: 1.9, w: 4.65, h: 4.65, objectName: 'grafico-peyea' });
  const D = [['+4.40', 'Fuerza financiera', 'Sin deudas, pagos a tiempo y ahorro'], ['−3.33', 'Ventaja competitiva', 'Puntualidad docente con −6'], ['−3.33', 'Estabilidad del entorno', 'La IA cambia rápido el estudio'], ['+3.50', 'Fuerza de la industria', 'Conectividad y demanda de refuerzo']];
  D.forEach(([n, l, d], i) => {
    const x = 5.95 + (i % 2) * 3.45, y = 1.75 + Math.floor(i / 2) * 1.55;
    card(s, x, y, 3.25, 1.4, i === 1 ? C.accent3 : C.background1, `dim-${i}`);
    txt(s, n, { x: x + 0.25, y: y + 0.15, w: 1.5, h: 0.55, fontSize: 26, bold: true, color: C.text1 });
    txt(s, l, { x: x + 0.25, y: y + 0.72, w: 2.85, h: 0.3, fontSize: 13, bold: true, color: C.accent2 });
    txt(s, d, { x: x + 0.25, y: y + 1.0, w: 2.85, h: 0.3, fontSize: 11, color: C.accent4 });
  });
  card(s, 5.95, 4.95, 6.7, 1.75, C.text2, 'resultado-peyea');
  txt(s, 'Vector X = 0.17 · Y = 1.07  →  cuadrante agresivo', { x: 6.2, y: 5.12, w: 6.3, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
  txt(s, 'Estrategias: E1, E2, E3, E4, E5 y E10. Crecer con lo que ya hace bien y corregir la puntualidad, que es lo que más la debilita.', { x: 6.2, y: 5.6, w: 6.3, h: 0.95, fontSize: 14, color: C.background1 });
  s.addNotes('La fuerza financiera es buena: la empresa no tiene deudas, paga a tiempo y tiene ahorro. La industria también tiene demanda. Lo débil es la ventaja competitiva, sobre todo por la puntualidad de los profesores. El vector cae en el cuadrante agresivo, pero muy cerca del eje. De aquí salen seis estrategias.');
}

// ---------- 6. BCG ----------
{
  const s = slide('CLARO', 'Generación de estrategias');
  kicker(s, '05 / MATRIZ BCG');
  s.addText('En la BCG, dos vacas sostienen el 90 % de los ingresos', { placeholder: 'title' });
  card(s, 0.6, 1.75, 5.6, 4.95, C.background1, 'marco-bcg');
  s.addImage({ path: IMG('bcg_ppt.png'), x: 0.8, y: 1.95, w: 5.2, h: 4.5, objectName: 'grafico-bcg' });
  const L = [['Vaca', 'Clases individuales · 60 %', 'Nadie ofrece asesor académico y jefe del área. Estrategias: E1, E2, E5 y E7.'], ['Vaca', 'Programas de verano · 30 %', 'Muchos alumnos y metodología dinámica. Estrategia: E10.'], ['Interrogante', 'Talleres y empresas · 10 %', 'Hay demanda, pero la línea está en pausa. Estrategia: E6.']];
  L.forEach(([q, t, d], i) => {
    const y = 1.75 + i * 1.7;
    card(s, 6.6, y, 6.05, 1.5, i === 2 ? C.accent3 : C.background1, `linea-${i}`);
    txt(s, q.toUpperCase(), { x: 6.9, y: y + 0.18, w: 5.5, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: C.accent2 });
    txt(s, t, { x: 6.9, y: y + 0.48, w: 5.5, h: 0.4, fontSize: 17, bold: true, color: C.text1 });
    txt(s, d, { x: 6.9, y: y + 0.9, w: 5.55, h: 0.5, fontSize: 13, color: C.accent4 });
  });
  s.addNotes('Las clases individuales y los programas de verano son vacas: tienen buena participación, pero ya no crecen. Hay que cuidarlas, porque financian al resto. Talleres y empresas es un interrogante: hay demanda, pero hoy casi no participamos. La participación y el crecimiento son estimaciones mías, porque no hay datos públicos del mercado.');
}

// ---------- 7. MDE ----------
pres.addSection({ title: 'Decisión' });
{
  const s = slide('OSCURO', 'Decisión');
  kicker(s, '06 / MATRIZ DE DECISIÓN', true);
  s.addText('Seis estrategias se repiten tres veces o más', { placeholder: 'title' });
  const rows = [
    ['E1', 'Nivelación por grado escolar', 1, 1, 1, 1], ['E2', 'Campañas y diagnóstico gratuito', 1, 1, 1, 1], ['E10', 'Grupos de tres o cuatro alumnos', 1, 1, 1, 1],
    ['E4', 'Orden comercial en el CRM', 1, 1, 1, 0], ['E5', 'Medir y premiar el cumplimiento docente', 1, 0, 1, 1], ['E7', 'Destacar profesor fijo, asesor y reporte', 1, 1, 0, 1],
    ['E3', 'Familias de otras ciudades', 1, 0, 1, 0], ['E6', 'Alianzas con colegios y empresas', 1, 0, 0, 1], ['E8', 'Línea de carrera docente', 1, 0, 0, 0], ['E9', 'Profesores suplentes', 1, 0, 0, 0],
  ];
  const head = ['Cód.', 'Estrategia', 'FODA', 'MIE', 'PEYEA', 'BCG', 'Total'].map((t) => ({ text: t, options: { bold: true, color: HEX.dk2, fill: { color: HEX.accent1 }, align: 'center' } }));
  const body = rows.map((r, i) => {
    const tot = r.slice(2).reduce((a, b) => a + b, 0);
    const keep = tot >= 3;
    const o = { color: keep ? HEX.lt1 : HEX.accent5, fill: { color: keep ? '0F3256' : HEX.dk2 } };
    return [
      { text: r[0], options: { ...o, bold: true, align: 'center' } },
      { text: r[1], options: { ...o } },
      ...r.slice(2).map((x) => ({ text: x ? '●' : '', options: { ...o, align: 'center', color: HEX.accent1 } })),
      { text: String(tot), options: { ...o, bold: true, align: 'center' } },
    ];
  });
  s.addTable([head, ...body], { x: 0.6, y: 1.7, w: 8.4, colW: [0.75, 3.6, 0.8, 0.7, 0.9, 0.75, 0.9], fontSize: 13, rowH: 0.42, border: { type: 'solid', pt: 0.5, color: HEX.dk2 }, valign: 'middle', objectName: 'tabla-mde' });
  card(s, 9.4, 1.7, 3.3, 2.2, '0F3256', 'retenidas');
  txt(s, '6 retenidas', { x: 9.65, y: 1.9, w: 2.9, h: 0.55, fontSize: 26, bold: true, color: C.accent1 });
  txt(s, 'E1, E2, E10, E4, E5 y E7 pasan a la MCPE.', { x: 9.65, y: 2.55, w: 2.85, h: 1.1, fontSize: 15, color: C.background1 });
  card(s, 9.4, 4.1, 3.3, 2.2, C.text1, 'contingencia');
  txt(s, '4 de contingencia', { x: 9.65, y: 4.3, w: 2.9, h: 0.5, fontSize: 20, bold: true, color: C.background1 });
  txt(s, 'E3, E6, E8 y E9 se usan si las primeras no bastan.', { x: 9.65, y: 4.9, w: 2.85, h: 1.1, fontSize: 15, color: C.accent5 });
  s.addNotes('Junté las estrategias de las cuatro matrices y conté cuántas veces aparece cada una. Me quedé con las que aparecen tres veces o más. Son seis, y cubren las dos partes del objetivo: más alumnos y profesores que cumplen. Las otras cuatro quedan como contingencia.');
}

// ---------- 8. MCPE ----------
{
  const s = slide('CLARO', 'Decisión');
  kicker(s, '07 / MATRIZ CUANTITATIVA MCPE');
  s.addText('La MCPE elige cuatro estrategias', { placeholder: 'title' });
  const data = [['E4 Orden comercial', 4.21], ['E7 Profesor fijo y asesor', 4.73], ['E2 Campañas y diagnóstico', 4.94], ['E1 Nivelación por grado', 5.29], ['E10 Grupos de 3 o 4', 5.40], ['E5 Cumplimiento docente', 5.48]];
  card(s, 0.6, 1.75, 8.1, 4.95, C.background1, 'marco-mcpe');
  s.addChart(pres.charts.BAR, [{ name: 'Total', labels: data.map((d) => d[0]), values: data.map((d) => d[1]) }], {
    x: 0.8, y: 1.9, w: 7.7, h: 4.65, barDir: 'bar', chartColors: [HEX.accent5, HEX.accent5, HEX.accent2, HEX.accent2, HEX.accent2, HEX.accent2],
    valAxisMinVal: 2, valAxisMaxVal: 6, valAxisMajorUnit: 1, showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.00',
    dataLabelFontSize: 13, dataLabelColor: HEX.dk1, dataLabelFontFace: '+mn-lt', catAxisLabelFontSize: 13, catAxisLabelColor: HEX.dk1, catAxisLabelFontFace: '+mn-lt',
    valAxisLabelFontSize: 11, valAxisLabelColor: HEX.accent4, valAxisLabelFontFace: '+mn-lt', valGridLine: { color: 'E1E8ED', size: 0.5 }, catGridLine: { style: 'none' },
    showLegend: false, barGapWidthPct: 60, objectName: 'grafico-mcpe',
  });
  card(s, 9.1, 1.75, 3.6, 2.35, C.text2, 'regla');
  txt(s, 'Cómo se lee', { x: 9.35, y: 1.95, w: 3.1, h: 0.4, fontSize: 18, bold: true, color: C.accent1 });
  txt(s, '23 factores de la MEFE y la MEFI, con sus mismos pesos. El total va de 2 a 8 y desde 5 es aceptable.', { x: 9.35, y: 2.4, w: 3.15, h: 1.55, fontSize: 14, color: C.background1 });
  card(s, 9.1, 4.3, 3.6, 2.4, C.accent3, 'nota-e2');
  txt(s, 'E2 con 4.94', { x: 9.35, y: 4.5, w: 3.1, h: 0.4, fontSize: 18, bold: true, color: C.text1 });
  txt(s, 'Queda apenas debajo de 5. La mantengo porque es la que más ayuda a conseguir inscripciones.', { x: 9.35, y: 4.95, w: 3.15, h: 1.6, fontSize: 14, color: C.text1 });
  s.addNotes('En la MCPE califiqué qué tan atractiva es cada estrategia frente a cada factor de la MEFE y la MEFI. Gana E5, el cumplimiento de los profesores, porque ataca las debilidades que más pesan. Las cuatro primeras pasan a la Matriz de Ética. E7 y E4 quedan como contingencia.');
}

// ---------- 9. Ética ----------
{
  const s = slide('CLARO', 'Decisión');
  kicker(s, '08 / MATRIZ DE ÉTICA');
  s.addText('Las cuatro estrategias pasan la Matriz de Ética', { placeholder: 'title' });
  const E = [
    ['E5', '5.48', 'Medir y premiar el cumplimiento docente', 'Promueve el debido proceso: reglas claras y cada profesor ve sus datos. Es justa porque el bono sigue la misma regla para todos.'],
    ['E10', '5.40', 'Grupos de tres o cuatro alumnos', 'Es justa en la distribución: baja el precio y llega a familias que no pagan una clase individual.'],
    ['E1', '5.29', 'Nivelación por grado escolar', 'Es justa en la distribución: acerca el refuerzo en matemática a más familias.'],
    ['E2', '4.94', 'Campañas y diagnóstico gratuito', 'Neutral en privacidad solo con autorización de los padres y cumpliendo la ley de datos.'],
  ];
  E.forEach(([c, p, t, d], i) => {
    const x = 0.6 + i * 3.075;
    card(s, x, 1.75, 2.9, 4.95, i === 3 ? C.accent3 : C.background1, `etica-${c}`);
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: 2.0, w: 0.85, h: 0.85, fill: { color: C.text2 }, line: { color: C.text2 }, objectName: `codigo-${c}` });
    txt(s, c, { x: x + 0.25, y: 2.0, w: 0.85, h: 0.85, align: 'center', valign: 'middle', fontSize: 16, bold: true, color: C.accent1 });
    txt(s, p, { x: x + 1.3, y: 2.15, w: 1.4, h: 0.55, fontSize: 22, bold: true, color: C.accent2 });
    txt(s, t, { x: x + 0.25, y: 3.1, w: 2.45, h: 0.95, fontSize: 15, bold: true, color: C.text1 });
    txt(s, d, { x: x + 0.25, y: 4.1, w: 2.45, h: 1.75, fontSize: 13, color: C.accent4 });
    txt(s, 'SE ACEPTA', { x: x + 0.25, y: 6.1, w: 2.45, h: 0.35, fontSize: 12, bold: true, charSpacing: 2, color: C.accent2 });
  });
  s.addNotes('La Matriz de Ética revisa derechos, justicia y utilitarismo. Ninguna estrategia viola derechos ni es injusta. E2 tiene una condición: el examen recoge datos de menores, así que solo se aplica con autorización de los padres y cumpliendo la ley de datos, como dice nuestro código de ética.');
}

// ---------- 10. Conclusiones ----------
pres.addSection({ title: 'Cierre' });
{
  const s = slide('OSCURO', 'Cierre');
  kicker(s, '09 / CONCLUSIONES', true);
  s.addText('El primer paso es que los profesores cumplan', { placeholder: 'title' });
  const F = [['1', 'E5', 'Medir y premiar el cumplimiento docente'], ['2', 'E10', 'Grupos de tres o cuatro alumnos'], ['3', 'E1', 'Nivelación por grado escolar'], ['4', 'E2', 'Campañas y diagnóstico gratuito']];
  F.forEach(([n, c, t], i) => {
    const y = 1.75 + i * 0.85;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.65, h: 0.65, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: `paso-${n}` });
    txt(s, n, { x: 0.6, y, w: 0.65, h: 0.65, align: 'center', valign: 'middle', fontSize: 18, bold: true, color: C.text2 });
    txt(s, [{ text: `${c}  `, options: { bold: true, color: C.accent1 } }, { text: t, options: { color: C.background1 } }], { x: 1.45, y: y + 0.12, w: 6.2, h: 0.45, fontSize: 18 });
  });
  card(s, 8.0, 1.75, 4.7, 3.25, '0F3256', 'cierre');
  txt(s, 'Por qué E5 primero', { x: 8.3, y: 1.95, w: 4.1, h: 0.4, fontSize: 18, bold: true, color: C.accent1 });
  txt(s, 'Las demás estrategias solo funcionan si el profesor cumple. Juntas apuntan a volver a unos 500 alumnos en octubre y noviembre de 2028.', { x: 8.3, y: 2.45, w: 4.15, h: 2.4, fontSize: 15, color: C.background1 });
  txt(s, 'REFERENCIAS', { x: 0.6, y: 5.3, w: 4, h: 0.25, fontSize: 10, bold: true, charSpacing: 2, color: C.accent1 });
  txt(s, [
    { text: 'David, F. R. (2017). Conceptos de administración estratégica (15.ª ed.). Pearson Educación.', options: { breakLine: true } },
    { text: 'Forbes Perú. (2026, 20 de mayo). Hogares peruanos con acceso a internet fijo o móvil se elevó a 96 % en el 2025, según Erestel. https://forbes.pe/actualidad/2026-05-20/hogares-peruanos-con-acceso-a-internet-fijo-o-movil-se-elevo-a-96-en-el-2025-segun-erestel/', options: { breakLine: true } },
    { text: 'MatLab Perú. (s. f.). Inicio [Sitio web]. https://matlab.com.pe/', options: { breakLine: true } },
    { text: 'Ministerio de Educación. (2026). ENLA 2025: resumen ejecutivo. https://repositorio.minedu.gob.pe/handle/20.500.12799/12773' },
  ], { x: 0.6, y: 5.6, w: 12.1, h: 1.2, fontSize: 10, color: C.accent5, paraSpaceAfter: 2 });
  s.addNotes('Las cuatro estrategias finales son E5, E10, E1 y E2. Recomiendo empezar por E5, porque si el profesor no cumple, ninguna otra estrategia funciona. Las demás quedan como contingencia, sobre todo el orden del área comercial y el grupo de profesores suplentes. Gracias.');
}

pres.writeFile({ fileName: OUT }).then(async () => { await applyTheme(OUT, THEME); console.log('ok', OUT); });
