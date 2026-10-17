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

// ---------- 2. La empresa ----------
pres.addSection({ title: 'Diagnóstico' });
{
  const s = slide('CLARO', 'Diagnóstico');
  kicker(s, '01 / LA EMPRESA');
  s.addText('MatLab Perú enseña matemáticas 1 a 1 y de forma virtual', { placeholder: 'title' });
  const stats = [['2012', 'Año de fundación'], ['2020', 'Todas las clases pasan a ser virtuales'], ['40', 'Personas, con 27 profesores'], ['S/55 a S/65', 'Por clase individual de 50 minutos']];
  stats.forEach(([n, l], i) => {
    const x = 0.6 + i * 3.075;
    card(s, x, 1.75, 2.9, 1.75, C.background1, `dato-${i}`);
    txt(s, n, { x: x + 0.25, y: 1.95, w: 2.5, h: 0.65, fontSize: i === 3 ? 26 : 34, bold: true, color: C.text1 });
    txt(s, l, { x: x + 0.25, y: 2.7, w: 2.5, h: 0.6, fontSize: 13, color: C.accent4 });
  });
  const B = [
    ['MISIÓN', 'Enseñamos matemáticas y educación financiera de forma personalizada y virtual, al ritmo de cada persona, para que entienda los números y los use en su vida.'],
    ['VISIÓN', 'Ser la academia virtual más reconocida del Perú en matemáticas y educación financiera, y que cada alumno pierda el miedo a los números.'],
    ['VALORES', 'Personalización, cercanía y confianza, comprensión, compromiso y puntualidad, innovación e integridad.'],
  ];
  B.forEach(([t, d], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 3.75, 3.9, 2.95, i === 1 ? C.accent3 : C.background1, `identidad-${i}`);
    txt(s, t, { x: x + 0.3, y: 3.98, w: 3.3, h: 0.3, fontSize: 12, bold: true, charSpacing: 2, color: C.accent2 });
    txt(s, d, { x: x + 0.3, y: 4.4, w: 3.35, h: 2.1, fontSize: 15, color: C.text1 });
  });
  s.addNotes('MatLab Perú es una academia fundada en 2012 por Marco Loret de Mola. Desde 2020 todas las clases son virtuales. Somos 40 personas con 27 profesores. Como la empresa no tenía misión, visión ni valores escritos, los propuse yo en el primer avance.');
}

// ---------- 3. Diagnóstico externo ----------
{
  const s = slide('OSCURO', 'Diagnóstico');
  kicker(s, '02 / DIAGNÓSTICO EXTERNO', true);
  s.addText('El entorno tiene demanda, pero mucha competencia', { placeholder: 'title' });
  const P = [
    ['POLÍTICO', 'Cambios de gobierno y nuevas normas', 'A3'], ['ECONÓMICO', 'Economía de las familias y alza de precios', 'O3 · A1'],
    ['SOCIOCULTURAL', 'Solo 12.8 % logra el nivel en matemática', 'O2 · O4'], ['TECNOLÓGICO', '96 % de hogares con internet, y la IA', 'O1 · O5 · A5'],
    ['ECOLÓGICO', 'El Niño puede cortar luz o internet', 'A2'], ['LEGAL', 'Protección de datos de menores', 'A4'],
  ];
  P.forEach(([t, d, c], i) => {
    const x = 0.6 + (i % 3) * 2.62, y = 1.75 + Math.floor(i / 3) * 2.05;
    card(s, x, y, 2.47, 1.9, '0F3256', `pestel-${i}`);
    txt(s, t, { x: x + 0.2, y: y + 0.18, w: 2.1, h: 0.3, fontSize: 10, bold: true, charSpacing: 1, color: C.accent1 });
    txt(s, d, { x: x + 0.2, y: y + 0.52, w: 2.1, h: 0.9, fontSize: 14, color: C.background1 });
    txt(s, c, { x: x + 0.2, y: y + 1.5, w: 2.1, h: 0.25, fontSize: 11, bold: true, color: C.accent5 });
  });
  txt(s, 'PESTEL', { x: 0.6, y: 5.9, w: 3, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: C.accent5 });
  card(s, 8.6, 1.75, 4.1, 1.45, C.accent1, 'mefe');
  txt(s, '2.34', { x: 8.85, y: 1.9, w: 1.6, h: 0.75, fontSize: 36, bold: true, color: C.text2 });
  txt(s, 'MEFE, debajo de 2.5. No aprovecha del todo el entorno.', { x: 10.35, y: 1.95, w: 2.2, h: 1.1, fontSize: 13, color: C.text2 });
  card(s, 8.6, 3.35, 4.1, 0.95, '0F3256', 'porter');
  txt(s, [{ text: 'Porter  ', options: { bold: true, color: C.accent1 } }, { text: 'las cinco fuerzas son altas', options: { color: C.background1 } }], { x: 8.85, y: 3.6, w: 3.7, h: 0.45, fontSize: 15 });
  card(s, 8.6, 4.45, 4.1, 2.25, '0F3256', 'mpc');
  txt(s, 'MPC', { x: 8.85, y: 4.6, w: 3, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: C.accent1 });
  [['Plataformas de tutorías', '2.84'], ['MatLab Perú', '2.68'], ['Academias virtuales', '2.62'], ['Tutores independientes', '2.32']].forEach(([n, v], i) => {
    const on = i === 1;
    txt(s, n, { x: 8.85, y: 4.98 + i * 0.4, w: 2.8, h: 0.35, fontSize: 14, bold: on, color: on ? C.accent1 : C.background1 });
    txt(s, v, { x: 11.7, y: 4.98 + i * 0.4, w: 0.8, h: 0.35, fontSize: 14, bold: on, align: 'right', color: on ? C.accent1 : C.background1 });
  });
  s.addNotes('En el entorno hay oportunidades claras: casi todos los hogares tienen internet y pocos alumnos logran el nivel esperado en matemática. Pero las cinco fuerzas de Porter son altas. La MEFE dio 2.34 y en la MPC quedamos segundos, detrás de las plataformas de tutorías.');
}

// ---------- 4. Diagnóstico interno ----------
{
  const s = slide('CLARO', 'Diagnóstico');
  kicker(s, '03 / DIAGNÓSTICO INTERNO');
  s.addText('Adentro, el problema es el cumplimiento docente', { placeholder: 'title' });
  card(s, 0.6, 1.75, 3.6, 1.6, C.text2, 'mefi');
  txt(s, '2.40', { x: 0.85, y: 1.9, w: 1.6, h: 0.75, fontSize: 36, bold: true, color: C.accent1 });
  txt(s, 'MEFI, debajo de 2.5. Pesan más las debilidades.', { x: 2.35, y: 1.95, w: 1.7, h: 1.25, fontSize: 13, color: C.background1 });
  card(s, 0.6, 3.5, 3.6, 3.2, C.background1, 'amofhit');
  txt(s, 'AMOFHIT Y CADENA DE VALOR', { x: 0.85, y: 3.7, w: 3.2, h: 0.3, fontSize: 11, bold: true, charSpacing: 1, color: C.accent2 });
  txt(s, [
    { text: 'Crea valor en las clases y en el seguimiento a las familias.', options: { breakLine: true } },
    { text: 'Fallan recursos humanos y marketing y ventas.', options: { breakLine: true } },
    { text: 'Algunos profesores faltan, llegan tarde o no prenden la cámara.' },
  ], { x: 0.85, y: 4.1, w: 3.15, h: 2.5, fontSize: 14, color: C.text1, paraSpaceAfter: 8 });
  const Q = [
    ['FORTALEZAS', ['F1 Enseñanza 1 a 1', 'F2 Trayectoria y marca', 'F3 Experiencia virtual', 'F4 Seguimiento a padres', 'F5 Oferta variada', 'F6 Innovación'], C.background1],
    ['DEBILIDADES', ['D1 Dependencia docente', 'D2 Faltas y tardanzas', 'D3 Malas prácticas', 'D4 Desorden comercial', 'D5 Pocas inscripciones', 'D6 Datos e IA en proceso'], C.accent3],
    ['OPORTUNIDADES', ['O1 Conectividad 96 %', 'O2 Bajo logro en matemática', 'O3 Mejora económica', 'O4 Demanda de empresas', 'O5 Tecnología e IA'], C.accent3],
    ['AMENAZAS', ['A1 Alza de precios', 'A2 Clima', 'A3 Cambios de gobierno', 'A4 Datos de menores', 'A5 Bajo costo e IA gratis', 'A6 Disputa por docentes'], C.background1],
  ];
  Q.forEach(([t, items, fill], i) => {
    const x = 4.4 + (i % 2) * 4.2, y = 1.75 + Math.floor(i / 2) * 2.55;
    card(s, x, y, 4.05, 2.4, fill, `foda-${i}`);
    txt(s, t, { x: x + 0.25, y: y + 0.18, w: 3.5, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: C.accent2 });
    const half = Math.ceil(items.length / 2);
    [items.slice(0, half), items.slice(half)].forEach((col, k) => {
      txt(s, col.map((it, j) => ({ text: it, options: { breakLine: j < col.length - 1 } })), { x: x + 0.25 + k * 1.95, y: y + 0.6, w: 1.85, h: 1.65, fontSize: 12, color: C.text1, paraSpaceAfter: 5 });
    });
  });
  s.addNotes('Por dentro, la MEFI dio 2.40. La empresa crea valor en sus clases y en el seguimiento a los padres, pero depende mucho de los profesores y algunos no cumplen. El área comercial también está desordenada. Con todo esto armé el FODA, que es la base de las estrategias.');
}

// ---------- 5. Objetivo y FODA cruzado ----------
pres.addSection({ title: 'Plan estratégico' });
{
  const s = slide('CLARO', 'Plan estratégico');
  kicker(s, '04 / OBJETIVO Y FODA CRUZADO');
  s.addText('Hacia 500 alumnos: el FODA cruzado dio diez estrategias', { placeholder: 'title' });
  card(s, 0.6, 1.65, 12.1, 0.75, C.text2, 'objetivo');
  txt(s, [{ text: 'OBJETIVO   ', options: { bold: true, color: C.accent1, fontSize: 12, charSpacing: 2 } }, { text: 'Volver a unos 500 alumnos en octubre y noviembre de 2028, mejorando el cumplimiento de los profesores.', options: { color: C.background1 } }], { x: 0.9, y: 1.65, w: 11.6, h: 0.75, fontSize: 16, valign: 'middle' });
  const Q = [
    ['FO · EXPLOTAR', [['E1', 'Nivelación por grado escolar', 'F1, F4, O2'], ['E2', 'Campañas y diagnóstico gratuito con IA', 'F3, F6, O1, O5'], ['E3', 'Familias de otras ciudades', 'F2, F3, O1, O2']]],
    ['DO · BUSCAR', [['E4', 'Orden comercial en el CRM', 'D4, D5, O5'], ['E5', 'Medir y premiar el cumplimiento docente', 'D2, D3, D6, O5'], ['E6', 'Alianzas con colegios y empresas', 'D5, O2, O4']]],
    ['FA · CONFRONTAR', [['E7', 'Destacar profesor fijo, asesor y reporte', 'F1, F4, A5'], ['E8', 'Línea de carrera docente', 'F2, F6, A6']]],
    ['DA · EVITAR', [['E9', 'Profesores suplentes', 'D1, D2, A6'], ['E10', 'Grupos de tres o cuatro alumnos', 'D5, A1, A5']]],
  ];
  Q.forEach(([t, items], i) => {
    const x = 0.6 + (i % 2) * 6.15, y = 2.6 + Math.floor(i / 2) * 2.1;
    card(s, x, y, 5.95, 1.95, i % 3 === 0 ? C.accent3 : C.background1, `cuadrante-${i}`);
    txt(s, t, { x: x + 0.3, y: y + 0.17, w: 5.4, h: 0.3, fontSize: 12, bold: true, charSpacing: 2, color: C.accent2 });
    s.addText(items.flatMap(([c, n, f], k) => [
      { text: `${c}  `, options: { bold: true, color: C.text1 } },
      { text: n, options: { color: C.text1 } },
      { text: `   ${f}`, options: { color: C.accent4, fontSize: 11, breakLine: k < items.length - 1 } },
    ]), { x: x + 0.3, y: y + 0.55, w: 5.45, h: 1.3, fontSize: 15, valign: 'top', margin: 0, paraSpaceAfter: 4, isTextBox: true });
  });
  s.addNotes('Con el diagnóstico planteé el objetivo: volver a unos 500 alumnos en octubre y noviembre de 2028, que son los meses de más demanda, y que los profesores cumplan. Luego crucé los factores del FODA y salieron diez estrategias, cada una con los factores que la sostienen.');
}

// ---------- 6. MIE y PEYEA ----------
{
  const s = slide('OSCURO', 'Plan estratégico');
  kicker(s, '05 / MIE Y PEYEA', true);
  s.addText('MIE y PEYEA: crecer en lo que ya hace bien', { placeholder: 'title' });
  txt(s, 'MIE · CELDA V, RETENER Y MANTENER', { x: 0.6, y: 1.7, w: 5.6, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: C.accent1 });
  const gx = 1.25, gy = 2.4, cw = 1.2, ch = 0.9;
  const zone = { I: 'g', II: 'g', IV: 'g', III: 'y', V: 'y', VII: 'y', VI: 'r', VIII: 'r', IX: 'r' };
  ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX'].forEach((c, i) => {
    const x = gx + (i % 3) * (cw + 0.06), y = gy + Math.floor(i / 3) * (ch + 0.06);
    const on = c === 'V';
    const fill = on ? C.accent1 : zone[c] === 'g' ? C.accent2 : zone[c] === 'y' ? C.accent4 : '0F3256';
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: ch, fill: { color: fill }, line: { color: fill }, objectName: `celda-${c}` });
    txt(s, on ? [{ text: 'V', options: { bold: true, fontSize: 18, breakLine: true } }, { text: 'MatLab', options: { fontSize: 11, bold: true } }] : c,
      { x, y, w: cw, h: ch, align: 'center', valign: 'middle', fontSize: 15, color: on ? C.text2 : C.background1 });
  });
  ['Fuerte', 'Promedio', 'Débil'].forEach((t, i) => txt(s, t, { x: gx + i * (cw + 0.06), y: gy - 0.32, w: cw, h: 0.25, align: 'center', fontSize: 11, color: C.accent5 }));
  ['Alto', 'Medio', 'Bajo'].forEach((t, i) => txt(s, t, { x: 0.6, y: gy + i * (ch + 0.06) + 0.32, w: 0.55, h: 0.25, align: 'right', fontSize: 11, color: C.accent5 }));
  txt(s, 'MEFI 2.40 · MEFE 2.34. Recomienda penetración de mercado y desarrollo de productos: E1, E2, E4, E7 y E10.', { x: 0.6, y: 5.4, w: 5.4, h: 1.0, fontSize: 14, color: C.background1 });
  txt(s, 'PEYEA · CUADRANTE AGRESIVO', { x: 6.6, y: 1.7, w: 6, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: C.accent1 });
  card(s, 6.6, 2.1, 3.55, 3.55, C.background1, 'marco-peyea');
  s.addImage({ path: IMG('peyea_ppt.png'), x: 6.7, y: 2.2, w: 3.35, h: 3.35, objectName: 'grafico-peyea' });
  [['FF', '+4.40'], ['VC', '−3.33'], ['EE', '−3.33'], ['FI', '+3.50']].forEach(([l, v], i) => {
    const y = 2.1 + i * 0.6;
    txt(s, l, { x: 10.45, y, w: 0.6, h: 0.45, fontSize: 14, bold: true, color: C.accent1, valign: 'middle' });
    txt(s, v, { x: 11.05, y, w: 1.6, h: 0.45, fontSize: 22, bold: true, color: C.background1, valign: 'middle' });
  });
  txt(s, 'Vector X 0.17 · Y 1.07', { x: 10.45, y: 4.6, w: 2.3, h: 0.9, fontSize: 13, color: C.accent5 });
  txt(s, 'Sin deudas y con ahorro, pero la puntualidad docente, con −6, frena la ventaja competitiva. Estrategias: E1, E2, E3, E4, E5 y E10.', { x: 6.6, y: 5.8, w: 6.1, h: 0.95, fontSize: 14, color: C.background1 });
  s.addNotes('La MIE nos ubica en la celda V, de retener y mantener: conviene penetración de mercado y desarrollo de productos. La PEYEA da cuadrante agresivo, porque la empresa no tiene deudas y tiene ahorro, pero el vector está muy cerca del eje porque la puntualidad de los profesores debilita la ventaja competitiva.');
}

// ---------- 7. BCG ----------
{
  const s = slide('CLARO', 'Plan estratégico');
  kicker(s, '06 / MATRIZ BCG');
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

// ---------- 8. MDE ----------
pres.addSection({ title: 'Decisión' });
{
  const s = slide('OSCURO', 'Decisión');
  kicker(s, '07 / MATRIZ DE DECISIÓN', true);
  s.addText('Seis estrategias se repiten tres veces o más', { placeholder: 'title' });
  const rows = [
    ['E1', 'Nivelación por grado escolar', 1, 1, 1, 1], ['E2', 'Campañas y diagnóstico gratuito', 1, 1, 1, 1], ['E10', 'Grupos de tres o cuatro alumnos', 1, 1, 1, 1],
    ['E4', 'Orden comercial en el CRM', 1, 1, 1, 0], ['E5', 'Medir y premiar el cumplimiento docente', 1, 0, 1, 1], ['E7', 'Destacar profesor fijo, asesor y reporte', 1, 1, 0, 1],
    ['E3', 'Familias de otras ciudades', 1, 0, 1, 0], ['E6', 'Alianzas con colegios y empresas', 1, 0, 0, 1], ['E8', 'Línea de carrera docente', 1, 0, 0, 0], ['E9', 'Profesores suplentes', 1, 0, 0, 0],
  ];
  const head = ['Cód.', 'Estrategia', 'FODA', 'MIE', 'PEYEA', 'BCG', 'Total'].map((t) => ({ text: t, options: { bold: true, color: HEX.dk2, fill: { color: HEX.accent1 }, align: 'center' } }));
  const body = rows.map((r) => {
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

// ---------- 9. MCPE y Ética ----------
{
  const s = slide('CLARO', 'Decisión');
  kicker(s, '08 / MCPE Y MATRIZ DE ÉTICA');
  s.addText('MCPE y Ética: quedan cuatro estrategias', { placeholder: 'title' });
  const data = [['E4 Orden comercial', 4.21], ['E7 Profesor fijo y asesor', 4.73], ['E2 Campañas y diagnóstico', 4.94], ['E1 Nivelación por grado', 5.29], ['E10 Grupos de 3 o 4', 5.40], ['E5 Cumplimiento docente', 5.48]];
  card(s, 0.6, 1.75, 6.6, 4.95, C.background1, 'marco-mcpe');
  txt(s, 'MCPE · TOTAL DE 2 A 8, DESDE 5 ES ACEPTABLE', { x: 0.85, y: 1.92, w: 6.1, h: 0.3, fontSize: 10, bold: true, charSpacing: 1, color: C.accent2 });
  s.addChart(pres.charts.BAR, [{ name: 'Total', labels: data.map((d) => d[0]), values: data.map((d) => d[1]) }], {
    x: 0.75, y: 2.25, w: 6.3, h: 4.35, barDir: 'bar', chartColors: [HEX.accent5, HEX.accent5, HEX.accent2, HEX.accent2, HEX.accent2, HEX.accent2],
    valAxisMinVal: 2, valAxisMaxVal: 6, valAxisMajorUnit: 1, showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.00',
    dataLabelFontSize: 12, dataLabelColor: HEX.dk1, dataLabelFontFace: '+mn-lt', catAxisLabelFontSize: 12, catAxisLabelColor: HEX.dk1, catAxisLabelFontFace: '+mn-lt',
    valAxisLabelFontSize: 10, valAxisLabelColor: HEX.accent4, valAxisLabelFontFace: '+mn-lt', valGridLine: { color: 'E1E8ED', size: 0.5 }, catGridLine: { style: 'none' },
    showLegend: false, barGapWidthPct: 60, objectName: 'grafico-mcpe',
  });
  txt(s, 'MATRIZ DE ÉTICA · LAS CUATRO SE ACEPTAN', { x: 7.6, y: 1.75, w: 5.1, h: 0.3, fontSize: 10, bold: true, charSpacing: 1, color: C.accent2 });
  const E = [
    ['E5', 'Promueve el debido proceso y es justa: misma regla de bono para todos.'],
    ['E10', 'Justa en la distribución: llega a familias que no pagan una clase individual.'],
    ['E1', 'Justa en la distribución: acerca el refuerzo a más familias.'],
    ['E2', 'Neutral en privacidad solo con autorización de los padres.'],
  ];
  E.forEach(([c, d], i) => {
    const y = 2.15 + i * 1.15;
    card(s, 7.6, y, 5.1, 1.0, i === 3 ? C.accent3 : C.background1, `etica-${c}`);
    s.addShape(pres.shapes.OVAL, { x: 7.8, y: y + 0.17, w: 0.66, h: 0.66, fill: { color: C.text2 }, line: { color: C.text2 }, objectName: `codigo-${c}` });
    txt(s, c, { x: 7.8, y: y + 0.17, w: 0.66, h: 0.66, align: 'center', valign: 'middle', fontSize: 13, bold: true, color: C.accent1 });
    txt(s, d, { x: 8.65, y: y + 0.12, w: 3.9, h: 0.78, fontSize: 13, color: C.text1, valign: 'middle' });
  });
  s.addNotes('En la MCPE califiqué cada estrategia frente a los 23 factores de la MEFE y la MEFI. Gana E5, el cumplimiento de los profesores. Las cuatro primeras pasan a la Matriz de Ética y ninguna viola derechos ni es injusta. E2 tiene una condición: el examen recoge datos de menores, así que solo va con autorización de los padres.');
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
