// Presentación del APF3, máximo 10 láminas, con el estilo de los avances anteriores.
const path = require('path');
const fs = require('fs');
const pptxgen = require('pptxgenjs');
const { applyTheme } = require(process.env.PPTX_SKILL + '/scripts/apply_theme.js');

const DIR = path.join(__dirname, '..');
const IMG = (f) => path.join(DIR, 'figuras', f);
const OUT = path.join(DIR, 'APF3_Presentacion_MatLab_Peru.pptx');

const THEME = {
  name: 'MatLab APF3',
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
pres.title = 'Plan estratégico e implementación de MatLab Perú';
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
const circ = (s, x, y, d, t, fill, color, size, name) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill }, objectName: name });
  txt(s, t, { x, y, w: d, h: d, align: 'center', valign: 'middle', fontSize: size, bold: true, color });
};

// ---------- 1. Portada ----------
pres.addSection({ title: 'Inicio' });
{
  const s = slide('OSCURO', 'Inicio');
  s.addImage({ path: IMG('portada.jpg'), x: 7.6, y: 0, w: 5.733, h: 7.5, sizing: { type: 'cover', w: 5.733, h: 7.5 }, objectName: 'foto' });
  txt(s, 'PLANEAMIENTO ESTRATÉGICO  ·  SECCIÓN 20726', { x: 0.7, y: 1.0, w: 6.6, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: C.accent1 });
  txt(s, 'Avance de Proyecto Final 3', { x: 0.7, y: 1.55, w: 6.6, h: 0.45, fontSize: 18, color: C.background1 });
  txt(s, 'Plan estratégico e implementación de MatLab Perú', { x: 0.7, y: 2.05, w: 6.6, h: 2.0, fontSize: 36, bold: true, color: C.background1 });
  s.addImage({ path: IMG('logo_utp_banda.png'), x: 0.7, y: 4.35, w: LOGO_W, h: LOGO_W * LOGO_R, objectName: 'logo' });
  txt(s, 'Pierre Gallegos Mallqui  ·  Código U23211135', { x: 0.7, y: 5.45, w: 6.6, h: 0.35, fontSize: 14, color: C.background1 });
  txt(s, 'Docente: Zeyla Casaverde Warthon', { x: 0.7, y: 5.85, w: 6.6, h: 0.35, fontSize: 14, color: C.accent5 });
  s.addNotes('Presento el tercer avance de mi proyecto sobre MatLab Perú. Primero repaso rápido el diagnóstico y las estrategias elegidas, y luego explico cómo se van a implementar: objetivos, matrices, Balanced Scorecard y estructura.');
}

// ---------- 2. Diagnóstico ----------
pres.addSection({ title: 'Lo ya analizado' });
{
  const s = slide('CLARO', 'Lo ya analizado');
  kicker(s, '01 / DIAGNÓSTICO');
  s.addText('Partimos de un diagnóstico débil por dentro y por fuera', { placeholder: 'title' });
  const stats = [['2.34', 'MEFE', 'No aprovecha del todo el entorno'], ['2.40', 'MEFI', 'Pesan más las debilidades'], ['2.68', 'MPC', 'Segunda detrás de las plataformas'], ['5 altas', 'PORTER', 'Las cinco fuerzas son altas']];
  stats.forEach(([n, l, d], i) => {
    const x = 0.6 + i * 3.075;
    card(s, x, 1.75, 2.9, 2.0, C.background1, `stat-${i}`);
    txt(s, n, { x: x + 0.25, y: 1.95, w: 2.5, h: 0.75, fontSize: 36, bold: true, color: C.text1 });
    txt(s, l, { x: x + 0.25, y: 2.75, w: 2.5, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: C.accent2 });
    txt(s, d, { x: x + 0.25, y: 3.08, w: 2.5, h: 0.55, fontSize: 13, color: C.accent4 });
  });
  card(s, 0.6, 4.0, 6.0, 2.7, C.background1, 'fortalezas');
  txt(s, 'LO QUE HACE BIEN', { x: 0.9, y: 4.2, w: 5.4, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: C.accent2 });
  txt(s, [{ text: 'Enseñanza personalizada 1 a 1', options: { breakLine: true } }, { text: 'Experiencia en clases 100 % virtuales desde 2020', options: { breakLine: true } }, { text: 'Seguimiento y reportes a los padres' }], { x: 0.9, y: 4.6, w: 5.5, h: 1.9, fontSize: 16, color: C.text1, paraSpaceAfter: 8 });
  card(s, 6.75, 4.0, 5.95, 2.7, C.text2, 'debilidades');
  txt(s, 'LO QUE DEBE CORREGIR', { x: 7.05, y: 4.2, w: 5.4, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: C.accent1 });
  txt(s, [{ text: 'Faltas, tardanzas y cámara apagada', options: { breakLine: true } }, { text: 'Desorden comercial y pocas inscripciones', options: { breakLine: true } }, { text: 'Ingresos de S/400 000 a S/230 000 al mes' }], { x: 7.05, y: 4.6, w: 5.45, h: 1.9, fontSize: 16, color: C.background1, paraSpaceAfter: 8 });
  s.addNotes('En el primer avance la MEFE dio 2.34 y la MEFI 2.40, los dos debajo de 2.5. Quedamos segundos en la MPC y las cinco fuerzas son altas. Hacemos bien la enseñanza personalizada, pero fallamos en el cumplimiento de los profesores y en la captación.');
}

// ---------- 3. Estrategias elegidas ----------
{
  const s = slide('OSCURO', 'Lo ya analizado');
  kicker(s, '02 / ESTRATEGIAS ELEGIDAS', true);
  s.addText('De diez estrategias quedaron cuatro', { placeholder: 'title' });
  const steps = [['10', 'FODA cruzado'], ['6', 'Matriz de Decisión'], ['4', 'MCPE y Ética']];
  steps.forEach(([n, l], i) => {
    const x = 0.6 + i * 1.75;
    circ(s, x + 0.15, 1.9, 1.2, n, i === 2 ? C.accent1 : '0F3256', i === 2 ? C.text2 : C.background1, 30, `paso-${i}`);
    txt(s, l, { x, y: 3.2, w: 1.5, h: 0.6, fontSize: 13, align: 'center', color: C.accent5 });
  });
  txt(s, 'MIE celda V, retener y mantener · PEYEA agresiva · BCG con dos vacas y un interrogante', { x: 0.6, y: 4.2, w: 5.1, h: 1.2, fontSize: 15, color: C.background1 });
  txt(s, 'Objetivo: 500 alumnos en octubre y noviembre de 2028', { x: 0.6, y: 5.55, w: 5.1, h: 0.9, fontSize: 16, bold: true, color: C.accent1 });
  const E = [['E5', '5.48', 'Medir y premiar el cumplimiento docente'], ['E10', '5.40', 'Clases en grupos de tres o cuatro alumnos'], ['E1', '5.29', 'Programas de nivelación por grado escolar'], ['E2', '4.94', 'Campañas en redes con diagnóstico gratuito']];
  E.forEach(([c, p, t], i) => {
    const y = 1.75 + i * 1.22;
    card(s, 6.2, y, 6.5, 1.07, '0F3256', `estrategia-${c}`);
    circ(s, 6.4, y + 0.18, 0.7, c, C.accent1, C.text2, 13, `cod-${c}`);
    txt(s, t, { x: 7.35, y: y + 0.18, w: 4.2, h: 0.7, fontSize: 16, color: C.background1, valign: 'middle' });
    txt(s, p, { x: 11.55, y: y + 0.18, w: 0.95, h: 0.7, fontSize: 18, bold: true, color: C.accent1, valign: 'middle', align: 'right' });
  });
  s.addNotes('En el segundo avance el FODA cruzado dio diez estrategias. La Matriz de Decisión retuvo seis y la MCPE eligió cuatro, que pasaron la Matriz de Ética. Estas cuatro son las que ahora implementamos.');
}

// ---------- 4. Objetivos de largo plazo ----------
pres.addSection({ title: 'Implementación' });
{
  const s = slide('CLARO', 'Implementación');
  kicker(s, '03 / OBJETIVOS DE LARGO PLAZO');
  s.addText('Tres objetivos para 2028, calculados con una tasa anual', { placeholder: 'title' });
  const O = [
    ['OLP1', 'Alumnos', '288 → 500', 'en octubre y noviembre', '√(500 / 288) − 1 = 31.8 % al año', '2027: 380 alumnos'],
    ['OLP2', 'Ingresos', '230 → 300', 'miles de soles, promedio al mes', '√(300 000 / 230 000) − 1 = 14.2 % al año', '2027: S/262 700'],
    ['OLP3', 'Cumplimiento', '25 → 5', 'clases con problemas de cada 100', '√(5 / 25) = 0.447 por año', '2027: 11 de cada 100'],
  ];
  O.forEach(([c, t, n, d, f, m], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.75, 3.9, 4.95, i === 2 ? C.accent3 : C.background1, `olp-${i}`);
    txt(s, `${c} · ${t.toUpperCase()}`, { x: x + 0.3, y: 1.98, w: 3.3, h: 0.3, fontSize: 12, bold: true, charSpacing: 2, color: C.accent2 });
    txt(s, n, { x: x + 0.3, y: 2.45, w: 3.4, h: 0.75, fontSize: 34, bold: true, color: C.text1 });
    txt(s, d, { x: x + 0.3, y: 3.2, w: 3.3, h: 0.6, fontSize: 14, color: C.accent4 });
    card(s, x + 0.3, 4.05, 3.3, 1.15, C.text2, `calc-${i}`);
    txt(s, f, { x: x + 0.45, y: 4.12, w: 3.05, h: 1.0, fontSize: 14, color: C.background1, valign: 'middle' });
    txt(s, m, { x: x + 0.3, y: 5.55, w: 3.3, h: 0.5, fontSize: 18, bold: true, color: C.accent2 });
  });
  s.addNotes('Planteé tres objetivos para 2028. Para cada uno calculé una tasa fija por año, así sé cuánto debo avanzar en 2027. Por ejemplo, para pasar de 288 a 500 alumnos en dos años necesitamos crecer 31.8 por ciento cada año, o sea llegar a 380 en 2027.');
}

// ---------- 5. Objetivos de corto plazo ----------
{
  const s = slide('CLARO', 'Implementación');
  kicker(s, '04 / OBJETIVOS DE CORTO PLAZO');
  s.addText('Nueve objetivos de corto plazo, tres por cada OLP', { placeholder: 'title' });
  const R = [
    ['OLP1', [['OCP1.1', '380 alumnos en 2027', 'E1 · E2 · E10'], ['OCP1.2', 'Compras de 3 a 4 y 5.2 por cada 100 interesados', 'E2'], ['OCP1.3', '50 en nivelación y 40 en grupos en 2027', 'E1 · E10']]],
    ['OLP2', [['OCP2.1', 'S/262 700 al mes en 2027', 'E1 · E2 · E10'], ['OCP2.2', 'Nuevos programas: 10 % y luego 20 % de ingresos', 'E1 · E10'], ['OCP2.3', 'S/25 por alumno en grupo: S/75 a S/100 por clase', 'E10']]],
    ['OLP3', [['OCP3.1', '11 de cada 100 clases con problemas en 2027', 'E5'], ['OCP3.2', 'Los 27 profesores con tablero desde marzo 2027', 'E5'], ['OCP3.3', 'Profesores que se van: de 10 a 6 y a 4', 'E5']]],
  ];
  R.forEach(([o, items], i) => {
    const y = 1.7 + i * 1.7;
    card(s, 0.6, y, 1.4, 1.55, C.text2, `fila-${o}`);
    txt(s, o, { x: 0.6, y, w: 1.4, h: 1.55, align: 'center', valign: 'middle', fontSize: 18, bold: true, color: C.accent1 });
    items.forEach(([c, t, e], k) => {
      const x = 2.15 + k * 3.55;
      card(s, x, y, 3.4, 1.55, C.background1, `${c}`);
      txt(s, c, { x: x + 0.2, y: y + 0.15, w: 1.5, h: 0.3, fontSize: 11, bold: true, color: C.accent2 });
      txt(s, e, { x: x + 1.7, y: y + 0.15, w: 1.5, h: 0.3, fontSize: 11, bold: true, color: C.accent4, align: 'right' });
      txt(s, t, { x: x + 0.2, y: y + 0.5, w: 3.0, h: 0.95, fontSize: 14, color: C.text1 });
    });
  });
  s.addNotes('Cada objetivo de largo plazo tiene tres de corto plazo. Todos salen de un cálculo. Por ejemplo, el precio de 25 soles por alumno en grupo hace que una clase de tres o cuatro alumnos deje entre 75 y 100 soles, más que una clase individual. Antes probamos 35 soles y no funcionó.');
}

// ---------- 6. MEOLP y MEPCS ----------
{
  const s = slide('OSCURO', 'Implementación');
  kicker(s, '05 / MEOLP Y MEPCS', true);
  s.addText('MEOLP y MEPCS: E5 es la más difícil de copiar', { placeholder: 'title' });
  const head = ['MEOLP', 'OLP1', 'OLP2', 'OLP3'].map((t) => ({ text: t, options: { bold: true, color: HEX.dk2, fill: { color: HEX.accent1 }, align: 'center' } }));
  const rows = [['E5', 1, 0, 1], ['E10', 1, 1, 0], ['E1', 1, 1, 0], ['E2', 1, 1, 0]].map((r) => [
    { text: r[0], options: { bold: true, color: HEX.lt1, fill: { color: '0F3256' }, align: 'center' } },
    ...r.slice(1).map((x) => ({ text: x ? '●' : '', options: { color: HEX.accent1, fill: { color: '0F3256' }, align: 'center' } })),
  ]);
  s.addTable([head, ...rows], { x: 0.6, y: 1.75, w: 5.2, colW: [1.6, 1.2, 1.2, 1.2], fontSize: 15, rowH: 0.55, border: { type: 'solid', pt: 0.5, color: HEX.dk2 }, valign: 'middle', objectName: 'tabla-meolp' });
  txt(s, 'OLP3 depende solo de E5. Por eso salió primera en la MCPE.', { x: 0.6, y: 4.75, w: 5.2, h: 0.9, fontSize: 15, color: C.background1 });
  const lv = { A: ['Alta', HEX.accent1, HEX.dk2], M: ['Media', HEX.accent4, HEX.lt1], B: ['Baja', '0F3256', HEX.lt1], N: ['No aplica', '0F3256', HEX.accent5] };
  const hd = ['MEPCS', 'Plataformas', 'Academias', 'Tutores', 'IA gratis', 'Nuevos'].map((t) => ({ text: t, options: { bold: true, color: HEX.dk2, fill: { color: HEX.accent1 }, align: 'center' } }));
  const mp = [['E5', 'B', 'M', 'B', 'N', 'B'], ['E10', 'M', 'A', 'M', 'A', 'M'], ['E1', 'M', 'A', 'B', 'M', 'M'], ['E2', 'A', 'M', 'B', 'A', 'M']].map((r) => [
    { text: r[0], options: { bold: true, color: HEX.lt1, fill: { color: '0F3256' }, align: 'center' } },
    ...r.slice(1).map((k) => ({ text: lv[k][0], options: { color: lv[k][2], fill: { color: lv[k][1] }, align: 'center' } })),
  ]);
  s.addTable([hd, ...mp], { x: 6.3, y: 1.75, w: 6.4, colW: [0.9, 1.3, 1.2, 0.95, 1.05, 1.0], fontSize: 12, rowH: 0.55, border: { type: 'solid', pt: 0.5, color: HEX.dk2 }, valign: 'middle', objectName: 'tabla-mepcs' });
  txt(s, 'Qué tanto puede responder cada competidor. E10 y E2 son las más expuestas, así que E10 no debe competir solo por precio.', { x: 6.3, y: 4.75, w: 6.4, h: 1.2, fontSize: 15, color: C.background1 });
  s.addNotes('En la MEOLP las cuatro estrategias cubren los tres objetivos, y el cumplimiento depende solo de E5. En la MEPCS vimos que E5 es la más difícil de copiar, porque nadie más tiene datos por clase y seguimiento a los padres. E10 y E2 son las más expuestas a precios bajos y contenido gratis.');
}

// ---------- 7. Mapa estratégico ----------
{
  const s = slide('CLARO', 'Implementación');
  kicker(s, '06 / BALANCED SCORECARD · MAPA');
  s.addText('Todo empieza en los profesores', { placeholder: 'title' });
  card(s, 0.6, 1.65, 8.3, 5.1, C.background1, 'marco-mapa');
  s.addImage({ path: IMG('mapa_bsc_ppt.png'), x: 1.25, y: 1.75, w: 7.0, h: 4.9, objectName: 'mapa' });
  card(s, 9.2, 1.65, 3.5, 5.1, C.text2, 'lectura-mapa');
  txt(s, 'Cómo se lee', { x: 9.45, y: 1.9, w: 3.0, h: 0.4, fontSize: 18, bold: true, color: C.accent1 });
  txt(s, [
    { text: 'Si los profesores tienen su tablero y se quedan,', options: { breakLine: true } },
    { text: 'las clases salen sin faltas ni tardanzas.', options: { breakLine: true } },
    { text: 'Con mejores clases y nuevos programas llegan más alumnos,', options: { breakLine: true } },
    { text: 'y con más alumnos suben los ingresos.' },
  ], { x: 9.45, y: 2.45, w: 3.05, h: 4.0, fontSize: 15, color: C.background1, paraSpaceAfter: 8 });
  s.addNotes('El mapa estratégico tiene once objetivos en las cuatro perspectivas. Se lee de abajo hacia arriba: todo empieza en los profesores. Si cumplen y se quedan, mejoran las clases, llegan más alumnos y suben los ingresos.');
}

// ---------- 8. Tablero ----------
{
  const s = slide('CLARO', 'Implementación');
  kicker(s, '07 / BALANCED SCORECARD · TABLERO');
  s.addText('Cada objetivo tiene indicador, meta y responsable', { placeholder: 'title' });
  const rows = [
    ['Financiera', 'Ingresos promedio al mes', 'S/230 000', 'S/262 700', 'S/300 000', 'Gerencia general'],
    ['Financiera', 'Ingresos de nivelación y grupos', '0 %', '10 %', '20 %', 'Coordinación académica'],
    ['Clientes', 'Alumnos en octubre y noviembre', '288', '380', '500', 'Gerencia general'],
    ['Clientes', 'Compras por cada 100 interesados', '3', '4', '5.2', 'Ventas'],
    ['Procesos', 'Clases con problemas por cada 100', '25', '11', '5', 'Calidad docente'],
    ['Procesos', 'Grados con programa listo', '0', '5', '11', 'Diseño de programas'],
    ['Procesos', 'Interesados que hacen el diagnóstico', '0 %', '30 %', '50 %', 'Marketing y datos'],
    ['Aprendizaje', 'Profesores con tablero mensual', '0', '27', '27', 'Calidad docente'],
    ['Aprendizaje', 'Profesores que se van al año', '10', '6', '4', 'Calidad docente'],
    ['Aprendizaje', 'Profesores capacitados en grupo y nivelación', '0', '10', '20', 'Capacitación'],
  ];
  const head = ['Perspectiva', 'Indicador', 'Base 2026', 'Meta 2027', 'Meta 2028', 'Responsable'].map((t) => ({ text: t, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, align: 'center' } }));
  const body = rows.map((r, i) => r.map((c, k) => ({ text: c, options: { color: HEX.dk1, fill: { color: i % 2 ? HEX.lt1 : HEX.accent6 }, align: k >= 2 && k <= 4 ? 'center' : 'left', bold: k === 4 } })));
  s.addTable([head, ...body], { x: 0.6, y: 1.65, w: 12.1, colW: [1.6, 4.3, 1.35, 1.35, 1.35, 2.15], fontSize: 13, rowH: 0.45, border: { type: 'solid', pt: 0.5, color: HEX.accent5 }, valign: 'middle', objectName: 'tablero' });
  s.addNotes('El tablero pone a cada objetivo un indicador, la base de 2026, las metas de 2027 y 2028 y un responsable. Se revisa una vez al mes en la reunión de áreas que ya existe, así que no hace falta crear reuniones nuevas.');
}

// ---------- 9. Estructura ----------
{
  const s = slide('CLARO', 'Implementación');
  kicker(s, '08 / ESTRUCTURA ORGANIZACIONAL');
  s.addText('Mismas 13 personas, con un dueño para cada estrategia', { placeholder: 'title' });
  txt(s, '¿LA ESTRUCTURA ACTUAL SIRVE?', { x: 0.6, y: 1.65, w: 4.5, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: C.accent2 });
  const P = [['E5', 'En parte', 'Revisan, pero nadie responde por el resultado'], ['E10', 'En parte', 'Nadie llena los grupos'], ['E1', 'Sí', 'Talleres y capacitación ya tienen responsables'], ['E2', 'En parte', 'Falta quien maneje campañas y datos']];
  P.forEach(([c, v, d], i) => {
    const y = 2.05 + i * 1.18;
    card(s, 0.6, y, 4.6, 1.05, i === 2 ? C.accent3 : C.background1, `pert-${c}`);
    circ(s, 0.75, y + 0.17, 0.7, c, C.text2, C.accent1, 12, `pc-${c}`);
    txt(s, v, { x: 1.6, y: y + 0.12, w: 3.4, h: 0.35, fontSize: 15, bold: true, color: C.text1 });
    txt(s, d, { x: 1.6, y: y + 0.48, w: 3.5, h: 0.5, fontSize: 12, color: C.accent4 });
  });
  card(s, 5.5, 1.65, 7.2, 5.1, C.background1, 'marco-org');
  s.addImage({ path: IMG('organigrama_nuevo_ppt.png'), x: 6.04, y: 1.75, w: 6.125, h: 4.9, objectName: 'organigrama' });
  s.addNotes('La estructura actual sirve para E1 y solo en parte para las otras. En académica cada persona es responsable de un tema, pero nadie une ese trabajo ni responde por el resultado. Propongo una coordinación académica fija, una unidad de calidad docente para E5 y una función de marketing y datos para E2, sin contratar a nadie.');
}

// ---------- 10. Conclusiones ----------
pres.addSection({ title: 'Cierre' });
{
  const s = slide('OSCURO', 'Cierre');
  kicker(s, '09 / CONCLUSIONES', true);
  s.addText('El primer paso es que los profesores cumplan', { placeholder: 'title' });
  const F = [
    ['1', 'Tres objetivos para 2028, cada uno con tres objetivos de corto plazo calculados'],
    ['2', 'Las cuatro estrategias cubren los objetivos y E5 es la más difícil de copiar'],
    ['3', 'Once objetivos en el mapa, con indicador, meta y responsable en el tablero'],
    ['4', 'Nueva estructura sin contratar: cada estrategia con un dueño y un indicador'],
  ];
  F.forEach(([n, t], i) => {
    const y = 1.75 + i * 0.85;
    circ(s, 0.6, y, 0.65, n, C.accent1, C.text2, 18, `paso-${n}`);
    txt(s, t, { x: 1.45, y: y + 0.1, w: 11, h: 0.5, fontSize: 18, color: C.background1 });
  });
  txt(s, 'REFERENCIAS', { x: 0.6, y: 5.3, w: 4, h: 0.25, fontSize: 10, bold: true, charSpacing: 2, color: C.accent1 });
  txt(s, [
    { text: 'David, F. R. (2017). Conceptos de administración estratégica (15.ª ed.). Pearson Educación.', options: { breakLine: true } },
    { text: 'Forbes Perú. (2026, 20 de mayo). Hogares peruanos con acceso a internet fijo o móvil se elevó a 96 % en el 2025, según Erestel. https://forbes.pe/actualidad/2026-05-20/hogares-peruanos-con-acceso-a-internet-fijo-o-movil-se-elevo-a-96-en-el-2025-segun-erestel/', options: { breakLine: true } },
    { text: 'MatLab Perú. (s. f.). Inicio [Sitio web]. https://matlab.com.pe/', options: { breakLine: true } },
    { text: 'Ministerio de Educación. (2026). ENLA 2025: resumen ejecutivo. https://repositorio.minedu.gob.pe/handle/20.500.12799/12773' },
  ], { x: 0.6, y: 5.6, w: 12.1, h: 1.2, fontSize: 10, color: C.accent5, paraSpaceAfter: 2 });
  s.addNotes('En resumen, el plan tiene tres objetivos medibles para 2028, cuatro estrategias que los cubren, un tablero para seguir el avance cada mes y una estructura donde cada estrategia tiene un responsable. Lo primero es aplicar E5, porque todo empieza en que los profesores cumplan. Gracias.');
}

pres.writeFile({ fileName: OUT }).then(async () => { await applyTheme(OUT, THEME); console.log('ok', OUT); });
