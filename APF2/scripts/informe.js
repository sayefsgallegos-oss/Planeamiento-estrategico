// Arma el informe APF2 completo en Word con formato APA 7.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, WidthType, ShadingType, BorderStyle, Header, PageNumber,
  HeadingLevel, LevelFormat, PageBreak, VerticalAlign,
} = require('docx');

const DIR = path.join(__dirname, '..');
const FIG = (f) => fs.readFileSync(path.join(DIR, 'figuras', f));
const FONT = 'Times New Roman';
const W = 9026; // ancho útil A4 con márgenes de 1 pulgada

// ---------- piezas de texto ----------
const runs = (txt, base = {}) => {
  // **negrita** y *cursiva* simples dentro del texto
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/g;
  let last = 0, m;
  while ((m = re.exec(txt))) {
    if (m.index > last) out.push(new TextRun({ text: txt.slice(last, m.index), ...base }));
    const t = m[0];
    if (t.startsWith('**')) out.push(new TextRun({ ...base, text: t.slice(2, -2), bold: true }));
    else out.push(new TextRun({ ...base, text: t.slice(1, -1), italics: true }));
    last = m.index + t.length;
  }
  if (last < txt.length) out.push(new TextRun({ text: txt.slice(last), ...base }));
  return out;
};

const P = (txt) => new Paragraph({ children: runs(txt), indent: { firstLine: 720 }, spacing: { line: 480 } });
const PC = (txt, o = {}) => new Paragraph({ children: runs(txt, o), alignment: AlignmentType.CENTER, spacing: { line: 480 } });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(t)] });
const LI = (t) => new Paragraph({ children: runs(t), numbering: { reference: 'etica', level: 0 }, spacing: { line: 480 } });
const BR = () => new Paragraph({ children: [new PageBreak()] });

let nTabla = 0, nFig = 0;
const caption = (kind, n, title) => [
  new Paragraph({ children: [new TextRun({ text: `${kind} ${n}`, bold: true })], spacing: { before: 240, line: 240 }, keepNext: true }),
  new Paragraph({ children: [new TextRun({ text: title, italics: true })], spacing: { after: 60, line: 240 }, keepNext: true }),
];
const nota = (txt) => new Paragraph({
  children: [new TextRun({ text: 'Nota. ', italics: true, size: 18 }), ...runs(txt, { size: 18 })],
  spacing: { before: 60, after: 240, line: 240 },
});

// ---------- tablas ----------
const border = { style: BorderStyle.SINGLE, size: 4, color: '808080' };
const borders = { top: border, bottom: border, left: border, right: border };
const GRIS = 'D9D9D9', GRIS2 = 'F2F2F2';

function cell(content, width, o = {}) {
  const lines = Array.isArray(content) ? content : [content];
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    verticalAlign: VerticalAlign.CENTER,
    columnSpan: o.span,
    children: lines.map((l) => new Paragraph({
      alignment: o.align || AlignmentType.LEFT,
      spacing: { line: 240 },
      children: runs(String(l), { size: o.size || 20, bold: o.bold }),
    })),
  });
}

// rows: arreglo de filas; cada fila es arreglo de celdas (texto o {t, span, fill, bold, align})
function tabla(widths, header, rows, o = {}) {
  const size = o.size || 20;
  const center = o.center || [];
  const mk = (r, isHead) => new TableRow({
    tableHeader: isHead,
    cantSplit: true,
    children: (() => {
      let col = 0;
      return r.map((c) => {
        const obj = typeof c === 'object' && !Array.isArray(c) ? c : { t: c };
        const span = obj.span || 1;
        const w = widths.slice(col, col + span).reduce((a, b) => a + b, 0);
        const align = obj.align || (isHead || center.includes(col) ? AlignmentType.CENTER : AlignmentType.LEFT);
        col += span;
        return cell(obj.t, w, { span: span > 1 ? span : undefined, fill: isHead ? GRIS : obj.fill, bold: isHead || obj.bold, align, size });
      });
    })(),
  });
  const all = [];
  if (header) (Array.isArray(header[0]) ? header : [header]).forEach((h) => all.push(mk(h, true)));
  rows.forEach((r) => all.push(mk(r, false)));
  return new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows: all });
}

const T = (title, widths, header, rows, notaTxt, o) => [
  ...caption('Tabla', ++nTabla, title), tabla(widths, header, rows, o), nota(notaTxt),
];
const FIGURA = (title, file, w, h, notaTxt) => [
  ...caption('Figura', ++nFig, title),
  new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, children: [new ImageRun({ type: 'png', data: FIG(file), transformation: { width: w, height: h } })] }),
  nota(notaTxt),
];
const SEC = (t) => ({ t, span: 99, fill: GRIS2, bold: true });
const fixSpan = (rows, n) => rows.map((r) => r.map((c) => (c && c.span === 99 ? { ...c, span: n } : c)));

// ---------- datos de la MCPE ----------
const E_MCPE = ['E1', 'E2', 'E4', 'E5', 'E7', 'E10'];
const MCPE = [
  ['O1', 'Alta conectividad de los hogares', 0.10, [3, 4, 2, 1, 2, 3]],
  ['O2', 'Bajo logro de los escolares en matemática', 0.12, [4, 3, 2, 3, 3, 3]],
  ['O3', 'Posible mejora de la economía de las familias', 0.07, [3, 2, 2, 1, 2, 4]],
  ['O4', 'Empresas que buscan educación financiera', 0.08, [2, 2, 3, 1, 2, 3]],
  ['O5', 'Tecnologías digitales e inteligencia artificial', 0.08, [2, 4, 3, 4, 1, 1]],
  ['A1', 'Posible alza de precios y de costos', 0.08, [2, 3, 2, 1, 2, 4]],
  ['A2', 'Riesgo de fenómenos climáticos', 0.05, [2, 1, 1, 3, 2, 2]],
  ['A3', 'Cambios de gobierno y nuevas normas', 0.05, [3, 2, 1, 2, 2, 2]],
  ['A4', 'Normas de protección de datos de menores', 0.08, [2, 1, 2, 3, 2, 2]],
  ['A5', 'Competencia de bajo costo y sustitutos gratuitos', 0.15, [3, 3, 2, 3, 4, 4]],
  ['A6', 'Disputa por docentes calificados', 0.14, [2, 1, 1, 4, 2, 3]],
  ['F1', 'Enseñanza personalizada 1 a 1', 0.12, [4, 3, 2, 3, 4, 2]],
  ['F2', 'Trayectoria y marca conocida', 0.06, [3, 3, 2, 2, 4, 2]],
  ['F3', 'Experiencia en clases 100 % virtuales', 0.10, [3, 4, 2, 3, 2, 3]],
  ['F4', 'Seguimiento y comunicación con los padres', 0.08, [4, 2, 3, 3, 4, 3]],
  ['F5', 'Oferta variada de servicios', 0.06, [4, 2, 3, 1, 2, 4]],
  ['F6', 'Cultura de innovación y mejora de procesos', 0.06, [2, 4, 3, 4, 1, 2]],
  ['D1', 'Dependencia de los profesores', 0.11, [2, 1, 1, 4, 2, 3]],
  ['D2', 'Faltas, tardanzas y cámara apagada', 0.11, [1, 1, 1, 4, 1, 2]],
  ['D3', 'Malas prácticas de algunos docentes', 0.08, [2, 1, 1, 4, 1, 2]],
  ['D4', 'Desorden en el área comercial', 0.09, [2, 3, 4, 1, 3, 2]],
  ['D5', 'Falta de inscripciones', 0.08, [4, 4, 4, 2, 3, 3]],
  ['D6', 'Datos e inteligencia artificial en proceso', 0.05, [1, 3, 3, 4, 1, 1]],
];
const totMCPE = E_MCPE.map((_, i) => MCPE.reduce((s, f) => s + f[2] * f[3][i], 0));

// ---------- contenido ----------
const portada = [
  new Paragraph({ spacing: { before: 1200 }, children: [] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new ImageRun({ type: 'png', data: FIG('logo_utp.png'), transformation: { width: 240, height: 91 } })] }),
  PC('**Universidad Tecnológica del Perú**'),
  PC('Facultad de Ingeniería'),
  PC(''),
  PC('**Planeamiento Estratégico**'),
  PC('Sección 20726'),
  PC(''),
  PC('**Avance de Proyecto Final 2**'),
  PC('*Diseño del plan estratégico de MatLab Perú*'),
  PC(''),
  PC('Alumno: Pierre Gallegos Mallqui'),
  PC('Código: U23211135'),
  PC('Docente: Zeyla Casaverde Warthon'),
  PC('Lima, Perú'),
  PC('2026'),
  BR(),
];

const sec1 = [
  H1('1. La presentación de la empresa'),
  H2('1.1. La introducción y la misión, visión y valores'),
  H3('Introducción'),
  P('Este trabajo es el segundo avance de mi proyecto final de Planeamiento Estratégico. En el primer avance hice el diagnóstico de MatLab Perú, una academia que enseña matemáticas y educación financiera. En este avance uso ese diagnóstico para diseñar el plan estratégico y elegir las estrategias de mayor impacto. Trabajo en la empresa, por eso uso información interna y mi propia experiencia, y la empresa autorizó su uso en este trabajo. Sigo la metodología de Fred David. Los pesos y las calificaciones de las matrices los puse yo, y en cada una explico por qué.'),
  H3('Presentación de la empresa'),
  P('MatLab Perú es una academia peruana. Su página web dice que quiere que las personas se lleven bien con las matemáticas y las finanzas (MatLab Perú, s. f.).'),
  P('Marco Loret de Mola la fundó en 2012 para enseñar matemáticas con clases personalizadas. En 2020 pasó a dar todas sus clases de forma virtual, y eso le dio más clientes y más alcance en redes. En los meses de más demanda ha llegado a atender unos 500 alumnos, aunque es un cálculo aproximado.'),
  P('Atiende a familias y estudiantes, desde primaria hasta la universidad, y también a empresas que compran programas de educación financiera. Las clases individuales duran unos 50 minutos y cuestan entre S/55 y S/65. También hay paquetes de horas, talleres en grupo y programas de verano.'),
  H3('Misión, visión, valores y código de ética'),
  P('La empresa no tiene una misión, una visión, valores ni un código de ética escritos de forma estable. Los propongo yo, a partir de lo que conozco de ella.'),
  P('La misión que propongo es esta: enseñamos matemáticas y educación financiera de forma personalizada y virtual a estudiantes, familias y organizaciones. Adaptamos cada clase al ritmo de la persona y seguimos su avance con ayuda de la tecnología, para que entienda los números y los use en su vida.'),
  P('La visión que propongo es esta: ser la academia virtual más reconocida del Perú en matemáticas y educación financiera, y que cada persona que pasa por MatLab pierda el miedo a los números.'),
  P('Los valores que propongo son personalización, cercanía y confianza, comprensión, compromiso y puntualidad, innovación e integridad. El código de ética tiene seis compromisos:'),
  LI('Tratar con respeto a cada estudiante, sobre todo a los menores de edad.'),
  LI('Cuidar los datos de alumnos, familias y clientes.'),
  LI('Informar con honestidad sobre los servicios y los precios.'),
  LI('Cumplir los horarios y las clases programadas.'),
  LI('Preparar las clases y mantener la cámara encendida.'),
  LI('Usar la tecnología y la inteligencia artificial de forma responsable, sin reemplazar al profesor.'),
];

const sec2 = [
  H1('2. El diagnóstico interno y externo de la empresa'),
  H2('2.1. El contexto global, los factores críticos de éxito y la evaluación externa PESTEL'),
  P('Tres cambios mundiales afectan a MatLab. El internet es parte de la vida diaria, la inteligencia artificial está cambiando la forma de estudiar y las plataformas de tutorías permiten contratar profesores de cualquier lugar.'),
  P('Los factores críticos de éxito son lo que más pesa para ganar clientes en un sector. Para una academia como MatLab elegí ocho, pensando en lo que mira una familia al escoger: calidad y personalización, confianza en la marca, precio, puntualidad y cumplimiento, cobertura de cursos, seguimiento a los padres, tecnología, y captación y atención comercial. Los uso en la MPC.'),
  ...T('Análisis PESTEL de MatLab Perú', [1500, 5926, 1600], ['Dimensión', 'Análisis', 'Factor'], [
    [{ t: 'Político', bold: true }, 'Los cambios de gobierno pueden traer nuevas normas en educación, y eso genera incertidumbre.', 'Amenaza A3'],
    [{ t: 'Económico', bold: true }, 'Si la economía de las familias mejora, pueden gastar más en clases extra. Si suben los precios, pueden gastar menos.', ['Oportunidad O3', 'Amenaza A1']],
    [{ t: 'Sociocultural', bold: true }, 'Pocos estudiantes logran el nivel esperado en matemática. Solo el 12.8 % de los alumnos de quinto de secundaria lo logra (Ministerio de Educación, 2026). Eso crea demanda de refuerzo. MatLab también ha trabajado con empresas que quieren programas de educación financiera.', ['Oportunidad O2', 'Oportunidad O4']],
    [{ t: 'Tecnológico', bold: true }, 'El 96 % de los hogares tuvo internet en 2025 (Forbes Perú, 2026), y por eso todas las clases pueden ser virtuales. También es una amenaza, porque hoy los estudiantes pueden usar inteligencia artificial para resolver dudas sin pagar clases.', ['Oportunidad O1', 'Oportunidad O5', 'Amenaza A5']],
    [{ t: 'Ecológico', bold: true }, 'Fenómenos como El Niño Costero pueden afectar a las familias y cortar la luz o el internet. Como las clases son virtuales, un corte afecta el servicio.', 'Amenaza A2'],
    [{ t: 'Legal', bold: true }, 'La ley de protección de datos personales obliga a cuidar los datos de los alumnos. Muchos son menores de edad, así que hay que tener más cuidado.', 'Amenaza A4'],
  ], 'Elaboración propia. Los códigos corresponden a la MEFE de la Tabla 2.'),
  H2('2.2. Matriz de evaluación de factores externos MEFE'),
  P('En la MEFE cada factor tiene un peso, y todos suman 1.0, y un valor de 1 a 4 según qué tan bien responde la empresa: 4 es muy bien y 1 es mal. El promedio es 2.5.'),
  ...T('Matriz de evaluación de factores externos MEFE', [600, 2400, 650, 650, 800, 3926], ['Cód.', 'Factor externo clave', 'Peso', 'Valor', 'Puntaje', 'Por qué ese valor'], fixSpan([
    [SEC('OPORTUNIDADES')],
    ['O1', 'Alta conectividad de los hogares, con 96 % de acceso a internet en 2025', '0.10', '4', '0.40', 'Todas las clases son virtuales desde 2020, así que la empresa aprovecha muy bien esto.'],
    ['O2', 'Bajo logro de los escolares en matemática', '0.12', '3', '0.36', 'La enseñanza personalizada responde a esta necesidad, pero en la empresa cuesta conseguir inscripciones.'],
    ['O3', 'Posible mejora de la economía de las familias, que podrían gastar más en educación', '0.07', '2', '0.14', 'La empresa no ha logrado aprovechar esto para conseguir más alumnos, y sus ingresos se han debilitado.'],
    ['O4', 'Empresas e instituciones que buscan programas de educación financiera y apoyo educativo', '0.08', '3', '0.24', 'MatLab ya ofrece educación financiera y programas para empresas, y tiene experiencia con estos clientes.'],
    ['O5', 'Tecnologías digitales e inteligencia artificial para personalizar y automatizar', '0.08', '3', '0.24', 'La empresa usa CRM, tableros y pizarra virtual, y ha probado inteligencia artificial.'],
    [SEC('AMENAZAS')],
    ['A1', 'Posible alza de precios y de costos, que afecta a las familias y a la empresa', '0.08', '2', '0.16', 'La empresa cambió el pago a los profesores para bajar costos, pero sus ingresos se han debilitado.'],
    ['A2', 'Riesgo de fenómenos climáticos, como El Niño Costero', '0.05', '2', '0.10', 'Las clases virtuales evitan traslados, pero un corte de luz o de internet afecta la clase.'],
    ['A3', 'Cambios de gobierno que pueden traer nuevas normas en educación', '0.05', '2', '0.10', 'Es algo que la empresa no controla.'],
    ['A4', 'Normas de protección de datos personales, sobre todo de menores', '0.08', '2', '0.16', 'La empresa maneja datos de alumnos, muchos menores de edad, con herramientas digitales.'],
    ['A5', 'Competencia de bajo costo y sustitutos gratuitos, como la inteligencia artificial', '0.15', '2', '0.30', 'MatLab se diferencia por la personalización, pero le cuesta conseguir alumnos nuevos.'],
    ['A6', 'Disputa por docentes calificados con otras academias y plataformas', '0.14', '1', '0.14', 'Le di 1 porque en la empresa hay profesores que faltan, llegan tarde o tienen malas prácticas.'],
    ['', { t: 'TOTAL', bold: true }, { t: '1.00', bold: true }, '', { t: '2.34', bold: true }, ''],
  ], 6), 'Elaboración propia.', { size: 18, center: [0, 2, 3, 4] }),
  P('El total es 2.34, menor que 2.5. Hoy la empresa no aprovecha del todo las oportunidades ni se protege bien de las amenazas. Solo responde muy bien a la conectividad. Le va peor con las amenazas de más peso, que son la competencia de bajo costo y la disputa por docentes.'),
  H2('2.3. Modelo de las 5 fuerzas de Porter'),
  P('Las cinco fuerzas de Porter muestran qué tan atractivo es un sector (David, 2017). A cada una le puse un nivel.'),
  ...T('Las cinco fuerzas de Porter en el sector de clases de matemática', [2000, 900, 6126], ['Fuerza', 'Nivel', 'Por qué'], [
    [{ t: 'Rivalidad entre competidores', bold: true }, 'Alta', 'Hay muchas academias, tutores independientes y plataformas de tutorías. Los padres pueden comparar y cambiar fácil. En MatLab nos cuesta conseguir alumnos nuevos.'],
    [{ t: 'Nuevos competidores', bold: true }, 'Alta', 'Empezar cuesta poco: se necesita una computadora, internet y un profesor. Lo difícil es ganar la confianza de las familias.'],
    [{ t: 'Productos sustitutos', bold: true }, 'Alta', 'Hay videos y contenido gratis, aplicaciones y herramientas de inteligencia artificial que resuelven dudas. También el refuerzo del colegio.'],
    [{ t: 'Poder de los clientes', bold: true }, 'Alta', 'Los padres comparan precios y cambian de academia con facilidad. Las empresas que contratan programas negocian más porque compran por volumen.'],
    [{ t: 'Poder de los proveedores', bold: true }, 'Alta', 'Los profesores son lo más importante del servicio y pueden trabajar en otras academias o por su cuenta. En MatLab dependemos mucho de ellos. La tecnología que usamos se puede reemplazar, así que ahí el poder es bajo.'],
  ], 'Elaboración propia.', { center: [1] }),
  P('Las cinco fuerzas son altas, así que hay mucha competencia. Para sostenerse, MatLab tiene que diferenciarse con la personalización y cuidar el cumplimiento de sus profesores.'),
  H2('2.4. Matriz de perfil competitivo MPC'),
  P('La MPC compara a la empresa con sus competidores en los factores críticos de éxito. Cada factor tiene un peso y una calificación de 1 a 4: 4 es fortaleza mayor, 3 fortaleza menor, 2 debilidad menor y 1 debilidad mayor. Comparé a MatLab con tres grupos: plataformas donde se elige entre muchos tutores, academias virtuales y tutores independientes. Las calificaciones de los competidores son estimaciones mías.'),
  ...T('Matriz de perfil competitivo MPC', [2226, 600, 775, 775, 775, 775, 775, 775, 775, 775],
    [['Factor crítico de éxito', 'Peso', { t: 'MatLab Perú', span: 2 }, { t: 'Plataformas de tutorías', span: 2 }, { t: 'Academias virtuales', span: 2 }, { t: 'Tutores independientes', span: 2 }],
     ['', '', 'Calif.', 'Pond.', 'Calif.', 'Pond.', 'Calif.', 'Pond.', 'Calif.', 'Pond.']], [
    ['Calidad y personalización de la enseñanza', '0.20', '4', '0.80', '2', '0.40', '3', '0.60', '3', '0.60'],
    ['Confianza y reputación de la marca', '0.12', '3', '0.36', '3', '0.36', '2', '0.24', '2', '0.24'],
    ['Precio', '0.12', '2', '0.24', '4', '0.48', '3', '0.36', '4', '0.48'],
    ['Puntualidad y cumplimiento del servicio', '0.14', '1', '0.14', '2', '0.28', '2', '0.28', '2', '0.28'],
    ['Cobertura de niveles y cursos', '0.08', '3', '0.24', '4', '0.32', '3', '0.24', '2', '0.16'],
    ['Seguimiento y reportes a los padres', '0.12', '3', '0.36', '1', '0.12', '2', '0.24', '1', '0.12'],
    ['Tecnología y plataforma virtual', '0.10', '3', '0.30', '4', '0.40', '3', '0.30', '2', '0.20'],
    ['Captación y atención comercial', '0.12', '2', '0.24', '4', '0.48', '3', '0.36', '2', '0.24'],
    [{ t: 'Total', bold: true }, { t: '1.00', bold: true }, '', { t: '2.68', bold: true }, '', { t: '2.84', bold: true }, '', { t: '2.62', bold: true }, '', { t: '2.32', bold: true }],
  ], 'Elaboración propia.', { size: 18, center: [1, 2, 3, 4, 5, 6, 7, 8, 9] }),
  P('MatLab saca 2.68 y queda segunda, detrás de las plataformas de tutorías, que sacan 2.84. Gana en personalización y en seguimiento a los padres. Pierde sobre todo en puntualidad y cumplimiento, donde le puse 1 por los problemas con los profesores que conozco de la empresa. También queda baja en precio y captación comercial.'),
  H2('2.5. La evaluación interna AMOFHIT'),
  P('El AMOFHIT revisa siete áreas de la empresa. Lo escribo con información interna y con lo que conozco de mi trabajo.'),
  H3('Administración'),
  P('La empresa tiene tres áreas: Gerencia con 2 personas, Comercial con 6 y Académica con 5. Con los 27 profesores, somos 40 personas. Antes había más áreas. Las áreas se reúnen en persona una vez por semana para ver avances. Por mi experiencia, el área comercial es la que tiene más desorden.'),
  H3('Marketing'),
  P('Los alumnos llegan sobre todo por internet: la página web, las redes y las recomendaciones. Hay un protocolo de ventas y un CRM para seguir a los interesados. Desde 2020 la empresa creció mucho en redes y clientes, pero hoy cuesta conseguir inscripciones.'),
  H3('Operaciones'),
  P('Las clases son por Google Meet, con pizarra virtual y reservas en línea. Cada profesor llena un informe después de cada clase. La capacidad de la empresa depende de las horas libres de los profesores. Por mi experiencia, algunos faltan, llegan tarde o no prenden la cámara, y eso pesa mucho porque el cliente solo ve la clase.'),
  H3('Finanzas'),
  P('La empresa gana con clases individuales, paquetes de horas, talleres, programas de verano y programas para empresas. El costo más grande son las horas de los profesores. En mi experiencia, los ingresos se han debilitado, sobre todo en el producto que depende de los profesores.'),
  H3('Recursos humanos'),
  P('Hay 13 personas en las áreas de gestión y 27 profesores. Los profesores reciben asesoría académica cada tres semanas y hay una integración cada dos meses. Todo el personal recibió una capacitación con un asesor sobre protocolos de trabajo. Aun así, falta controlar mejor el desempeño de los profesores.'),
  H3('Informática'),
  P('La empresa usa Google Workspace para guardar los datos de las clases, una plataforma de reservas, el CRM Kommo para las ventas y Looker Studio para hacer tableros. Está tratando de convertir los datos de cada clase en indicadores.'),
  H3('Tecnología'),
  P('Se implementó una pizarra virtual y se han probado ideas con inteligencia artificial, como reportes automáticos para los padres, un examen de diagnóstico y un asistente para ventas. Están en etapas distintas de avance.'),
  H2('2.6. Matriz de evaluación de factores internos MEFI'),
  P('La MEFI se arma igual que la MEFE, pero con fortalezas y debilidades. El 4 es fortaleza mayor, el 3 fortaleza menor, el 2 debilidad menor y el 1 debilidad mayor. Un total menor a 2.5 indica una posición interna débil.'),
  ...T('Matriz de evaluación de factores internos MEFI', [600, 2400, 650, 650, 800, 3926], ['Cód.', 'Factor interno clave', 'Peso', 'Valor', 'Puntaje', 'Por qué ese valor'], fixSpan([
    [SEC('FORTALEZAS')],
    ['F1', 'Enseñanza personalizada 1 a 1 y objetivo de que el alumno se lleve bien con las matemáticas', '0.12', '4', '0.48', 'Es lo que más nos diferencia. Cada clase se adapta al nivel y al ritmo del alumno.'],
    ['F2', 'Más de diez años de trayectoria y marca conocida', '0.06', '3', '0.18', 'La empresa se fundó en 2012 y tiene experiencia con familias y empresas.'],
    ['F3', 'Experiencia en clases 100 % virtuales desde 2020', '0.10', '4', '0.40', 'Todas las clases son virtuales, con pizarra, reservas en línea y CRM.'],
    ['F4', 'Seguimiento académico y comunicación con los padres', '0.08', '3', '0.24', 'Cada profesor llena un informe por clase y se informa a las familias.'],
    ['F5', 'Oferta variada de servicios', '0.06', '3', '0.18', 'Hay clases individuales, talleres, programas de verano y educación financiera.'],
    ['F6', 'Cultura de innovación y mejora de procesos', '0.06', '3', '0.18', 'La empresa cambió el pago a los profesores, usa tableros y prueba inteligencia artificial.'],
    [SEC('DEBILIDADES')],
    ['D1', 'Dependencia de los profesores para dar el servicio', '0.11', '1', '0.11', 'El servicio y los costos dependen de las horas de los profesores. Los ingresos del producto que depende de ellos se han debilitado.'],
    ['D2', 'Faltas, tardanzas y cámara apagada en las clases virtuales', '0.11', '1', '0.11', 'Por mi experiencia en la empresa, es un problema que se repite y el cliente lo nota enseguida.'],
    ['D3', 'Malas prácticas de algunos docentes', '0.08', '2', '0.16', 'La calidad de las clases no es igual con todos los profesores.'],
    ['D4', 'Desorden en el área comercial', '0.09', '2', '0.18', 'En mi experiencia, las ventas y el seguimiento a los interesados están desordenados.'],
    ['D5', 'Falta de inscripciones', '0.08', '1', '0.08', 'Las inscripciones no alcanzan. En los meses de más demanda se llegó a unos 500 alumnos.'],
    ['D6', 'Datos e inteligencia artificial todavía en proceso', '0.05', '2', '0.10', 'Los datos de las clases se están convirtiendo en indicadores, pero es un trabajo en proceso.'],
    ['', { t: 'TOTAL', bold: true }, { t: '1.00', bold: true }, '', { t: '2.40', bold: true }, ''],
  ], 6), 'Elaboración propia.', { size: 18, center: [0, 2, 3, 4] }),
  P('El total es 2.40, un poco menos que 2.5. Las fortalezas son buenas, pero las debilidades pesan más de la mitad del total. Las más graves son la dependencia de los profesores y sus faltas y tardanzas, porque el cliente ve el servicio casi solo en la clase. El método es bueno, pero solo sirve si el profesor cumple.'),
  H2('2.7. Cadena de valor'),
  P('La cadena de valor muestra dónde la empresa crea valor y dónde tiene problemas (David, 2017).'),
  ...T('Cadena de valor de MatLab Perú', [1100, 2000, 3726, 2200], ['Tipo', 'Actividad', 'En MatLab', 'Valoración'], [
    ['Primaria', 'Logística interna', 'Asignar profesores y recibir los datos del alumno.', 'Debilidad'],
    ['Primaria', 'Operaciones', 'Clases individuales y talleres virtuales.', 'Fortaleza, con fallas de cumplimiento'],
    ['Primaria', 'Logística externa', 'Reserva de horarios y conexión por videollamada.', 'Fortaleza'],
    ['Primaria', 'Marketing y ventas', 'Página web, redes, CRM y protocolo de ventas.', 'Debilidad'],
    ['Primaria', 'Servicio posterior', 'Informes por clase y avisos a los padres.', 'Fortaleza'],
    ['Apoyo', 'Infraestructura', 'Tres áreas y reuniones semanales.', 'Fortaleza moderada'],
    ['Apoyo', 'Recursos humanos', 'Selección, capacitación y asesoría a profesores.', 'Debilidad'],
    ['Apoyo', 'Desarrollo tecnológico', 'Pizarra virtual, tableros e inteligencia artificial.', 'Fortaleza'],
    ['Apoyo', 'Abastecimiento', 'Herramientas y licencias digitales.', 'Adecuado'],
  ], 'Elaboración propia.'),
  P('El valor se crea sobre todo en las clases y en el seguimiento a las familias. Los puntos débiles son recursos humanos y marketing y ventas.'),
  H2('2.8. Matriz FODA'),
  P('La matriz FODA junta los factores internos de la MEFI y los externos de la MEFE, con sus códigos.'),
  ...T('Matriz FODA de MatLab Perú', [4513, 4513], ['FORTALEZAS', 'DEBILIDADES'], [
    [['**F1.** Enseñanza personalizada 1 a 1 y objetivo de que el alumno se lleve bien con las matemáticas', '**F2.** Más de diez años de trayectoria y marca conocida', '**F3.** Experiencia en clases 100 % virtuales desde 2020', '**F4.** Seguimiento académico y comunicación con los padres', '**F5.** Oferta variada de servicios', '**F6.** Cultura de innovación y mejora de procesos'],
     ['**D1.** Dependencia de los profesores para dar el servicio', '**D2.** Faltas, tardanzas y cámara apagada en las clases virtuales', '**D3.** Malas prácticas de algunos docentes', '**D4.** Desorden en el área comercial', '**D5.** Falta de inscripciones', '**D6.** Datos e inteligencia artificial todavía en proceso']],
    [{ t: 'OPORTUNIDADES', bold: true, fill: GRIS, align: AlignmentType.CENTER }, { t: 'AMENAZAS', bold: true, fill: GRIS, align: AlignmentType.CENTER }],
    [['**O1.** Alta conectividad de los hogares', '**O2.** Bajo logro de los escolares en matemática', '**O3.** Posible mejora de la economía de las familias', '**O4.** Empresas que buscan educación financiera y apoyo educativo', '**O5.** Tecnologías digitales e inteligencia artificial'],
     ['**A1.** Posible alza de precios y de costos', '**A2.** Riesgo de fenómenos climáticos', '**A3.** Cambios de gobierno y nuevas normas', '**A4.** Normas de protección de datos de menores', '**A5.** Competencia de bajo costo y sustitutos gratuitos', '**A6.** Disputa por docentes calificados']],
  ], 'Elaboración propia. Las fortalezas y debilidades salen de la MEFI de la Tabla 5, y las oportunidades y amenazas, de la MEFE de la Tabla 2.'),
  H2('2.9. Conclusiones'),
  P('MatLab hace bien su enseñanza personalizada y tiene mucha experiencia en clases virtuales. Además, da seguimiento a las familias con informes por clase.'),
  P('Sus principales problemas están adentro. Depende mucho de los profesores, algunos faltan o llegan tarde, el área comercial está desordenada y no alcanzan las inscripciones. Por eso la MEFI dio 2.40 y la MEFE 2.34, los dos por debajo de 2.5.'),
  P('Tiene oportunidades claras: casi todos los hogares tienen internet, pocos estudiantes logran el nivel esperado en matemática y hay empresas interesadas en educación financiera. Pero el sector tiene mucha competencia.'),
  P('Según el diagnóstico, lo primero que debería atender es el cumplimiento de los profesores y la captación de alumnos, para aprovechar mejor su ventaja en personalización.'),
];

const cuadr = (titulo, items) => [`**${titulo}**`, ...items.flatMap(([c, t, f]) => [`**${c}.** ${t}`, `*Factores: ${f}*`])];

const sec3 = [
  H1('3. Diseño del plan estratégico'),
  P('En esta parte uso los resultados del diagnóstico para elegir las estrategias de MatLab Perú. Primero cruzo los factores del FODA para generar estrategias. Luego las reviso con otras matrices y al final me quedo con las que más convienen a la empresa (David, 2017).'),
  H3('Objetivo planteado'),
  P('Todas las estrategias apuntan a un mismo objetivo: volver a tener unos 500 alumnos en octubre y noviembre de 2028, mejorando el cumplimiento de los profesores. Son los meses de más demanda, porque los padres buscan clases para que sus hijos no desaprueben el año escolar. Elegí este objetivo porque el diagnóstico mostró dos problemas principales, la falta de inscripciones y las faltas y tardanzas de algunos docentes. Por mi experiencia en la empresa, la cantidad de alumnos es lo que más mueve los ingresos. Entre 2020 y 2022 llegamos a unos S/400 000 al mes y hoy estamos cerca de S/230 000.'),

  H2('3.1. Matriz FODA cruzada'),
  P('El FODA cruzado combina los factores internos y externos de la Tabla 7. Las estrategias FO usan las fortalezas para aprovechar las oportunidades. Las DO corrigen debilidades para aprovechar oportunidades. Las FA usan las fortalezas para enfrentar las amenazas y las DA reducen debilidades para evitar las amenazas (David, 2017). Debajo de cada estrategia pongo los códigos de los factores que la sostienen.'),
  ...T('Matriz FODA cruzada de MatLab Perú', [1900, 3563, 3563], ['', 'FORTALEZAS F1 a F6', 'DEBILIDADES D1 a D6'], [
    [{ t: 'OPORTUNIDADES O1 a O5', bold: true, fill: GRIS, align: AlignmentType.CENTER },
      cuadr('FO. Explotar', [
        ['E1', 'Crear programas de nivelación en matemática por grado escolar, con metas claras y reportes a los padres.', 'F1, F4, O2'],
        ['E2', 'Captar más alumnos con campañas en redes y un examen de diagnóstico gratuito en la web, hecho con inteligencia artificial, que termine en una clase de prueba.', 'F3, F6, O1, O5'],
        ['E3', 'Llegar a familias de otras ciudades del país con campañas por región.', 'F2, F3, O1, O2']]),
      cuadr('DO. Buscar', [
        ['E4', 'Reordenar el área comercial con un solo proceso de ventas en el CRM, metas por asesor y seguimiento a cada interesado.', 'D4, D5, O5'],
        ['E5', 'Medir cada mes la puntualidad, la asistencia y el uso de cámara de cada profesor en un tablero, y dar un bono a quienes cumplen.', 'D2, D3, D6, O5'],
        ['E6', 'Hacer alianzas con colegios y empresas para dar refuerzo en matemática y educación financiera a sus alumnos o trabajadores.', 'D5, O2, O4']])],
    [{ t: 'AMENAZAS A1 a A6', bold: true, fill: GRIS, align: AlignmentType.CENTER },
      cuadr('FA. Confrontar', [
        ['E7', 'Destacar en la venta lo que una plataforma barata o la inteligencia artificial no dan: el profesor fijo, el asesor académico y el reporte a los padres.', 'F1, F4, A5'],
        ['E8', 'Crear una línea de carrera para los profesores, con más horas y mejor pago para los que tienen buenas evaluaciones.', 'F2, F6, A6']]),
      cuadr('DA. Evitar', [
        ['E9', 'Armar un grupo de profesores suplentes ya capacitados para cubrir las faltas el mismo día.', 'D1, D2, A6'],
        ['E10', 'Ofrecer clases en grupos de tres o cuatro alumnos, a menor precio por alumno, sin perder el seguimiento.', 'D5, A1, A5']])],
  ], 'Elaboración propia. Los códigos de los factores corresponden a la Tabla 7.', { size: 18 }),
  P('En la Tabla 9 explico de qué tipo es cada estrategia y cómo ayuda a llegar al objetivo.'),
  ...T('Sustento de las estrategias del FODA cruzado', [700, 1900, 6426], ['Cód.', 'Tipo de estrategia', 'Cómo ayuda al objetivo'], [
    ['E1', 'Desarrollo de productos', 'Solo el 12.8 % de los alumnos de quinto de secundaria logra el nivel esperado en matemática (Ministerio de Educación, 2026). Por mi experiencia, en octubre y noviembre los padres buscan sobre todo que sus hijos no desaprueben. Un programa por grado, con metas que los padres entienden, es más fácil de vender que horas sueltas y trae alumnos nuevos.'],
    ['E2', 'Penetración de mercado', 'El 96 % de los hogares tiene internet (Forbes Perú, 2026). La empresa ya probó un examen de diagnóstico con inteligencia artificial. Terminarlo y usarlo como puerta de entrada convierte más visitas en clases de prueba.'],
    ['E3', 'Desarrollo de mercados', 'Por mi experiencia, cerca del 90 % de los alumnos son de Lima. Las clases son virtuales, así que no hace falta abrir sedes para llegar a familias de otras ciudades.'],
    ['E4', 'Interna, reingeniería de procesos', 'Por mi experiencia, se pierden interesados porque nadie les hace seguimiento. Con un solo proceso y metas por asesor, más consultas terminan en inscripción.'],
    ['E5', 'Interna, calidad total', 'Ataca de forma directa la segunda parte del objetivo. Los datos de cada clase ya se guardan, solo falta convertirlos en indicadores por profesor.'],
    ['E6', 'Alianza estratégica', 'Una alianza con un colegio o una empresa trae varios alumnos a la vez. Además reactiva la línea para empresas, que hoy está en pausa por motivos externos a la empresa.'],
    ['E7', 'Diferenciación', 'Ninguna plataforma ni herramienta de inteligencia artificial da un asesor que siga al alumno y un reporte a los padres después de cada clase. Mostrarlo en la venta ayuda a cerrar más inscripciones.'],
    ['E8', 'Interna, recursos humanos', 'Si los buenos profesores tienen razones para quedarse, no se van a otras academias y el servicio se mantiene estable.'],
    ['E9', 'Interna, recursos humanos', 'Si un profesor falta, la clase igual se dicta. Así el cliente no ve el problema y no se retira.'],
    ['E10', 'Desarrollo de productos', 'Da una opción más barata a familias que hoy eligen tutores independientes o plataformas, y suma alumnos sin bajar la calidad.'],
  ], 'Elaboración propia.', { size: 18, center: [0] }),

  H2('3.2. Matriz Interna-Externa MIE'),
  P('La MIE ubica a la empresa en una de nueve celdas. En el eje X va el total de la MEFI y en el eje Y el total de la MEFE (David, 2017). MatLab Perú tiene 2.40 en la MEFI y 2.34 en la MEFE, así que cae en la celda V.'),
  ...(() => {
    const V = '#FFF2CC', C = 'E2EFDA', R = 'F8CBAD', Y = 'FFF2CC';
    const c = (t, fill) => ({ t, fill, align: AlignmentType.CENTER });
    return T('Matriz Interna-Externa de MatLab Perú', [2426, 2200, 2200, 2200],
      [['Total MEFE / Total MEFI', 'Fuerte 3.0 a 4.0', 'Promedio 2.0 a 2.99', 'Débil 1.0 a 1.99']], [
        [{ t: 'Alto 3.0 a 4.0', bold: true }, c('I', C), c('II', C), c('III', Y)],
        [{ t: 'Medio 2.0 a 2.99', bold: true }, c('IV', C), { t: ['**V**', '**MatLab Perú**', 'MEFI 2.40 y MEFE 2.34'], fill: Y, align: AlignmentType.CENTER }, c('VI', R)],
        [{ t: 'Bajo 1.0 a 1.99', bold: true }, c('VII', Y), c('VIII', R), c('IX', R)],
      ], 'Elaboración propia. Las celdas I, II y IV, en verde, son de crecer y construir. Las celdas III, V y VII, en amarillo, son de retener y mantener. Las celdas VI, VIII y IX, en rojo, son de cosechar o reducir.');
  })(),
  P('La celda V está en la zona de retener y mantener. Para esta zona se recomiendan la penetración de mercado y el desarrollo de productos (David, 2017). Tiene sentido para MatLab. La empresa no está en una posición para arriesgarse con negocios nuevos, pero sí puede vender más de lo que ya hace bien y mejorar su oferta.'),
  P('Esto coincide con varias estrategias del FODA cruzado. E2, E4 y E7 buscan más alumnos en el mercado actual, así que son de penetración de mercado. E1 y E10 son de desarrollo de productos, porque crean formas nuevas de dar el mismo servicio. Estas cinco estrategias salen a la vez del FODA cruzado y de la MIE, por eso pesan más en la Matriz de Decisión.'),

  H2('3.3. Matriz de posición estratégica y evaluación de acción MPEYEA'),
  P('La PEYEA mide la posición de la empresa con cuatro dimensiones. Dos son internas, la fuerza financiera y la ventaja competitiva. Dos son externas, la estabilidad del entorno y la fuerza de la industria (David, 2017). La fuerza financiera y la fuerza de la industria se califican de +1 a +6, donde +6 es lo mejor. La ventaja competitiva y la estabilidad del entorno se califican de −1 a −6, donde −1 es lo mejor y −6 lo peor.'),
  P('Las variables salen de la MEFE, la MEFI y la MPC. Las de fuerza financiera salen de lo que conozco de la empresa por mi trabajo. Uso cifras aproximadas, porque los montos exactos son reservados.'),
  ...T('Calificación de las variables de la PEYEA', [1500, 2500, 700, 4326], ['Dimensión', 'Variable', 'Valor', 'Por qué ese valor'], fixSpan([
    [SEC('Fuerza financiera FF')],
    ['', 'Rentabilidad', '+4', 'La empresa tiene utilidades, aunque los ingresos bajaron de unos S/400 000 al mes entre 2020 y 2022 a unos S/230 000 hoy.'],
    ['', 'Nivel de deuda', '+6', 'No tiene préstamos. Solo hubo una inversión de capital al inicio.'],
    ['', 'Cumplimiento de pagos', '+5', 'Siempre paga a tiempo a profesores y proveedores.'],
    ['', 'Dinero ahorrado', '+4', 'Tiene un ahorro cercano a S/100 000 para invertir en cambios.'],
    ['', 'Estabilidad de los ingresos', '+3', 'Cada mes se mueven entre S/180 000 y S/300 000, pero vienen bajando.'],
    [{ t: 'Promedio FF', bold: true }, '', { t: '+4.40', bold: true }, ''],
    [SEC('Ventaja competitiva VC')],
    ['', 'Calidad y personalización de la enseñanza', '−1', 'Es nuestra mayor fortaleza, F1, y en la MPC tiene 4.'],
    ['', 'Seguimiento a los padres', '−2', 'Cada clase tiene informe y hay asesores académicos que siguen al alumno, F4.'],
    ['', 'Trayectoria y marca', '−3', 'Más de diez años en el mercado, F2, pero la marca no es la más conocida del sector.'],
    ['', 'Puntualidad y cumplimiento', '−6', 'Es el factor más bajo de la MPC, con 1, por las faltas y tardanzas, D2.'],
    ['', 'Precio', '−4', 'Plataformas y tutores independientes cobran menos. En la MPC tiene 2.'],
    ['', 'Captación comercial', '−4', 'Desorden en el área comercial y falta de inscripciones, D4 y D5.'],
    [{ t: 'Promedio VC', bold: true }, '', { t: '−3.33', bold: true }, ''],
    [SEC('Estabilidad del entorno EE')],
    ['', 'Cambios de gobierno y de normas', '−4', 'Pueden cambiar las reglas del sector educativo, A3.'],
    ['', 'Variación de precios y de la economía familiar', '−4', 'Si suben los precios, las familias gastan menos en clases, A1 y O3.'],
    ['', 'Fenómenos climáticos', '−3', 'Un corte de luz o internet afecta la clase virtual, A2.'],
    ['', 'Avance de la inteligencia artificial y de sustitutos gratuitos', '−5', 'Cambia rápido la forma en que los alumnos resuelven dudas, A5.'],
    ['', 'Normas de protección de datos', '−2', 'Son reglas conocidas que la empresa puede cumplir, A4.'],
    ['', 'Necesidad de refuerzo escolar', '−2', 'El bajo logro en matemática se mantiene en el tiempo, O2.'],
    [{ t: 'Promedio EE', bold: true }, '', { t: '−3.33', bold: true }, ''],
    [SEC('Fuerza de la industria FI')],
    ['', 'Conectividad de los hogares', '+6', 'El 96 % de los hogares tiene internet (Forbes Perú, 2026), O1.'],
    ['', 'Demanda de refuerzo en matemática', '+5', 'Solo el 12.8 % de los alumnos de quinto de secundaria logra el nivel esperado (Ministerio de Educación, 2026), O2.'],
    ['', 'Uso de tecnología para personalizar', '+4', 'Las herramientas digitales ayudan a dar mejor seguimiento, O5.'],
    ['', 'Facilidad de entrada de competidores', '+2', 'Empezar cuesta poco, según el análisis de Porter.'],
    ['', 'Rivalidad y presión en precios', '+2', 'Hay muchas academias, tutores y plataformas.'],
    ['', 'Poder de los profesores', '+2', 'Los buenos docentes pueden irse a otras academias, A6.'],
    [{ t: 'Promedio FI', bold: true }, '', { t: '+3.50', bold: true }, ''],
  ], 4), 'Elaboración propia. Los códigos corresponden a la MEFE, la MEFI y la MPC.', { size: 18, center: [2] }),
  P('Con los promedios se calcula el punto del vector. En el eje X se suman la ventaja competitiva y la fuerza de la industria. En el eje Y se suman la fuerza financiera y la estabilidad del entorno.'),
  ...T('Resultado de la PEYEA', [1500, 4526, 3000], ['Eje', 'Cálculo', 'Resultado'], [
    ['X', 'VC + FI = −3.33 + 3.50', { t: '+0.17', bold: true }],
    ['Y', 'FF + EE = 4.40 − 3.33', { t: '+1.07', bold: true }],
  ], 'Elaboración propia.', { center: [0, 1, 2] }),
  ...FIGURA('Vector de la PEYEA de MatLab Perú', 'peyea.png', 400, 400, 'Elaboración propia.'),
  P('El vector cae en el cuadrante agresivo, pero muy cerca del eje vertical. Quiere decir que MatLab tiene dinero para invertir y está en un sector con demanda, pero su ventaja competitiva es débil. Lo que más la debilita es la puntualidad y el cumplimiento de los profesores, que califiqué con −6.'),
  P('En el cuadrante agresivo la empresa puede usar sus fortalezas para aprovechar oportunidades, superar sus debilidades y evitar amenazas. Para este cuadrante se recomiendan la penetración de mercado, el desarrollo de mercados y el desarrollo de productos, entre otras (David, 2017). Como el vector está tan cerca del eje, no conviene arriesgarse con negocios nuevos. Conviene usar el ahorro para crecer en lo que ya hace bien y, al mismo tiempo, corregir lo que la debilita. De la PEYEA salen seis estrategias.'),
  ...T('Estrategias que salen de la PEYEA', [700, 3000, 5326], ['Cód.', 'Estrategia', 'Por qué la recomienda la PEYEA'], [
    ['E1', 'Programas de nivelación por grado escolar', 'Es desarrollo de productos. La fuerza financiera permite diseñar y lanzar el programa sin endeudarse.'],
    ['E2', 'Campañas en redes con examen de diagnóstico gratuito', 'Es penetración de mercado. Aprovecha la alta conectividad, que es la variable más fuerte de la industria.'],
    ['E3', 'Campañas para familias de otras ciudades', 'Es desarrollo de mercados. Las clases virtuales permiten crecer fuera de Lima con poca inversión.'],
    ['E4', 'Reordenar el área comercial en el CRM', 'Es penetración de mercado y corrige la captación comercial, una de las variables más bajas de la ventaja competitiva.'],
    ['E5', 'Medir y premiar el cumplimiento de los profesores', 'La puntualidad es la peor variable de la ventaja competitiva. Mejorarla mueve el vector hacia la derecha y aleja a la empresa del cuadrante conservador.'],
    ['E10', 'Clases en grupos de tres o cuatro alumnos', 'Es desarrollo de productos y responde a la presión en precios de la industria.'],
  ], 'Elaboración propia.', { size: 18, center: [0] }),

  H2('3.4. Matriz Boston Consulting Group MBCG'),
  P('La BCG ordena las líneas de negocio según dos cosas. La primera es cuánto crece su mercado. La segunda es qué parte del mercado tiene la empresa frente a su competidor más fuerte (David, 2017). Con eso cada línea cae en uno de cuatro cuadrantes: estrella, interrogante, vaca o perro.'),
  P('MatLab tiene tres líneas. No hay datos públicos del tamaño de este mercado, así que la participación y el crecimiento son estimaciones mías, a partir de lo que veo en la empresa y de los resultados del diagnóstico.'),
  ...T('Líneas de negocio de MatLab Perú en la BCG', [1800, 1000, 2600, 2326, 1300], ['Línea', 'Parte de los ingresos', 'Participación frente a la competencia', 'Crecimiento del mercado', 'Cuadrante'], [
    ['Clases individuales por paquetes de horas', '60 %', 'Alta. En clases individuales con asesor académico y jefe del área académica no conozco otra academia que ofrezca lo mismo.', 'Bajo. Hay mucha competencia y nuestras ventas vienen bajando.', { t: 'Vaca', bold: true, align: AlignmentType.CENTER }],
    ['Programas de verano', '30 %', 'Alta. Tienen muchos alumnos y una metodología dinámica.', 'Bajo. Las ventas también vienen bajando.', { t: 'Vaca', bold: true, align: AlignmentType.CENTER }],
    ['Talleres en grupo y programas para empresas', '10 %', 'Baja. La línea para empresas está en pausa por motivos externos.', 'Alto. Hay empresas e instituciones que buscan educación financiera y apoyo educativo, O4.', { t: 'Interrogante', bold: true, align: AlignmentType.CENTER }],
  ], 'Elaboración propia. Los porcentajes son aproximados y salen de mi experiencia en la empresa.', { size: 18, center: [1] }),
  ...FIGURA('Matriz BCG de MatLab Perú', 'bcg.png', 420, 364, 'Elaboración propia. El tamaño de cada círculo muestra su parte de los ingresos.'),
  P('Las dos líneas más grandes son vacas. Dan el 90 % de los ingresos y son la base de la empresa, pero ya no crecen. Una vaca hay que cuidarla, porque financia al resto. Por eso las estrategias para ellas buscan vender más en el mismo mercado y no perder calidad. La línea de talleres y empresas es un interrogante. Tiene un mercado con interés, pero hoy MatLab casi no participa. Conviene invertir en ella con cuidado, usando lo que dejan las vacas, para ver si puede convertirse en estrella.'),
  ...T('Estrategias que salen de la BCG', [700, 2700, 1600, 4026], ['Cód.', 'Estrategia', 'Línea', 'Por qué'], [
    ['E1', 'Programas de nivelación por grado escolar', 'Clases individuales', 'Da una forma nueva de vender la vaca en octubre y noviembre, cuando los padres buscan que sus hijos no desaprueben.'],
    ['E2', 'Campañas en redes con examen de diagnóstico gratuito', 'Clases individuales', 'Trae más alumnos a la línea que más ingresos deja.'],
    ['E5', 'Medir y premiar el cumplimiento de los profesores', 'Clases individuales', 'La vaca depende de que el profesor cumpla. Si falla, el cliente se va y la línea pierde su ventaja.'],
    ['E7', 'Destacar el profesor fijo, el asesor académico y el reporte a los padres', 'Clases individuales', 'Defiende la alta participación frente a plataformas baratas y la inteligencia artificial.'],
    ['E10', 'Clases en grupos de tres o cuatro alumnos', 'Programas de verano', 'La metodología dinámica del verano ya funciona en grupo. Llevarla a grupos chicos todo el año suma alumnos a menor precio.'],
    ['E6', 'Alianzas con colegios y empresas', 'Talleres y empresas', 'Es la forma de invertir en el interrogante con poco riesgo, porque cada alianza trae varios alumnos a la vez.'],
  ], 'Elaboración propia.', { size: 18, center: [0] }),

  H2('3.5. Matriz de Decisión Estratégica MDE'),
  P('La Matriz de Decisión junta las estrategias que salieron de las matrices anteriores y cuenta cuántas veces aparece cada una. Las que se repiten más tienen más respaldo, porque varias herramientas llegan a la misma conclusión (David, 2017). Me quedo con las que aparecen tres veces o más. Las demás quedan como estrategias de contingencia, para usarlas si las primeras no funcionan o si cambia el entorno.'),
  ...(() => {
    const X = 'X', o = '';
    const r = (c, e, t, a, b, d, f, n) => [c, e, t, a, b, d, f, n >= 3 ? { t: String(n), bold: true } : String(n)];
    return T('Matriz de Decisión Estratégica de MatLab Perú', [600, 3206, 1800, 760, 600, 800, 620, 640],
      ['Cód.', 'Estrategia', 'Tipo', 'FODA', 'MIE', 'PEYEA', 'BCG', 'Total'], [
        r('E1', 'Crear programas de nivelación en matemática por grado escolar', 'Desarrollo de productos', X, X, X, X, 4),
        r('E2', 'Captar alumnos con campañas en redes y un examen de diagnóstico gratuito', 'Penetración de mercado', X, X, X, X, 4),
        r('E10', 'Ofrecer clases en grupos de tres o cuatro alumnos', 'Desarrollo de productos', X, X, X, X, 4),
        r('E4', 'Reordenar el área comercial en el CRM', 'Penetración de mercado', X, X, X, o, 3),
        r('E5', 'Medir y premiar el cumplimiento de los profesores', 'Interna, calidad total', X, o, X, X, 3),
        r('E7', 'Destacar el profesor fijo, el asesor académico y el reporte a los padres', 'Diferenciación', X, X, o, X, 3),
        r('E3', 'Llegar a familias de otras ciudades', 'Desarrollo de mercados', X, o, X, o, 2),
        r('E6', 'Hacer alianzas con colegios y empresas', 'Alianza estratégica', X, o, o, X, 2),
        r('E8', 'Crear una línea de carrera para los profesores', 'Interna, recursos humanos', X, o, o, o, 1),
        r('E9', 'Armar un grupo de profesores suplentes', 'Interna, recursos humanos', X, o, o, o, 1),
      ], 'Elaboración propia. Se retienen las estrategias con tres repeticiones o más.', { size: 18, center: [0, 3, 4, 5, 6, 7] });
  })(),
  P('Quedan seis estrategias retenidas: E1, E2, E10, E4, E5 y E7. Las seis van con el objetivo. E1, E2, E4, E7 y E10 apuntan a tener más alumnos, y E5 apunta a que los profesores cumplan.'),
  P('E3, E6, E8 y E9 quedan como contingencia. E3 es buena idea, pero hoy cerca del 90 % de los alumnos son de Lima y primero conviene crecer donde la marca ya se conoce. E6 depende de reactivar la línea para empresas, que está en pausa. E8 y E9 ayudan al cumplimiento, pero E5 ya lo ataca de forma más directa. Si E5 no basta, E9 sería la primera en activarse.'),

  H2('3.6. Matriz cuantitativa del planeamiento estratégico MCPE'),
  P('La MCPE califica qué tan atractiva es cada estrategia frente a cada factor de la MEFE y la MEFI, con los mismos pesos (David, 2017). Por eso la suma de pesos es 2.00. Para cada factor me pregunté si cambia la elección entre estrategias, y le puse a cada una un puntaje de atractivo, PA, de 1 a 4. El 1 es nada atractiva, el 2 algo atractiva, el 3 atractiva y el 4 muy atractiva. Luego multipliqué el peso por el PA y obtuve el TPA. El total puede ir de 2 a 8, y un total de 5 o más indica una estrategia aceptable. Los puntajes los puse yo, a partir del diagnóstico y de lo que conozco de la empresa.'),
  ...(() => {
    const G = { O: 'Oportunidades', A: 'Amenazas', F: 'Fortalezas', D: 'Debilidades' };
    const rows = []; let last = '';
    MCPE.forEach(([c, n, w, pa]) => {
      if (c[0] !== last) { rows.push([{ t: G[c[0]], span: 14, fill: GRIS2, bold: true }]); last = c[0]; }
      rows.push([`${c}. ${n}`, w.toFixed(2), ...pa.flatMap((p) => [String(p), (w * p).toFixed(2)])]);
    });
    rows.push([{ t: 'Total', bold: true }, { t: '2.00', bold: true }, ...totMCPE.flatMap((t) => ['', { t: t.toFixed(2), bold: true }])]);
    const wd = [2186, 520, ...Array(12).fill(570)];
    return T('Matriz cuantitativa del planeamiento estratégico de MatLab Perú', wd,
      [['Factor clave', 'Peso', ...E_MCPE.map((e) => ({ t: e, span: 2 }))], ['', '', ...E_MCPE.flatMap(() => ['PA', 'TPA'])]],
      rows, 'Elaboración propia. PA es el puntaje de atractivo y TPA es el peso por el PA. E1 es nivelación por grado, E2 campañas con diagnóstico gratuito, E4 orden del área comercial, E5 cumplimiento de los profesores, E7 diferenciación en la venta y E10 grupos de tres o cuatro alumnos.',
      { size: 16, center: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13] });
  })(),
  ...T('Resultado de la MCPE', [800, 700, 4626, 900, 2000], ['Puesto', 'Cód.', 'Estrategia', 'Total', 'Decisión'], [
    ['1', 'E5', 'Medir y premiar el cumplimiento de los profesores', '5.48', 'Pasa a la Matriz de Ética'],
    ['2', 'E10', 'Ofrecer clases en grupos de tres o cuatro alumnos', '5.40', 'Pasa a la Matriz de Ética'],
    ['3', 'E1', 'Crear programas de nivelación por grado escolar', '5.29', 'Pasa a la Matriz de Ética'],
    ['4', 'E2', 'Captar alumnos con campañas y examen de diagnóstico gratuito', '4.94', 'Pasa a la Matriz de Ética'],
    ['5', 'E7', 'Destacar el profesor fijo, el asesor académico y el reporte a los padres', '4.73', 'Contingencia'],
    ['6', 'E4', 'Reordenar el área comercial en el CRM', '4.21', 'Contingencia'],
  ], 'Elaboración propia.', { size: 18, center: [0, 1, 3] }),
  P('E5 sale primera. Es la única que ataca de frente las debilidades más pesadas, la dependencia de los profesores, sus faltas y tardanzas y las malas prácticas. También responde a la disputa por docentes, una de las amenazas con más peso. Tiene sentido que gane, porque si el profesor no cumple, ninguna otra estrategia funciona.'),
  P('E10 sale segunda porque responde bien a la competencia de bajo costo y al alza de precios. Da una opción más barata a las familias sin dejar el seguimiento.'),
  P('E1 sale tercera. Aprovecha la oportunidad con más peso, el bajo logro en matemática, y usa las dos fortalezas principales, la personalización y el seguimiento a los padres.'),
  P('E2 queda cuarta con 4.94, apenas debajo de 5. La mantengo porque es la que más ayuda con la falta de inscripciones y aprovecha la conectividad y la inteligencia artificial.'),
  P('E7 y E4 quedan como contingencia. E4 sacó el puntaje más bajo porque ordena las ventas, pero no ayuda con los problemas de los profesores, que son los que más pesan en la matriz. Aun así, el área comercial necesita orden, y E4 se puede aplicar después junto con E2.'),

  H2('3.7. Matriz ética ME'),
  P('La Matriz de Ética es el último filtro. Revisa que las estrategias no violen derechos, que sean justas y que den buenos resultados con medios correctos. Si una estrategia viola un derecho, es injusta o es perjudicial, se descarta. Evalué las cuatro estrategias con mayor puntaje en la MCPE.'),
  ...(() => {
    const N = 'Neutral', Pm = 'Promueve', J = 'Justo', Ex = 'Excelente';
    return T('Matriz de Ética de MatLab Perú', [1300, 3326, 1100, 1100, 1100, 1100], ['Criterio', 'Impacto en', 'E5', 'E10', 'E1', 'E2'], fixSpan([
      [SEC('Derechos')],
      ['', '1. El derecho a la vida', N, N, N, N],
      ['', '2. El derecho a la propiedad', N, N, N, N],
      ['', '3. El derecho al libre pensamiento', N, N, N, N],
      ['', '4. El derecho a la privacidad', N, N, N, N],
      ['', '5. El derecho a la libertad de conciencia', N, N, N, N],
      ['', '6. El derecho a hablar libremente', N, N, N, N],
      ['', '7. El derecho al debido proceso', Pm, N, N, N],
      [SEC('Justicia')],
      ['', '8. La distribución', J, J, J, N],
      ['', '9. La equidad en la administración', J, N, N, N],
      ['', '10. Las normas de compensación', J, N, N, N],
      [SEC('Utilitarismo')],
      ['', '11. Los fines y resultados estratégicos', Ex, Ex, Ex, Ex],
      ['', '12. Los medios estratégicos empleados', Ex, Ex, Ex, Ex],
      [{ t: 'Decisión', bold: true }, '', { t: 'Se acepta', bold: true }, { t: 'Se acepta', bold: true }, { t: 'Se acepta', bold: true }, { t: 'Se acepta', bold: true }],
    ], 6), 'Elaboración propia. En derechos se marca promueve, neutral o viola. En justicia, justo, neutral o injusto. En utilitarismo, excelente, neutral o perjudicial.', { size: 18, center: [2, 3, 4, 5] });
  })(),
  P('Ninguna estrategia viola derechos ni es injusta, así que las cuatro se aceptan. E5 promueve el debido proceso porque las reglas serán claras y conocidas por todos los profesores, y cada uno podrá ver sus propios resultados y explicar una falta antes de cualquier decisión. Solo se miden datos de trabajo, como la hora de ingreso y si la cámara estuvo encendida, así que no afecta la privacidad. También es justa en la distribución, en la administración y en la compensación, porque el bono se da con la misma regla para todos. Esto va con el compromiso 4 del código de ética, cumplir los horarios y las clases programadas.'),
  P('E2 necesita un cuidado especial con la privacidad. El examen de diagnóstico recoge datos de alumnos, y muchos son menores de edad. Por eso solo se aplicará con autorización de los padres y cumpliendo la ley de protección de datos, como dice el compromiso 2 del código de ética. Con esa condición, el impacto en la privacidad es neutral. La publicidad también debe mostrar precios y resultados reales, según el compromiso 3.'),
  P('E1 y E10 son justas en la distribución porque acercan el refuerzo en matemática a más familias. E10 en especial baja el precio por alumno, así que llega a familias que hoy no pueden pagar una clase individual.'),

  H2('3.8. Conclusiones'),
  P('El FODA cruzado dio diez estrategias. La MIE ubicó a MatLab en la celda V, de retener y mantener, y la PEYEA en el cuadrante agresivo, aunque muy cerca del eje. La BCG mostró que el 90 % de los ingresos viene de dos líneas que ya no crecen, las clases individuales y los programas de verano.'),
  P('La Matriz de Decisión retuvo seis estrategias y la MCPE eligió cuatro. Las cuatro pasaron la Matriz de Ética. Son medir y premiar el cumplimiento de los profesores, ofrecer clases en grupos de tres o cuatro alumnos, crear programas de nivelación por grado escolar y captar alumnos con campañas en redes y un examen de diagnóstico gratuito.'),
  P('Juntas apuntan al objetivo de volver a unos 500 alumnos en octubre y noviembre de 2028. E5 cuida que los profesores cumplan, que es la base de todo. E1 y E2 traen alumnos nuevos en los meses de más demanda, y E10 da una opción más barata frente a la competencia.'),
  P('Según los resultados, lo primero que debería hacer la empresa es aplicar E5, porque las demás estrategias solo funcionan si el profesor cumple. Las estrategias que no pasaron quedan como contingencia, sobre todo el orden del área comercial y el grupo de profesores suplentes.'),
];

const REF = (parts) => new Paragraph({ children: parts.flatMap((p) => runs(p)), indent: { left: 720, hanging: 720 }, spacing: { line: 480 } });
const refs = [
  BR(),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 480 }, children: [new TextRun({ text: 'Referencias', bold: true })] }),
  REF(['David, F. R. (2017). *Conceptos de administración estratégica* (15.ª ed.). Pearson Educación.']),
  REF(['Forbes Perú. (2026, 20 de mayo). *Hogares peruanos con acceso a internet fijo o móvil se elevó a 96 % en el 2025, según Erestel*. https://forbes.pe/actualidad/2026-05-20/hogares-peruanos-con-acceso-a-internet-fijo-o-movil-se-elevo-a-96-en-el-2025-segun-erestel/']),
  REF(['MatLab Perú. (s. f.). *Inicio* [Sitio web]. https://matlab.com.pe/']),
  REF(['Ministerio de Educación. (2026). *ENLA 2025: resumen ejecutivo*. Oficina de Medición de la Calidad de los Aprendizajes. https://repositorio.minedu.gob.pe/handle/20.500.12799/12773']),
];

const hstyle = (id, name, o) => ({
  id, name, basedOn: 'Normal', next: 'Normal', quickFormat: true,
  run: { font: FONT, size: 24, bold: true, italics: !!o.italics, color: '000000' },
  paragraph: { alignment: o.center ? AlignmentType.CENTER : AlignmentType.LEFT, spacing: { before: o.before || 240, after: 0, line: 480 }, outlineLevel: o.level, keepNext: true },
});

const doc = new Document({
  creator: 'Pierre Gallegos Mallqui',
  title: 'Avance de Proyecto Final 2. Diseño del plan estratégico de MatLab Perú',
  styles: {
    default: { document: { run: { font: FONT, size: 24 } } },
    paragraphStyles: [
      hstyle('Heading1', 'Heading 1', { center: true, level: 0, before: 360 }),
      hstyle('Heading2', 'Heading 2', { level: 1 }),
      hstyle('Heading3', 'Heading 3', { italics: true, level: 2 }),
    ],
  },
  numbering: { config: [{ reference: 'etica', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 360 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: [PageNumber.CURRENT] })] })] }) },
    children: [...portada, ...sec1, ...sec2, ...sec3, ...refs],
  }],
});

const out = path.join(DIR, 'APF2_MatLab_Peru_Gallegos.docx');
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(out, b); console.log('ok', out, 'tablas', nTabla, 'figuras', nFig); });
