/* =========================================================
   PLANTILLA RESTAURANTE - CONFIGURACIÓN DESDE EXCEL
   =========================================================

   Esta versión está preparada para el nuevo archivo:

   Datos_Restaurante.xlsx

   Hojas principales:
   - NAV
   - INICIO
   - MENU
   - BENEFICIOS
   - GALERIA
   - FAQ
   - CONTACTO
   - FOOTER
   - COLORES

   ========================================================= */


/* =========================================================
   CONFIGURACIÓN PREDETERMINADA
   ========================================================= */

const FALLBACK = {

  nav: {
    logo: 'logo.png',
    nombre: 'Nombre del negocio',
    nombreColor: '#FFFFFF',
    opcionesColor: '#FFFFFF',
    fondoColor: '#000000'
  },

  inicio: {
    fondoTipo: 'Imagen',
    fondoImagen: 'portada.jpg',
    fondoColor: '#ffffff',
    titulo: 'Título principal',
    tituloColor: '#000000',
    propuesta: 'Propuesta de valor del negocio',
    propuestaColor: '#000000',
    ctaTexto: 'Llamada a la acción',
    ctaColor: '#ff0000',
    ctaTextoColor: '#FFFFFF',
    ctaTipo: 'WhatsApp',
    ctaUrl: ''
  },

  menu: {
    fondoTipo: 'Color',
    fondoImagen: 'menu.jpg',
    fondoColor: '#000000',
    etiqueta: 'NUESTRO MENÚ',
    etiquetaColor: '#000000',
    titulo: 'Descripción resumida del menú',
    tituloColor: '#000000',
    descripcion: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
    descripcionColor: '#000000',
    cartaColor: '#ffffff',
    botonTexto: 'Ver menú',
    botonColor: '#ff0000',
    botonTextoColor: '#FFFFFF',
    pdfUrl: 'menu.pdf'
  },

  beneficios: {
    fondoTipo: 'Color',
    fondoImagen: 'beneficios.jpg',
    fondoColor: '#ffffff',
    etiqueta: 'POR QUÉ ELEGIRNOS',
    etiquetaColor: '#000000',
    titulo: 'Beneficios principales',
    tituloColor: '#000000',
    descripcion: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
    descripcionColor: '#000000',

    tarjetas: [
      {
        titulo: 'Beneficio 1',
        tituloColor: '#ffffff',
        descripcion: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
        descripcionColor: '#ffffff',
        icono: '✅',
        iconoFondoColor: '#ffffff',
        cartaColor: '#000000'
      },
      {
        titulo: 'Beneficio 2',
        tituloColor: '#ffffff',
        descripcion: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
        descripcionColor: '#ffffff',
        icono: '✅',
        iconoFondoColor: '#ffffff',
        cartaColor: '#000000'
      },
      {
        titulo: 'Beneficio 3',
        tituloColor: '#ffffff',
        descripcion: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
        descripcionColor: '#ffffff',
        icono: '✅',
        iconoFondoColor: '#ffffff',
        cartaColor: '#000000'
      },
      {
        titulo: 'Beneficio 4',
        tituloColor: '#ffffff',
        descripcion: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
        descripcionColor: '#ffffff',
        icono: '✅',
        iconoFondoColor: '#ffffff',
        cartaColor: '#000000'
      }
    ]
  },

  galeria: {
    fondoTipo: 'Color',
    fondoImagen: 'galeria-fondo.jpg',
    fondoColor: '#000000',
    titulo: 'GALERÍA',
    tituloColor: '#ffffff',
    imagenes: [
      'galeria1.jpg',
      'galeria2.jpg',
      'galeria3.jpg',
      'galeria4.jpg'
    ]
  },

  faq: {
    fondoTipo: 'Color',
    fondoImagen: 'faq.jpg',
    fondoColor: '#ffffff',
    titulo: 'PREGUNTAS FRECUENTES',
    tituloColor: '#000000',
    preguntas: [
      {
        pregunta: 'Pregunta frecuente 1',
        preguntaColor: '#000000',
        respuesta: 'Respuesta de ejemplo.',
        respuestaColor: '#374151',
        cartaColor: '#FFFFFF'
      }
    ]
  },

  contacto: {
    fondoTipo: 'Color',
    fondoImagen: 'contacto.jpg',
    fondoColor: '#000000',
    etiqueta: 'CONTACTO',
    etiquetaColor: '#ffffff',
    titulo: 'Medios de contacto del negocio',
    tituloColor: '#ffffff',
    descripcion: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
    descripcionColor: '#ffffff',

    direccion: 'Calle 123, Colonia Vecindario, 00000, Ciudad, Estado, Mexico',
    direccionColor: '#000000',
    direccionCartaColor: '#FFFFFF',

    dias: 'Lun–Dom',
    diasColor: '#000000',
    horas: '0:00 AM – 11:59 PM',
    horasColor: '#000000',
    horarioCartaColor: '#FFFFFF',

    telefono: '12 3456 7890',
    telefonoColor: '#000000',
    telefonoCartaColor: '#FFFFFF',

    email: 'ejemplocorreo@gmail.com',
    emailColor: '#000000',
    emailCartaColor: '#FFFFFF',

    mapsTexto: 'Abrir en Google Maps',
    mapsBotonColor: '#FFFFFF',
    mapsTextoColor: '#0B2239',

    whatsappTexto: 'Enviar WhatsApp',
    whatsappBotonColor: '#16A34A',
    whatsappTextoColor: '#FFFFFF',
    whatsapp: '52 12 3456 7890',
    whatsappMensaje: 'Quiero reservar una mesa',

    zoomMapa: 16
  },

  footer: {
    fondoColor: '#FFFFFF',
    textoColor: '#000000',
    descripcion: 'Breve descripción del negocio.',
    instagram: 'https://instagram.com/',
    facebook: 'https://facebook.com/',
    copyright: '© Nombre del negocio. Todos los derechos reservados.',
    mostrarCredito: true
  }

};


/* =========================================================
   CONFIGURACIÓN CUANDO EL EXCEL NO ESTÁ CONECTADO
   =========================================================

   Objetivo:
   - No mostrar logo.png en el NAV.
   - No mostrar portada.jpg en INICIO.
   - No mostrar menu.jpg ni ninguna otra imagen de sección.
   - Mantener colores y textos genéricos para que la plantilla
     siga siendo visible antes de conectar el Excel.

   Cuando Datos_Restaurante.xlsx se conecta correctamente,
   applyConfig() reemplaza esta configuración por la del Excel.
   ========================================================= */

const NO_EXCEL_CONFIG = {

  nav: {
    ...FALLBACK.nav,
    logo: ''
  },

  inicio: {
    ...FALLBACK.inicio,
    fondoTipo: 'Color',
    fondoImagen: ''
  },

menu: {
    ...FALLBACK.menu,
    fondoTipo: 'Color',
    fondoImagen: ''
  },

  beneficios: {
    ...FALLBACK.beneficios,
    fondoTipo: 'Color',
    fondoImagen: ''
  },

  galeria: {
    ...FALLBACK.galeria,
    fondoTipo: 'Color',
    fondoImagen: '',
    imagenes: []
  },

  faq: {
    ...FALLBACK.faq,
    fondoTipo: 'Color',
    fondoImagen: ''
  },

  contacto: {
    ...FALLBACK.contacto,
    fondoTipo: 'Color',
    fondoImagen: ''
  },

  footer: {
    ...FALLBACK.footer
  }

};


/* =========================================================
   UTILIDADES DOM
   ========================================================= */

const $ = (id) => document.getElementById(id);

function setText(id, value) {
  const element = $(id);

  if (!element || value === undefined || value === null) {
    return;
  }

  element.textContent = String(value);
}

function setCssVar(name, value) {
  if (!name || !validHex(value)) {
    return;
  }

  document.documentElement.style.setProperty(name, value);
}

function validHex(value) {
  return /^#[0-9a-fA-F]{6}$/.test(String(value || '').trim());
}

function cleanPhone(value) {
  return String(value || '').replace(/\D/g, '');
}

function normalizeText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

function boolFromText(value, fallback = true) {
  const normalized = normalizeText(value);

  if (['si', 'sí', 'yes', 'true', '1'].includes(normalized)) {
    return true;
  }

  if (['no', 'false', '0'].includes(normalized)) {
    return false;
  }

  return fallback;
}

function numberInRange(value, min, max, fallback) {
  const parsed = Number(value);

  if (!Number.isFinite(parsed)) {
    return fallback;
  }

  return Math.min(max, Math.max(min, parsed));
}


/* =========================================================
   URL / WHATSAPP / MAPAS
   ========================================================= */

function buildWhatsappUrl(phone, message = '') {
  const clean = cleanPhone(phone);

  if (!clean) {
    return '#';
  }

  const text = String(message || '').trim();

  return text
    ? `https://wa.me/${clean}?text=${encodeURIComponent(text)}`
    : `https://wa.me/${clean}`;
}

function safeUrl(value, fallback = '#') {
  const raw = String(value || '').trim();

  if (!raw) {
    return fallback;
  }

  if (raw.startsWith('#')) {
    return raw;
  }

  /* Ruta relativa: menu.pdf, archivos/menu.pdf, etc. */
  if (!/^[a-zA-Z][a-zA-Z\d+.-]*:/.test(raw)) {
    return raw;
  }

  try {
    const url = new URL(raw, window.location.href);
    const allowed = ['http:', 'https:', 'mailto:', 'tel:'];

    return allowed.includes(url.protocol)
      ? url.href
      : fallback;
  }
  catch {
    return fallback;
  }
}

function buildGoogleMapsUrl(address) {
  const query = String(address || '').trim();

  if (!query) {
    return '#';
  }

  return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(query)}`;
}

function buildGoogleMapsEmbed(address, zoom = 16) {
  const cleanAddress = String(address || '').trim();
  const safeZoom = numberInRange(zoom, 3, 20, 16);

  const defaultAddress =
    'Calle 123, Colonia Vecindario, 00000, Ciudad, Estado, Mexico';

  if (!cleanAddress || cleanAddress === defaultAddress) {
    return 'https://www.google.com/maps?q=Mexico&z=5&output=embed';
  }

  return (
    `https://www.google.com/maps?q=${encodeURIComponent(cleanAddress)}` +
    `&z=${safeZoom}&output=embed`
  );
}


/* =========================================================
   ARCHIVOS DE IMAGEN
   ========================================================= */

function safeAssetName(value, fallback = 'logo.png') {
  const name = String(value || '').trim();
  const validExtension = /^[\w .()-]+\.(png|jpe?g|webp|svg)$/iu;

  return validExtension.test(name)
    ? name
    : fallback;
}

function getImageMimeType(filename) {
  const extension = String(filename || '')
    .trim()
    .toLowerCase()
    .split('.')
    .pop();

  const mimeTypes = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    webp: 'image/webp',
    svg: 'image/svg+xml'
  };

  return mimeTypes[extension] || 'image/png';
}

function setFavicon(filename) {
  let favicon = $('siteFavicon');

  if (!favicon) {
    favicon = document.createElement('link');
    favicon.id = 'siteFavicon';
    favicon.rel = 'icon';
    document.head.appendChild(favicon);
  }

  const safeFilename = safeAssetName(filename, 'logo.png');

  favicon.type = getImageMimeType(safeFilename);
  favicon.href = `img/${safeFilename}?v=${Date.now()}`;
}

function resetFavicon() {
  let favicon = $('siteFavicon');

  if (!favicon) {
    favicon = document.createElement('link');
    favicon.id = 'siteFavicon';
    favicon.rel = 'icon';
    document.head.appendChild(favicon);
  }

  favicon.type = 'image/svg+xml';
  favicon.href = `img/favicon.svg?v=${Date.now()}`;
}

function applyLogo(filename, businessName) {
  const logo = $('brandLogo');
  const brandMark = $('brandMark');

  const initials = String(businessName || 'R')
    .trim()
    .split(/\s+/)
    .slice(0, 2)
    .map((word) => word[0] || '')
    .join('')
    .toUpperCase();

  if (brandMark) {
    brandMark.textContent = initials || 'R';
  }

  if (!logo) {
    return;
  }

  /*
  Si no hay Excel conectado usamos logo = ''.
  En ese caso NO intentamos cargar img/logo.png.
  Mostramos únicamente las iniciales generadas por CSS/HTML.
  */
  const rawFilename = String(filename || '').trim();

  if (!rawFilename) {
    logo.removeAttribute('src');
    logo.classList.add('is-hidden');
    brandMark?.classList.add('is-visible');
    resetFavicon();
    return;
  }

  const safeFilename = safeAssetName(rawFilename, 'logo.png');
  const src = `img/${safeFilename}?v=${Date.now()}`;

  logo.onload = () => {
    logo.classList.remove('is-hidden');
    brandMark?.classList.remove('is-visible');
    setFavicon(safeFilename);
  };

  logo.onerror = () => {
    logo.classList.add('is-hidden');
    brandMark?.classList.add('is-visible');
    resetFavicon();
  };

  logo.src = src;
}


/* =========================================================
   FONDOS DE SECCIONES
   ========================================================= */

function applySectionBackground({
  selector,
  mode,
  filename,
  fallbackFilename
}) {
  const section = document.querySelector(selector);

  if (!section) {
    return;
  }

  const normalizedMode = normalizeText(mode || 'Color');

  const useImage = [
    'imagen',
    'image',
    'foto'
  ].includes(normalizedMode);


  /* =====================================================
     LIMPIAR EL ESTADO ANTERIOR
     ===================================================== */

  section.classList.remove('has-background-image');

  section.style.removeProperty('background-image');
  section.style.removeProperty('background-color');
  section.style.removeProperty('background-position');
  section.style.removeProperty('background-size');
  section.style.removeProperty('background-repeat');


  /* =====================================================
     MODO COLOR
     ===================================================== */

  if (!useImage) {
    /*
    En modo Color no colocamos estilos inline.

    El color vuelve a depender de las variables CSS:
    --hero-bg
    --menu-bg
    --benefits-bg
    --gallery-bg
    --faq-bg
    --contact-bg

    Esas variables ya son actualizadas por applyTheme().
    */
    return;
  }


  /* =====================================================
     MODO IMAGEN
     ===================================================== */

  const safeFilename = safeAssetName(
    filename,
    fallbackFilename
  );

  if (!safeFilename) {
    console.warn(
      `No se indicó una imagen válida para ${selector}`
    );
    return;
  }

  const src = `img/${safeFilename}?v=${Date.now()}`;

  const image = new Image();


  /* =====================================================
     IMAGEN CARGADA CORRECTAMENTE
     ===================================================== */

  image.onload = () => {
    /*
    Cuando Excel indica "Imagen", hacemos transparente
    el color de respaldo de la propia sección.

    De esta manera el color predeterminado no cubre
    visualmente la fotografía.
    */
    section.style.backgroundColor = 'transparent';

    section.style.backgroundImage = `url("${src}")`;

    section.style.backgroundPosition = 'center';
    section.style.backgroundSize = 'cover';
    section.style.backgroundRepeat = 'no-repeat';

    section.classList.add('has-background-image');
  };


  /* =====================================================
     ERROR AL CARGAR LA IMAGEN
     ===================================================== */

  image.onerror = () => {
    section.classList.remove('has-background-image');

    section.style.removeProperty('background-image');
    section.style.removeProperty('background-color');
    section.style.removeProperty('background-position');
    section.style.removeProperty('background-size');
    section.style.removeProperty('background-repeat');

    console.warn(
      `No se pudo cargar la imagen: ${safeFilename}`
    );
  };


  /* =====================================================
     INICIAR CARGA
     ===================================================== */

  image.src = src;
}


/* =========================================================
   COLORES DEL EXCEL
   ========================================================= */

function buildColorCatalog(workbook) {
  /*
  El Excel actual usa la hoja COLORES.
  También aceptamos CATALOGOS y LISTA_COLORES para mantener
  compatibilidad con versiones anteriores de la plantilla.
  */
  const sheet =
    workbook.Sheets.COLORES ||
    workbook.Sheets.CATALOGOS ||
    workbook.Sheets.LISTA_COLORES;

  const catalog = {};

  if (!sheet) {
    console.warn(
      'No se encontró la hoja COLORES/CATALOGOS/LISTA_COLORES. ' +
      'Se usarán los colores predeterminados.'
    );

    return catalog;
  }

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ''
  });

  rows.slice(1).forEach((row) => {
    const name = String(row[0] || '').trim();
    const hex = String(row[1] || '').trim();

    if (name && validHex(hex)) {
      /* Clave exacta para los valores del desplegable. */
      catalog[name] = hex;

      /* Clave normalizada para tolerar mayúsculas, acentos y espacios. */
      catalog[normalizeText(name)] = hex;
    }
  });

  return catalog;
}

function resolveColor(value, catalog, fallback) {
  const raw = String(value || '').trim();

  if (validHex(raw)) {
    return raw;
  }

  if (catalog[raw]) {
    return catalog[raw];
  }

  const normalized = normalizeText(raw);

  if (catalog[normalized]) {
    return catalog[normalized];
  }

  return fallback;
}

function getCell(workbook, sheetName, address, fallback = '') {
  const sheet = workbook.Sheets[sheetName];

  if (!sheet || !sheet[address]) {
    return fallback;
  }

  const value = sheet[address].v;

  return value === undefined || value === null
    ? fallback
    : value;
}


/* =========================================================
   LEER EL NUEVO EXCEL
   ========================================================= */

function parseNewWorkbook(workbook) {
  const colors = buildColorCatalog(workbook);
  const c = (sheet, address, fallback = '') =>
    getCell(workbook, sheet, address, fallback);
  const color = (sheet, address, fallback) =>
    resolveColor(c(sheet, address), colors, fallback);

  /* NAV */
  const nav = {
    logo: String(c('NAV', 'C2', FALLBACK.nav.logo)).trim(),
    nombre: String(c('NAV', 'C3', FALLBACK.nav.nombre)).trim(),
    nombreColor: color('NAV', 'C4', FALLBACK.nav.nombreColor),
    opcionesColor: color('NAV', 'C5', FALLBACK.nav.opcionesColor),
    fondoColor: color('NAV', 'C6', FALLBACK.nav.fondoColor)
  };

  /* INICIO */
  const inicio = {
    fondoTipo: c('INICIO', 'C2', FALLBACK.inicio.fondoTipo),
    fondoImagen: c('INICIO', 'C3', FALLBACK.inicio.fondoImagen),
    fondoColor: color('INICIO', 'C4', FALLBACK.inicio.fondoColor),
    titulo: c('INICIO', 'C5', FALLBACK.inicio.titulo),
    tituloColor: color('INICIO', 'C6', FALLBACK.inicio.tituloColor),
    propuesta: c('INICIO', 'C7', FALLBACK.inicio.propuesta),
    propuestaColor: color('INICIO', 'C8', FALLBACK.inicio.propuestaColor),
    ctaTexto: c('INICIO', 'C9', FALLBACK.inicio.ctaTexto),
    ctaColor: color('INICIO', 'C10', FALLBACK.inicio.ctaColor),
    ctaTextoColor: color('INICIO', 'C11', FALLBACK.inicio.ctaTextoColor),
    ctaTipo: c('INICIO', 'C12', FALLBACK.inicio.ctaTipo),
    ctaUrl: c('INICIO', 'C13', FALLBACK.inicio.ctaUrl)
  };

  /* MENÚ */
  const menu = {
    fondoTipo: c('MENU', 'C2', FALLBACK.menu.fondoTipo),
    fondoImagen: c('MENU', 'C3', FALLBACK.menu.fondoImagen),
    fondoColor: color('MENU', 'C4', FALLBACK.menu.fondoColor),
    etiqueta: c('MENU', 'C5', FALLBACK.menu.etiqueta),
    etiquetaColor: color('MENU', 'C6', FALLBACK.menu.etiquetaColor),
    titulo: c('MENU', 'C7', FALLBACK.menu.titulo),
    tituloColor: color('MENU', 'C8', FALLBACK.menu.tituloColor),
    descripcion: c('MENU', 'C9', FALLBACK.menu.descripcion),
    descripcionColor: color('MENU', 'C10', FALLBACK.menu.descripcionColor),
    cartaColor: color('MENU', 'C11', FALLBACK.menu.cartaColor),
    botonTexto: c('MENU', 'C12', FALLBACK.menu.botonTexto),
    botonColor: color('MENU', 'C13', FALLBACK.menu.botonColor),
    botonTextoColor: color('MENU', 'C14', FALLBACK.menu.botonTextoColor),
    pdfUrl: c('MENU', 'C15', FALLBACK.menu.pdfUrl)
  };

  /* BENEFICIOS */
  const beneficios = {
    fondoTipo: c('BENEFICIOS', 'C2', FALLBACK.beneficios.fondoTipo),
    fondoImagen: c('BENEFICIOS', 'C3', FALLBACK.beneficios.fondoImagen),
    fondoColor: color('BENEFICIOS', 'C4', FALLBACK.beneficios.fondoColor),
    etiqueta: c('BENEFICIOS', 'C5', FALLBACK.beneficios.etiqueta),
    etiquetaColor: color('BENEFICIOS', 'C6', FALLBACK.beneficios.etiquetaColor),
    titulo: c('BENEFICIOS', 'C7', FALLBACK.beneficios.titulo),
    tituloColor: color('BENEFICIOS', 'C8', FALLBACK.beneficios.tituloColor),
    descripcion: c('BENEFICIOS', 'C9', FALLBACK.beneficios.descripcion),
    descripcionColor: color('BENEFICIOS', 'C10', FALLBACK.beneficios.descripcionColor),
    tarjetas: []
  };

  const benefitSheet = workbook.Sheets.BENEFICIOS;

  if (benefitSheet) {
    const decoded = XLSX.utils.decode_range(benefitSheet['!ref'] || 'A1:K14');
    const finalRow = Math.max(14, decoded.e.r + 1);

    for (let row = 5; row <= finalRow; row += 1) {
      const title = String(c('BENEFICIOS', `E${row}`, '')).trim();

      if (!title) {
        continue;
      }

      beneficios.tarjetas.push({
        titulo: title,
        tituloColor: color('BENEFICIOS', `F${row}`, '#000000'),
        descripcion: c('BENEFICIOS', `G${row}`, ''),
        descripcionColor: color('BENEFICIOS', `H${row}`, '#374151'),
        icono: c('BENEFICIOS', `I${row}`, '✦'),
        iconoFondoColor: color('BENEFICIOS', `J${row}`, '#E8DCC8'),
        cartaColor: color('BENEFICIOS', `K${row}`, '#FFFFFF')
      });
    }
  }

  if (!beneficios.tarjetas.length) {
    beneficios.tarjetas = FALLBACK.beneficios.tarjetas;
  }

  /* GALERÍA */
  const galeria = {
    titulo: c('GALERIA', 'C2', FALLBACK.galeria.titulo),
    tituloColor: color('GALERIA', 'C3', FALLBACK.galeria.tituloColor),
    fondoTipo: c('GALERIA', 'C4', FALLBACK.galeria.fondoTipo),
    fondoColor: color('GALERIA', 'C5', FALLBACK.galeria.fondoColor),
    fondoImagen: c('GALERIA', 'C6', FALLBACK.galeria.fondoImagen),
    imagenes: []
  };

  const gallerySheet = workbook.Sheets.GALERIA;

  if (gallerySheet) {
    const decoded = XLSX.utils.decode_range(gallerySheet['!ref'] || 'A1:F6');
    const finalRow = Math.max(6, decoded.e.r + 1);

    for (let row = 2; row <= finalRow; row += 1) {
      const filename = String(c('GALERIA', `F${row}`, '')).trim();

      if (filename) {
        galeria.imagenes.push(filename);
      }
    }
  }
  else {
    galeria.imagenes = [...FALLBACK.galeria.imagenes];
  }

  /* FAQ */
  const faq = {
    titulo: c('FAQ', 'C2', FALLBACK.faq.titulo),
    tituloColor: color('FAQ', 'C3', FALLBACK.faq.tituloColor),
    fondoTipo: c('FAQ', 'C4', FALLBACK.faq.fondoTipo),
    fondoColor: color('FAQ', 'C5', FALLBACK.faq.fondoColor),
    fondoImagen: c('FAQ', 'C6', FALLBACK.faq.fondoImagen),
    preguntas: []
  };

  const faqSheet = workbook.Sheets.FAQ;

  if (faqSheet) {
    const decoded = XLSX.utils.decode_range(faqSheet['!ref'] || 'A1:C42');
    const finalRow = Math.max(12, decoded.e.r + 1);

    for (let row = 7; row <= finalRow; row += 1) {
      const label = normalizeText(c('FAQ', `B${row}`, ''));
      const match = label.match(/^texto pregunta\s+(\d+)$/);

      if (!match) {
        continue;
      }

      const pregunta = String(c('FAQ', `C${row}`, '')).trim();

      if (!pregunta) {
        continue;
      }

      faq.preguntas.push({
        pregunta,
        preguntaColor: color('FAQ', `C${row + 1}`, '#000000'),
        respuesta: c('FAQ', `C${row + 2}`, ''),
        respuestaColor: color('FAQ', `C${row + 3}`, '#374151'),
        cartaColor: color('FAQ', `C${row + 4}`, '#FFFFFF')
      });
    }
  }
  else {
    faq.preguntas = [...FALLBACK.faq.preguntas];
  }

  /* CONTACTO */
  const contacto = {
    fondoTipo: c('CONTACTO', 'C2', FALLBACK.contacto.fondoTipo),
    fondoImagen: c('CONTACTO', 'C3', FALLBACK.contacto.fondoImagen),
    fondoColor: color('CONTACTO', 'C4', FALLBACK.contacto.fondoColor),
    etiqueta: c('CONTACTO', 'C5', FALLBACK.contacto.etiqueta),
    etiquetaColor: color('CONTACTO', 'C6', FALLBACK.contacto.etiquetaColor),
    titulo: c('CONTACTO', 'C7', FALLBACK.contacto.titulo),
    tituloColor: color('CONTACTO', 'C8', FALLBACK.contacto.tituloColor),
    descripcion: c('CONTACTO', 'C9', FALLBACK.contacto.descripcion),
    descripcionColor: color('CONTACTO', 'C10', FALLBACK.contacto.descripcionColor),

    direccion: c('CONTACTO', 'F3', FALLBACK.contacto.direccion),
    direccionColor: color('CONTACTO', 'F4', FALLBACK.contacto.direccionColor),
    direccionCartaColor: color('CONTACTO', 'F5', FALLBACK.contacto.direccionCartaColor),

    dias: c('CONTACTO', 'I3', FALLBACK.contacto.dias),
    diasColor: color('CONTACTO', 'I4', FALLBACK.contacto.diasColor),
    horas: c('CONTACTO', 'I5', FALLBACK.contacto.horas),
    horasColor: color('CONTACTO', 'I6', FALLBACK.contacto.horasColor),
    horarioCartaColor: color('CONTACTO', 'I7', FALLBACK.contacto.horarioCartaColor),

    telefono: c('CONTACTO', 'L3', FALLBACK.contacto.telefono),
    telefonoColor: color('CONTACTO', 'L4', FALLBACK.contacto.telefonoColor),
    telefonoCartaColor: color('CONTACTO', 'L5', FALLBACK.contacto.telefonoCartaColor),

    email: c('CONTACTO', 'O3', FALLBACK.contacto.email),
    emailColor: color('CONTACTO', 'O4', FALLBACK.contacto.emailColor),
    emailCartaColor: color('CONTACTO', 'O5', FALLBACK.contacto.emailCartaColor),

    mapsTexto: c('CONTACTO', 'C12', FALLBACK.contacto.mapsTexto),
    mapsBotonColor: color('CONTACTO', 'C13', FALLBACK.contacto.mapsBotonColor),
    mapsTextoColor: color('CONTACTO', 'C14', FALLBACK.contacto.mapsTextoColor),

    whatsappTexto: c('CONTACTO', 'C15', FALLBACK.contacto.whatsappTexto),
    whatsappBotonColor: color('CONTACTO', 'C16', FALLBACK.contacto.whatsappBotonColor),
    whatsappTextoColor: color('CONTACTO', 'C17', FALLBACK.contacto.whatsappTextoColor),
    whatsapp: c('CONTACTO', 'C18', FALLBACK.contacto.whatsapp),
    whatsappMensaje: c('CONTACTO', 'C19', FALLBACK.contacto.whatsappMensaje),
    zoomMapa: numberInRange(c('CONTACTO', 'C20', FALLBACK.contacto.zoomMapa), 3, 20, 16)
  };

  /* FOOTER */
  const footer = {
    fondoColor: color('FOOTER', 'C2', FALLBACK.footer.fondoColor),
    textoColor: color('FOOTER', 'C3', FALLBACK.footer.textoColor),
    descripcion: c('FOOTER', 'C5', FALLBACK.footer.descripcion),
    instagram: c('FOOTER', 'C6', FALLBACK.footer.instagram),
    facebook: c('FOOTER', 'C7', FALLBACK.footer.facebook),
    copyright: c('FOOTER', 'C8', FALLBACK.footer.copyright),
    mostrarCredito: boolFromText(c('FOOTER', 'C9', 'Sí'), true)
  };

  return {
    nav,
    inicio,
    menu,
    beneficios,
    galeria,
    faq,
    contacto,
    footer
  };
}


/* =========================================================
   COMPATIBILIDAD CON EL EXCEL ANTERIOR
   ========================================================= */

function rowsToMap(rows, keyName, valueName) {
  const result = {};

  rows.forEach((row) => {
    const key = String(row[keyName] ?? '').trim();

    if (key) {
      result[key] = row[valueName] ?? '';
    }
  });

  return result;
}

function parseLegacyWorkbook(workbook) {
  const getRows = (name) => {
    const sheet = workbook.Sheets[name];

    if (!sheet) {
      return [];
    }

    return XLSX.utils.sheet_to_json(sheet, { defval: '' });
  };

  const contenido = rowsToMap(getRows('CONTENIDO'), 'Campo', 'Valor');
  const menuRows = rowsToMap(getRows('MENU'), 'Campo', 'Valor');
  const designRows = getRows('DISENO');
  const benefitRows = getRows('BENEFICIOS');
  const colorRows = getRows('LISTA_COLORES');

  const colors = {};

  colorRows.forEach((row) => {
    const name = String(row.Color ?? '').trim();
    const hex = String(row.HEX ?? '').trim();

    if (name && validHex(hex)) {
      colors[name] = hex;
    }
  });

  const design = {};

  designRows.forEach((row) => {
    const key = String(row['Configuración'] ?? '').trim();
    const rawHex = String(row.Valor ?? '').trim();
    const name = String(row.Color ?? '').trim();

    if (key) {
      design[key] = validHex(rawHex)
        ? rawHex
        : (colors[name] || rawHex);
    }
  });

  const legacyBenefitsConfig = {};

  benefitRows.forEach((row) => {
    const key = String(row['Configuración'] ?? '').trim();
    const selection = String(row['Selección'] ?? '').trim();
    const rawHex = String(row.Valor ?? '').trim();

    if (!key) {
      return;
    }

    legacyBenefitsConfig[key] = validHex(rawHex)
      ? rawHex
      : (colors[selection] || selection || rawHex);
  });

  const cards = benefitRows
    .map((row) => ({
      titulo: row.Titulo ?? row.Título ?? '',
      descripcion: row.Descripcion ?? row.Descripción ?? '',
      icono: row.Icono ?? row['Ícono'] ?? '✦',
      tituloColor: design.text || '#000000',
      descripcionColor: design.muted || '#374151',
      iconoFondoColor: design.accent || '#E8DCC8',
      cartaColor: design.surface_bg || '#FFFFFF'
    }))
    .filter((item) => String(item.titulo || '').trim());

  return {
    nav: {
      logo: contenido.logo_archivo || FALLBACK.nav.logo,
      nombre: contenido.nombre || FALLBACK.nav.nombre,
      nombreColor: design.header_text || FALLBACK.nav.nombreColor,
      opcionesColor: design.header_text || FALLBACK.nav.opcionesColor,
      fondoColor: design.header_bg || FALLBACK.nav.fondoColor
    },

    inicio: {
      ...FALLBACK.inicio,
      fondoTipo: 'Imagen',
      fondoImagen: contenido.portada_archivo || 'portada.jpg',
      titulo: contenido.hero_titulo || FALLBACK.inicio.titulo,
      propuesta: contenido.hero_propuesta || FALLBACK.inicio.propuesta,
      ctaTexto: contenido.hero_cta_texto || FALLBACK.inicio.ctaTexto,
      ctaTipo: contenido.hero_cta_tipo || FALLBACK.inicio.ctaTipo,
      ctaUrl: contenido.hero_cta_url || '',
      ctaColor: design.primary || FALLBACK.inicio.ctaColor,
      ctaTextoColor: design.primary_text || FALLBACK.inicio.ctaTextoColor
    },

    menu: {
      ...FALLBACK.menu,
      fondoTipo: menuRows.fondo_tipo || 'Color',
      fondoImagen: menuRows.fondo_imagen || 'menu.jpg',
      fondoColor: design.menu_bg || FALLBACK.menu.fondoColor,
      titulo: menuRows.titulo || FALLBACK.menu.titulo,
      descripcion: menuRows.descripcion || FALLBACK.menu.descripcion,
      botonTexto: menuRows.boton_texto || FALLBACK.menu.botonTexto,
      botonColor: design.primary || FALLBACK.menu.botonColor,
      botonTextoColor: design.primary_text || FALLBACK.menu.botonTextoColor,
      pdfUrl: menuRows.pdf_url || FALLBACK.menu.pdfUrl,
      cartaColor: design.surface_bg || '#FFFFFF'
    },

    beneficios: {
      ...FALLBACK.beneficios,
      fondoTipo: legacyBenefitsConfig.fondo_tipo || contenido.beneficios_fondo_tipo || 'Color',
      fondoImagen: legacyBenefitsConfig.fondo_imagen || contenido.beneficios_fondo_imagen || 'beneficios.jpg',
      fondoColor: legacyBenefitsConfig.fondo_color || design.beneficios_bg || FALLBACK.beneficios.fondoColor,
      etiqueta: legacyBenefitsConfig.encabezado_etiqueta || contenido.beneficios_etiqueta || FALLBACK.beneficios.etiqueta,
      titulo: legacyBenefitsConfig.encabezado_titulo || contenido.beneficios_titulo || FALLBACK.beneficios.titulo,
      descripcion: legacyBenefitsConfig.encabezado_descripcion || contenido.beneficios_descripcion || FALLBACK.beneficios.descripcion,
      etiquetaColor: legacyBenefitsConfig.texto_color || design.beneficios_text || '#FFFFFF',
      tituloColor: legacyBenefitsConfig.texto_color || design.beneficios_text || '#FFFFFF',
      descripcionColor: legacyBenefitsConfig.texto_color || design.beneficios_text || '#FFFFFF',
      tarjetas: cards.length ? cards : FALLBACK.beneficios.tarjetas
    },

    galeria: {
      ...FALLBACK.galeria,
      imagenes: []
    },

    faq: {
      ...FALLBACK.faq,
      preguntas: []
    },

    contacto: {
      ...FALLBACK.contacto,
      fondoColor: design.main_bg || FALLBACK.contacto.fondoColor,
      etiquetaColor: design.accent || FALLBACK.contacto.etiquetaColor,
      titulo: contenido.contacto_titulo || FALLBACK.contacto.titulo,
      descripcion: contenido.contacto_descripcion || FALLBACK.contacto.descripcion,
      direccion: contenido.direccion || FALLBACK.contacto.direccion,
      telefono: contenido.telefono || FALLBACK.contacto.telefono,
      email: contenido.email || FALLBACK.contacto.email,
      whatsapp: contenido.whatsapp || FALLBACK.contacto.whatsapp,
      whatsappMensaje: contenido.whatsapp_mensaje || FALLBACK.contacto.whatsappMensaje,
      dias: String(contenido.horario || '').split(/[·|]/)[0]?.trim() || FALLBACK.contacto.dias,
      horas: String(contenido.horario || '').split(/[·|]/).slice(1).join(' ').trim() || FALLBACK.contacto.horas,
      direccionCartaColor: design.surface_bg || '#FFFFFF',
      horarioCartaColor: design.surface_bg || '#FFFFFF',
      telefonoCartaColor: design.surface_bg || '#FFFFFF',
      emailCartaColor: design.surface_bg || '#FFFFFF'
    },

    footer: {
      ...FALLBACK.footer,
      fondoColor: design.footer_bg || FALLBACK.footer.fondoColor,
      textoColor: design.footer_text || FALLBACK.footer.textoColor,
      descripcion: contenido.footer_descripcion || contenido.sobre_descripcion || FALLBACK.footer.descripcion,
      instagram: contenido.instagram || FALLBACK.footer.instagram,
      facebook: contenido.facebook || FALLBACK.footer.facebook,
      copyright: contenido.aviso_footer || FALLBACK.footer.copyright
    }
  };
}


/* =========================================================
   LEER BUFFER EXCEL
   ========================================================= */

function parseExcelBuffer(buffer) {
  if (typeof XLSX === 'undefined') {
    throw new Error('No se pudo cargar la librería XLSX.');
  }

  const workbook = XLSX.read(buffer, { type: 'array' });

  if (workbook.Sheets.NAV) {
    return parseNewWorkbook(workbook);
  }

  if (workbook.Sheets.CONTENIDO) {
    return parseLegacyWorkbook(workbook);
  }

  throw new Error(
    'El Excel no tiene la estructura esperada. Usa el Datos_Restaurante.xlsx actualizado.'
  );
}


/* =========================================================
   TEMA / VARIABLES CSS
   ========================================================= */

function applyTheme(config) {
  const {
    nav,
    inicio,
    menu,
    beneficios,
    galeria,
    faq,
    contacto,
    footer
  } = config;

  /* Header */
  setCssVar('--header-bg', nav.fondoColor);
  setCssVar('--header-name-text', nav.nombreColor);
  setCssVar('--header-nav-text', nav.opcionesColor);

  /* Inicio */
  setCssVar('--hero-bg', inicio.fondoColor);
  setCssVar('--hero-title-text', inicio.tituloColor);
  setCssVar('--hero-copy-text', inicio.propuestaColor);
  setCssVar('--hero-cta-bg', inicio.ctaColor);
  setCssVar('--hero-cta-text', inicio.ctaTextoColor);


  /* Menú */
  setCssVar('--menu-bg', menu.fondoColor);
  setCssVar('--menu-card-bg', menu.cartaColor);
  setCssVar('--menu-eyebrow-text', menu.etiquetaColor);
  setCssVar('--menu-title-text', menu.tituloColor);
  setCssVar('--menu-description-text', menu.descripcionColor);
  setCssVar('--menu-button-bg', menu.botonColor);
  setCssVar('--menu-button-text', menu.botonTextoColor);

  /* Beneficios */
  setCssVar('--benefits-bg', beneficios.fondoColor);
  setCssVar('--benefits-eyebrow-text', beneficios.etiquetaColor);
  setCssVar('--benefits-title-text', beneficios.tituloColor);
  setCssVar('--benefits-description-text', beneficios.descripcionColor);

  /* Galería */
  setCssVar('--gallery-bg', galeria.fondoColor);
  setCssVar('--gallery-title-text', galeria.tituloColor);

  /* FAQ */
  setCssVar('--faq-bg', faq.fondoColor);
  setCssVar('--faq-title-text', faq.tituloColor);

  /* Contacto */
  setCssVar('--contact-bg', contacto.fondoColor);
  setCssVar('--contact-eyebrow-text', contacto.etiquetaColor);
  setCssVar('--contact-title-text', contacto.tituloColor);
  setCssVar('--contact-description-text', contacto.descripcionColor);
  setCssVar('--address-card-bg', contacto.direccionCartaColor);
  setCssVar('--address-text', contacto.direccionColor);
  setCssVar('--hours-card-bg', contacto.horarioCartaColor);
  setCssVar('--hours-days-text', contacto.diasColor);
  setCssVar('--hours-hours-text', contacto.horasColor);
  setCssVar('--phone-card-bg', contacto.telefonoCartaColor);
  setCssVar('--phone-text', contacto.telefonoColor);
  setCssVar('--email-card-bg', contacto.emailCartaColor);
  setCssVar('--email-text', contacto.emailColor);
  setCssVar('--maps-button-bg', contacto.mapsBotonColor);
  setCssVar('--maps-button-text', contacto.mapsTextoColor);
  setCssVar('--whatsapp-button-bg', contacto.whatsappBotonColor);
  setCssVar('--whatsapp-button-text', contacto.whatsappTextoColor);

  /* Footer */
  setCssVar('--footer-bg', footer.fondoColor);
  setCssVar('--footer-text', footer.textoColor);
}


/* =========================================================
   BENEFICIOS
   ========================================================= */

function renderBenefits(items) {
  const grid = $('benefitsGrid');

  if (!grid) {
    return;
  }

  grid.replaceChildren();

  const finalItems = Array.isArray(items) && items.length
    ? items
    : FALLBACK.beneficios.tarjetas;

  finalItems.forEach((item) => {
    const article = document.createElement('article');
    article.className = 'benefit-card';

    article.style.setProperty(
      '--benefit-card-bg',
      validHex(item.cartaColor) ? item.cartaColor : '#FFFFFF'
    );

    article.style.setProperty(
      '--benefit-title-text',
      validHex(item.tituloColor) ? item.tituloColor : '#000000'
    );

    article.style.setProperty(
      '--benefit-body-text',
      validHex(item.descripcionColor) ? item.descripcionColor : '#374151'
    );

    article.style.setProperty(
      '--benefit-icon-bg',
      validHex(item.iconoFondoColor) ? item.iconoFondoColor : '#E8DCC8'
    );

    const icon = document.createElement('div');
    icon.className = 'benefit-icon';
    icon.setAttribute('aria-hidden', 'true');
    icon.textContent = String(item.icono || '✦');

    const title = document.createElement('h3');
    title.textContent = String(item.titulo || '');

    const description = document.createElement('p');
    description.textContent = String(item.descripcion || '');

    article.append(icon, title, description);
    grid.appendChild(article);
  });
}


/* =========================================================
   GALERÍA
   ========================================================= */

function renderGallery(items) {
  const grid = $('galleryGrid');

  if (!grid) {
    return;
  }

  grid.replaceChildren();

  const finalItems = Array.isArray(items)
    ? items
      .map((item) => String(item || '').trim())
      .filter(Boolean)
    : [];

  finalItems.forEach((filename, index) => {
    const safeFilename = safeAssetName(filename, '');

    if (!safeFilename) {
      return;
    }

    const article = document.createElement('article');
    article.className = 'gallery-card';

    const image = document.createElement('img');
    image.className = 'gallery-image';
    image.alt = `Imagen de la galería ${index + 1}`;
    image.loading = 'lazy';
    image.decoding = 'async';

    image.onerror = () => {
      article.remove();
      console.warn(`No se pudo cargar la imagen de galería: ${safeFilename}`);
    };

    image.src = `img/${safeFilename}?v=${Date.now()}`;

    article.appendChild(image);
    grid.appendChild(article);
  });
}


/* =========================================================
   FAQ
   ========================================================= */

function renderFaq(items) {
  const list = $('faqList');

  if (!list) {
    return;
  }

  list.replaceChildren();

  const finalItems = Array.isArray(items)
    ? items
    : [];

  finalItems.forEach((item) => {
    const questionText = String(item.pregunta || '').trim();

    if (!questionText) {
      return;
    }

    const details = document.createElement('details');
    details.className = 'faq-item';

    details.style.setProperty(
      '--faq-card-bg',
      validHex(item.cartaColor) ? item.cartaColor : '#FFFFFF'
    );

    details.style.setProperty(
      '--faq-question-text',
      validHex(item.preguntaColor) ? item.preguntaColor : '#000000'
    );

    details.style.setProperty(
      '--faq-answer-text',
      validHex(item.respuestaColor) ? item.respuestaColor : '#374151'
    );

    const summary = document.createElement('summary');
    summary.className = 'faq-question';
    summary.textContent = questionText;

    const answer = document.createElement('div');
    answer.className = 'faq-answer';

    const paragraph = document.createElement('p');
    paragraph.textContent = String(item.respuesta || '');

    answer.appendChild(paragraph);
    details.append(summary, answer);
    list.appendChild(details);
  });
}


/* =========================================================
   CTA PRINCIPAL
   ========================================================= */

function configureHeroCta(inicio, waUrl) {
  const cta = $('heroCta');

  if (!cta) {
    return;
  }

  cta.textContent = inicio.ctaTexto || 'Escríbenos por WhatsApp';
  cta.removeAttribute('target');
  cta.removeAttribute('rel');

  const type = normalizeText(inicio.ctaTipo);

  if (type === 'whatsapp') {
    cta.href = waUrl;
    cta.target = '_blank';
    cta.rel = 'noopener';
    return;
  }

  if (type === 'menu') {
    cta.href = '#menu';
    return;
  }

  if (type === 'contacto' || type === 'contact') {
    cta.href = '#contacto';
    return;
  }

  const customUrl = safeUrl(inicio.ctaUrl, '#contacto');
  cta.href = customUrl;

  if (!customUrl.startsWith('#')) {
    cta.target = '_blank';
    cta.rel = 'noopener';
  }
}


/* =========================================================
   CONTACTO
   ========================================================= */

function applyContact(contacto) {
  setText('contactoEtiqueta', contacto.etiqueta);
  setText('contactoTitulo', contacto.titulo);
  setText('contactoDescripcion', contacto.descripcion);
  setText('direccion', contacto.direccion);

  const days = document.querySelector('.contact-days');
  const hours = document.querySelector('.contact-hours');

  if (days) {
    days.textContent = String(contacto.dias || '');
  }

  if (hours) {
    hours.textContent = String(contacto.horas || '');
  }

  const phone = $('telefono');
  const footerPhone = $('footerTelefono');
  const footerPhoneText = $('footerTelefonoTexto');
  const telHref = contacto.telefono
    ? `tel:${cleanPhone(contacto.telefono)}`
    : '#';

  if (phone) {
    phone.textContent = contacto.telefono || '';
    phone.href = telHref;
  }

  if (footerPhone) {
    footerPhone.href = telHref;
  }

  if (footerPhoneText) {
    footerPhoneText.textContent = contacto.telefono || '';
  }

  const email = $('email');
  const footerEmail = $('footerEmail');
  const footerEmailText = $('footerEmailTexto');
  const emailHref = contacto.email
    ? `mailto:${String(contacto.email).trim()}`
    : '#';

  if (email) {
    email.textContent = contacto.email || '';
    email.href = emailHref;
  }

  if (footerEmail) {
    footerEmail.href = emailHref;
  }

  if (footerEmailText) {
    footerEmailText.textContent = contacto.email || '';
  }

  const mapsUrl = buildGoogleMapsUrl(contacto.direccion);
  const mapsLink = $('mapsLink');
  const footerAddressLink = $('footerDireccionLink');
  const mapsFrame = $('mapsFrame');

  if (mapsLink) {
    mapsLink.href = mapsUrl;
    mapsLink.textContent = contacto.mapsTexto || 'Abrir en Google Maps';
  }

  if (footerAddressLink) {
    footerAddressLink.href = mapsUrl;
  }

  setText('footerDireccion', contacto.direccion);

  if (mapsFrame) {
    mapsFrame.src = buildGoogleMapsEmbed(
      contacto.direccion,
      contacto.zoomMapa
    );
  }

  const hasWhatsapp = Boolean(cleanPhone(contacto.whatsapp));
  const waUrl = hasWhatsapp
    ? buildWhatsappUrl(contacto.whatsapp, contacto.whatsappMensaje)
    : '#';

  const contactWhatsapp = $('contactWhatsapp');
  const floatingWhatsapp = $('floatingWhatsapp');

  if (contactWhatsapp) {
    contactWhatsapp.href = waUrl;
    contactWhatsapp.hidden = !hasWhatsapp;
  }

  if (floatingWhatsapp) {
    floatingWhatsapp.href = waUrl;
    floatingWhatsapp.hidden = !hasWhatsapp;
  }

  setText(
    'contactWhatsappTexto',
    contacto.whatsappTexto || 'Enviar WhatsApp'
  );

  return waUrl;
}


/* =========================================================
   FOOTER
   ========================================================= */

function applyFooter(footer, businessName) {
  setText('footerNombre', businessName);
  setText('footerDescripcion', footer.descripcion);

  const defaultCopyright =
    `© ${businessName}. Todos los derechos reservados.`;

  const copyrightRaw = String(footer.copyright || '').trim();
  const copyrightNormalized = normalizeText(copyrightRaw);
  const copyrightText = ['automatico', 'auto'].includes(copyrightNormalized)
    ? defaultCopyright
    : (copyrightRaw || defaultCopyright);

  setText('footerAviso', copyrightText);

  const instagram = $('instagram');
  const instagramUrl = String(footer.instagram || '').trim();

  if (instagram) {
    instagram.href = safeUrl(instagramUrl);
    instagram.hidden = !instagramUrl;
  }

  const facebook = $('facebook');
  const facebookUrl = String(footer.facebook || '').trim();

  if (facebook) {
    facebook.href = safeUrl(facebookUrl);
    facebook.hidden = !facebookUrl;
  }

  const credit = $('footerCredit');

  if (credit) {
    credit.hidden = !footer.mostrarCredito;
  }
}


/* =========================================================
   APLICAR CONFIGURACIÓN COMPLETA
   ========================================================= */

function applyConfig(config) {
  const finalConfig = {
    nav: { ...FALLBACK.nav, ...(config.nav || {}) },
    inicio: { ...FALLBACK.inicio, ...(config.inicio || {}) },
    menu: { ...FALLBACK.menu, ...(config.menu || {}) },
    beneficios: {
      ...FALLBACK.beneficios,
      ...(config.beneficios || {})
    },
    galeria: {
      ...FALLBACK.galeria,
      ...(config.galeria || {})
    },
    faq: {
      ...FALLBACK.faq,
      ...(config.faq || {})
    },
    contacto: { ...FALLBACK.contacto, ...(config.contacto || {}) },
    footer: { ...FALLBACK.footer, ...(config.footer || {}) }
  };

  applyTheme(finalConfig);

  const businessName =
    String(finalConfig.nav.nombre || FALLBACK.nav.nombre).trim() ||
    FALLBACK.nav.nombre;

  document.title = businessName;
  setText('brandNombre', businessName);
  applyLogo(finalConfig.nav.logo, businessName);

  /* Fondos */
  applySectionBackground({
    selector: '.hero',
    mode: finalConfig.inicio.fondoTipo,
    filename: finalConfig.inicio.fondoImagen,
    fallbackFilename: 'portada.jpg'
  });

  applySectionBackground({
    selector: '.menu-section',
    mode: finalConfig.menu.fondoTipo,
    filename: finalConfig.menu.fondoImagen,
    fallbackFilename: 'menu.jpg'
  });

  applySectionBackground({
    selector: '.benefits-section',
    mode: finalConfig.beneficios.fondoTipo,
    filename: finalConfig.beneficios.fondoImagen,
    fallbackFilename: 'beneficios.jpg'
  });

  applySectionBackground({
    selector: '.gallery-section',
    mode: finalConfig.galeria.fondoTipo,
    filename: finalConfig.galeria.fondoImagen,
    fallbackFilename: 'galeria-fondo.jpg'
  });

  applySectionBackground({
    selector: '.faq-section',
    mode: finalConfig.faq.fondoTipo,
    filename: finalConfig.faq.fondoImagen,
    fallbackFilename: 'faq.jpg'
  });

  applySectionBackground({
    selector: '.contact-section',
    mode: finalConfig.contacto.fondoTipo,
    filename: finalConfig.contacto.fondoImagen,
    fallbackFilename: 'contacto.jpg'
  });

  /* Inicio */
  setText('heroTitulo', finalConfig.inicio.titulo);
  setText('heroPropuesta', finalConfig.inicio.propuesta);


  /* Menú */
  setText('menuEtiqueta', finalConfig.menu.etiqueta);
  setText('menuTitulo', finalConfig.menu.titulo);
  setText('menuDescripcion', finalConfig.menu.descripcion);
  setText('menuPdfBtn', finalConfig.menu.botonTexto);

  const menuButton = $('menuPdfBtn');

  if (menuButton) {
    menuButton.href = safeUrl(finalConfig.menu.pdfUrl, 'menu.pdf');
  }

  /* Beneficios */
  setText('beneficiosEtiqueta', finalConfig.beneficios.etiqueta);
  setText('beneficiosTitulo', finalConfig.beneficios.titulo);
  setText('beneficiosDescripcion', finalConfig.beneficios.descripcion);
  renderBenefits(finalConfig.beneficios.tarjetas);

  /* Galería */
  setText('galeriaTitulo', finalConfig.galeria.titulo);
  renderGallery(finalConfig.galeria.imagenes);

  /* FAQ */
  setText('faqTitulo', finalConfig.faq.titulo);
  renderFaq(finalConfig.faq.preguntas);

  /* Contacto + WhatsApp */
  const waUrl = applyContact(finalConfig.contacto);
  configureHeroCta(finalConfig.inicio, waUrl);

  /* Footer */
  applyFooter(finalConfig.footer, businessName);
}


/* =========================================================
   CARGAR EXCEL REMOTO
   ========================================================= */

async function loadExcelConfigRemote() {
  const response = await fetch(
    `Datos_Restaurante.xlsx?v=${Date.now()}`,
    { cache: 'no-store' }
  );

  if (!response.ok) {
    throw new Error(
      `No se pudo leer Datos_Restaurante.xlsx (${response.status}).`
    );
  }

  return parseExcelBuffer(await response.arrayBuffer());
}


/* =========================================================
   CARGAR EXCEL LOCAL
   ========================================================= */

async function loadLocalFile(file) {
  const buffer = await file.arrayBuffer();
  const config = parseExcelBuffer(buffer);
  applyConfig(config);
}


/* =========================================================
   NOTIFICACIONES
   ========================================================= */

let noticeTimeout;

function showNotice(message, duration = 5000) {
  const notice = $('configNotice');

  if (!notice) {
    console.log(message);
    return;
  }

  clearTimeout(noticeTimeout);
  notice.textContent = message;
  notice.hidden = false;

  noticeTimeout = setTimeout(() => {
    notice.hidden = true;
  }, duration);
}


/* =========================================================
   PANEL DE EXCEL LOCAL
   ========================================================= */

function createLocalPreviewPanel() {
  if (window.location.protocol !== 'file:') {
    return;
  }

  const panel = document.createElement('div');
  panel.id = 'localExcelPanel';
  panel.className = 'local-excel-panel';

  panel.innerHTML = `
    <div class="local-excel-title">
      Vista previa local
    </div>

    <div class="local-excel-copy">
      Conecta
      <strong>Datos_Restaurante.xlsx</strong>.
      Después puedes editarlo y guardar los cambios.
    </div>

    <button
      id="connectExcelBtn"
      class="local-excel-button"
      type="button"
    >
      📊 Conectar Excel
    </button>

    <div
      id="localExcelStatus"
      class="local-excel-status"
    >
      Excel no conectado. Haz clic en Conectar Excel y selecciona Datos_Restaurante.xlsx.
    </div>

    <input
      id="fallbackExcelInput"
      class="local-excel-input"
      type="file"
      accept=".xlsx"
    >
  `;

  document.body.appendChild(panel);

  $('connectExcelBtn')?.addEventListener('click', connectLocalExcel);

  $('fallbackExcelInput')?.addEventListener('change', async (event) => {
    const file = event.target.files[0];

    if (!file) {
      return;
    }

    try {
      await loadLocalFile(file);
      setLocalStatus('✅ Excel cargado correctamente.');
      showNotice('Excel cargado correctamente.');
    }
    catch (error) {
      console.error(error);
      setLocalStatus('❌ No se pudo leer el Excel.');
      showNotice(error.message || 'No se pudo leer el Excel.');
    }
  });
}

function setLocalStatus(message) {
  const status = $('localExcelStatus');

  if (status) {
    status.textContent = message;
  }
}


/* =========================================================
   FILE SYSTEM ACCESS API
   ========================================================= */

let localExcelHandle = null;
let lastLocalVersion = null;
let localWatcher = null;

async function connectLocalExcel() {
  if ('showOpenFilePicker' in window) {
    try {
      const handles = await window.showOpenFilePicker({
        multiple: false,
        types: [
          {
            description: 'Archivo Excel',
            accept: {
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': [
                '.xlsx'
              ]
            }
          }
        ]
      });

      localExcelHandle = handles[0];
      lastLocalVersion = null;

      await refreshLocalExcel(true);
      startExcelWatcher();

      setLocalStatus(
        '✅ Excel conectado. Guarda cambios en Excel y la página se actualizará automáticamente.'
      );

      return;
    }
    catch (error) {
      if (error.name === 'AbortError') {
        return;
      }

      console.warn('File System Access API no disponible:', error);
    }
  }

  const input = $('fallbackExcelInput');

  if (input) {
    input.style.display = 'block';
    input.click();
    setLocalStatus('Selecciona manualmente Datos_Restaurante.xlsx.');
  }
}

async function refreshLocalExcel(force = false) {
  if (!localExcelHandle) {
    return;
  }

  try {
    const file = await localExcelHandle.getFile();
    const version = `${file.lastModified}-${file.size}`;

    if (!force && version === lastLocalVersion) {
      return;
    }

    if (!force) {
      await new Promise((resolve) => setTimeout(resolve, 300));
    }

    const updatedFile = await localExcelHandle.getFile();
    await loadLocalFile(updatedFile);

    lastLocalVersion =
      `${updatedFile.lastModified}-${updatedFile.size}`;

    const time = new Date().toLocaleTimeString();

    setLocalStatus(`✅ Actualizado automáticamente a las ${time}`);
    console.log('Excel actualizado:', time);
  }
  catch (error) {
    console.error('Error leyendo Excel local:', error);
    setLocalStatus(
      '⚠️ No se pudo volver a leer el Excel. Conéctalo nuevamente.'
    );
    stopExcelWatcher();
  }
}

function startExcelWatcher() {
  stopExcelWatcher();

  localWatcher = setInterval(() => {
    refreshLocalExcel(false);
  }, 1000);
}

function stopExcelWatcher() {
  if (localWatcher) {
    clearInterval(localWatcher);
    localWatcher = null;
  }
}


/* =========================================================
   MENÚ HAMBURGUESA
   ========================================================= */

function setupNavigation() {
  const button = document.querySelector('.nav-toggle');
  const nav = $('navLinks');

  if (!button || !nav) {
    return;
  }

  const closeMenu = () => {
    nav.classList.remove('open');
    button.classList.remove('is-open');
    button.setAttribute('aria-expanded', 'false');
    button.setAttribute('aria-label', 'Abrir menú');
  };

  button.addEventListener('click', () => {
    const open = nav.classList.toggle('open');

    button.classList.toggle('is-open', open);
    button.setAttribute('aria-expanded', String(open));
    button.setAttribute(
      'aria-label',
      open ? 'Cerrar menú' : 'Abrir menú'
    );
  });

  nav.querySelectorAll('a').forEach((link) => {
    link.addEventListener('click', closeMenu);
  });

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      closeMenu();
    }
  });

  window.addEventListener('resize', () => {
    if (window.innerWidth > 780) {
      closeMenu();
    }
  });
}


/* =========================================================
   INICIALIZACIÓN
   ========================================================= */

window.addEventListener('DOMContentLoaded', async () => {
  /*
  La página siempre inicia SIN imágenes de la carpeta img.
  Las imágenes se activan únicamente cuando el Excel se lee
  correctamente (local o remoto).
  */
  applyConfig(NO_EXCEL_CONFIG);
  setupNavigation();

  if (window.location.protocol === 'file:') {
    createLocalPreviewPanel();

    showNotice(
      'Vista local: conecta Datos_Restaurante.xlsx para activar la actualización automática.',
      8000
    );

    return;
  }

  try {
    const excelConfig = await loadExcelConfigRemote();
    applyConfig(excelConfig);
  }
  catch (error) {
    console.warn(error);

    showNotice(
      'No se pudo cargar Datos_Restaurante.xlsx. Se mantiene la vista sin imágenes hasta conectar el Excel.'
    );
  }
});
