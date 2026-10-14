export default {
  start: () => {
    const navigation = document.querySelector('#autocollapse .container-xxl');
    if (!navigation || navigation.querySelector('.inegi-header-brand')) return;

    const pageRoot = document.querySelector('meta[name="docfx:rel"]')?.content ?? '';
    const logo = document.createElement('img');
    logo.className = 'inegi-header-logo';
    logo.src = new URL(`${pageRoot}guia/imagenes/logos/INEGI_Logotipo.jpg`, window.location.href).href;
    logo.alt = 'INEGI';

    const link = document.createElement('a');
    link.className = 'inegi-header-brand';
    link.href = 'https://www.inegi.org.mx/';
    link.target = '_blank';
    link.rel = 'noopener noreferrer';
    link.setAttribute('aria-label', 'Visitar el sitio del INEGI (abre en otra pestaña)');
    link.appendChild(logo);
    navigation.appendChild(link);
  }
};
