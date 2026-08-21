/* main.js — SEO Auditor */

/**
 * Inicia una auditoría forzada desde la interfaz de usuario vía AJAX.
 * Mientras se ejecuta en segundo plano, mostramos una barra de progreso simulada 
 * y bloqueamos el botón para evitar que el usuario lance peticiones repetidas.
 */
async function forzarAuditoria(marcaId, marcaNombre, btn) {
  if (!confirm(`¿Iniciar auditoría en vivo para ${marcaNombre}?\nEsto puede tardar varios minutos dependiendo de la cantidad de URLs.`)) return;

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Auditando...';

  const progWrap  = document.getElementById(`progress-${marcaId}`);
  const progBar   = document.getElementById(`prog-bar-${marcaId}`);
  const progLabel = document.getElementById(`prog-label-${marcaId}`);

  if (progWrap) {
    progWrap.style.display = 'block';
    progBar.style.width = '5%';
    progLabel.textContent = 'Iniciando auditoría en segundo plano...';
  }

  try {
    const resp = await fetch(`/auditar/${marcaId}`, { method: 'POST' });
    const data = await resp.json();

    if (data.ok) {
      if (progLabel) progLabel.textContent = data.message;

      // Hacemos que la barra avance un poco al azar para que el usuario sienta
      // que el sistema está trabajando, mientras esperamos la confirmación del servidor.
      let pct = 10;
      const progressInterval = setInterval(() => {
        pct = Math.min(pct + Math.random() * 3, 92);
        if (progBar) progBar.style.width = pct + '%';
      }, 4000);

      // Consultamos el estado real al servidor cada 5 segundos (Polling)
      const pollInterval = setInterval(async () => {
        try {
          const statusResp = await fetch(`/audit_status/${marcaId}`);
          const statusData = await statusResp.json();

          if (statusData.status === 'done') {
            clearInterval(progressInterval);
            clearInterval(pollInterval);
            if (progBar)   progBar.style.width = '100%';
            if (progLabel) progLabel.textContent = '¡Completado! Recargando...';
            setTimeout(() => location.reload(), 1500);
          } else if (statusData.status === 'error') {
            clearInterval(progressInterval);
            clearInterval(pollInterval);
            if (progLabel) progLabel.textContent = statusData.message || 'Error en la auditoría.';
            if (progBar)   progBar.style.width = '100%';
            progBar.style.background = 'var(--danger, #e74c3c)';
            resetBtn(btn);
          }
          // Si el estado dice 'running', no hacemos nada y esperamos al siguiente ciclo
        } catch (e) {
          // Si hay algún problema de red al consultar el estado, lo ignoramos y 
          // volvemos a intentarlo en 5 segundos.
          console.warn('Error consultando estado de auditoría:', e);
        }
      }, 5000);

      // Ponemos un límite de 20 minutos por seguridad. Algunas páginas muy grandes
      // o que no son responsive pueden tardar mucho en terminar. Si pasamos el tiempo,
      // recargamos para ver cómo quedó.
      setTimeout(() => {
        clearInterval(progressInterval);
        clearInterval(pollInterval);
        if (progBar)   progBar.style.width = '100%';
        if (progLabel) progLabel.textContent = 'Tiempo máximo alcanzado. Recargando...';
        setTimeout(() => location.reload(), 1500);
      }, 1200000);

    } else {
      alert('Error: ' + (data.error || 'Desconocido'));
      resetBtn(btn);
    }
  } catch (e) {
    alert('Error de conexión: ' + e.message);
    resetBtn(btn);
  }
}

function resetBtn(btn) {
  btn.disabled = false;
  btn.innerHTML = '▶ Auditar ahora';
}

async function probarEmail(marcaId, btn) {
  const originalHtml = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = '...';
  
  try {
    const resp = await fetch(`/test_email/${marcaId}`, { method: 'POST' });
    const data = await resp.json();
    if (data.ok) {
      alert(data.message);
    } else {
      alert('Error: ' + data.error);
    }
  } catch (e) {
    alert('Error de conexión: ' + e.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = originalHtml;
  }
}

/* Los mensajes de éxito/error (flash) desaparecen solos después de 5 segundos
   para no saturar la pantalla. Y también inicializamos el toggle del tema oscuro. */
document.addEventListener('DOMContentLoaded', () => {
  // Ocultar notificaciones flash
  document.querySelectorAll('.flash').forEach(el => {
    setTimeout(() => {
      el.style.transition = 'opacity .5s';
      el.style.opacity = '0';
      setTimeout(() => el.remove(), 500);
    }, 5000);
  });

  // Lógica de alternancia de tema (Modo Oscuro)
  const themeToggle = document.getElementById('theme-toggle');
  if (themeToggle) {
    const lightIcon = document.getElementById('theme-toggle-light-icon');
    const darkIcon = document.getElementById('theme-toggle-dark-icon');

    // Función para actualizar la visibilidad de los iconos de Sol/Luna
    const updateThemeIcons = (theme) => {
      if (theme === 'dark') {
        if (lightIcon) lightIcon.style.display = 'block';
        if (darkIcon) darkIcon.style.display = 'none';
      } else {
        if (lightIcon) lightIcon.style.display = 'none';
        if (darkIcon) darkIcon.style.display = 'block';
      }
    };

    // Obtener tema actual inicializado en el HTML
    const currentTheme = document.documentElement.getAttribute('data-theme') || 'light';
    updateThemeIcons(currentTheme);

    // Manejar el evento de clic
    themeToggle.addEventListener('click', () => {
      const activeTheme = document.documentElement.getAttribute('data-theme') || 'light';
      const targetTheme = activeTheme === 'dark' ? 'light' : 'dark';
      
      document.documentElement.setAttribute('data-theme', targetTheme);
      localStorage.setItem('theme', targetTheme);
      updateThemeIcons(targetTheme);
    });
  }
});
