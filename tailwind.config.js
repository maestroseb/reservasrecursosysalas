// Configuración para regenerar tailwind-css.html (ver README > "Regenerar el CSS")
module.exports = {
  content: ['./*.html', './*.gs', '!./tailwind-css.html'],
  theme: { extend: {} },
  plugins: []
};
