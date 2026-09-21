const express = require('express');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Servir archivos estáticos de la aplicación
app.use(express.static(path.join(__dirname, 'sys-a-kyte', 'app')));

// Servir carpeta de datos si es requerida
app.use('/datos', express.static(path.join(__dirname, 'sys-a-kyte', 'datos')));

// Fallback: redirigir cualquier ruta a index.html
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'sys-a-kyte', 'app', 'index.html'));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`SYS a Kyte ejecutándose en el puerto ${PORT}`);
});
