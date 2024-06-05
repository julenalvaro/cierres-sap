// PATH: frontend/src/App.js

import React, { useState } from 'react';
import axios from 'axios';
import Button from '@mui/material/Button';
import Checkbox from '@mui/material/Checkbox';
import FormControlLabel from '@mui/material/FormControlLabel';

function App() {
  const [archivoStocks, setArchivoStocks] = useState(null);
  const [archivoCoois, setArchivoCoois] = useState(null);
  const [downloadEA, setDownloadEA] = useState(true);
  const [downloadEB, setDownloadEB] = useState(true);

  const handleFileChange = (e, setter) => setter(e.target.files[0]);

  const handleSubmit = async () => {
    const formData = new FormData();
    formData.append('archivo_stocks', archivoStocks);
    formData.append('archivo_coois', archivoCoois);
    formData.append('download_ea', downloadEA);
    formData.append('download_eb', downloadEB);

    try {
      // El frontend ahora siempre manejará un archivo ZIP
      const response = await axios.post('http://localhost:8000/excel/generate_excel/', formData, {
        responseType: 'blob',
      });
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'output.zip'); // Siempre descarga como ZIP
      document.body.appendChild(link);
      link.click();

    } catch (error) {
      console.error('Error al generar el archivo:', error);
    }
  };

  return (
    <div style={{ margin: 20 }}>
      <h1>Generador de Archivos Excel</h1>
      <input type="file" onChange={(e) => handleFileChange(e, setArchivoStocks)} />
      <input type="file" onChange={(e) => handleFileChange(e, setArchivoCoois)} />
      <FormControlLabel
        control={<Checkbox checked={downloadEA} onChange={() => setDownloadEA(!downloadEA)} />}
        label="Descargar EA"
      />
      <FormControlLabel
        control={<Checkbox checked={downloadEB} onChange={() => setDownloadEB(!downloadEB)} />}
        label="Descargar EB"
      />
      <Button variant="contained" color="primary" onClick={handleSubmit}>
        Generar
      </Button>
    </div>
  );
}

export default App;
