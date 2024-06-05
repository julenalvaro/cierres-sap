// PATH: frontend/src/App.js

// src/App.js
import React, { useState } from 'react';
import axios from 'axios';
import { Button, Checkbox, FormControlLabel, Typography, Box, Container, TextField } from '@mui/material';

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
      const response = await axios.post('http://localhost:8000/excel/generate_excel/', formData, {
        responseType: 'blob',
      });
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'output.zip');
      document.body.appendChild(link);
      link.click();
    } catch (error) {
      console.error('Error al generar el archivo:', error);
    }
  };

  return (
    <Container maxWidth="sm">
      <Box sx={{ my: 4 }}>
        <Typography variant="h4" gutterBottom>
          Generador de Archivos Excel
        </Typography>
        <TextField
          type="file"
          fullWidth
          variant="outlined"
          margin="normal"
          onChange={(e) => handleFileChange(e, setArchivoStocks)}
          InputLabelProps={{
            shrink: true,
          }}
          label="Cargar Archivo Stocks"
        />
        <TextField
          type="file"
          fullWidth
          variant="outlined"
          margin="normal"
          onChange={(e) => handleFileChange(e, setArchivoCoois)}
          InputLabelProps={{
            shrink: true,
          }}
          label="Cargar Archivo Coois"
        />
        <FormControlLabel
          control={<Checkbox checked={downloadEA} onChange={() => setDownloadEA(!downloadEA)} />}
          label="Descargar EA"
        />
        <FormControlLabel
          control={<Checkbox checked={downloadEB} onChange={() => setDownloadEB(!downloadEB)} />}
          label="Descargar EB"
        />
        <Button variant="contained" color="primary" fullWidth onClick={handleSubmit} sx={{ mt: 2 }}>
          Generar
        </Button>
      </Box>
    </Container>
  );
}

export default App;
