// src/App.js
import React, { useState } from 'react';
import axios from 'axios';
import { Button, Checkbox, FormControlLabel, Typography, Box, Container, TextField, CircularProgress, Backdrop, Paper } from '@mui/material';

function App() {
  const [archivoStocks, setArchivoStocks] = useState("STOCKS_EA_EB_2024_06_06.xlsx"); // Identificador del archivo predeterminado
  const [archivoCoois, setArchivoCoois] = useState(null);
  const [downloadEA, setDownloadEA] = useState(true);
  const [downloadEB, setDownloadEB] = useState(true);
  const [loading, setLoading] = useState(false);

  const handleFileChange = (e, setter) => setter(e.target.files[0]);

  const handleSubmit = async () => {
    setLoading(true);
    const formData = new FormData();
    if (archivoStocks instanceof File) {
      formData.append('archivo_stocks', archivoStocks);  // El usuario ha cargado un nuevo archivo
    } else {
      formData.append('archivo_stocks', archivoStocks);  // Envía el identificador del archivo predeterminado
    }
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
    } finally {
      setLoading(false);
    }
  };

  return (
    <Container maxWidth="sm">
      <Box sx={{ my: 4 }}>
        <Typography variant="h4" gutterBottom>
          Correcciones SAP División 3
        </Typography>
        <Paper elevation={3} sx={{ p: 2, mb: 3, backgroundColor: '#f3faff' }}>
          <Typography variant="h6" gutterBottom>
            Cargar Archivo de Órdenes de Producción (COOIS)
          </Typography>
          <TextField
            type="file"
            fullWidth
            variant="outlined"
            margin="normal"
            onChange={(e) => handleFileChange(e, setArchivoCoois)}
            InputLabelProps={{
              shrink: true,
            }}
            label="Seleccionar archivo"
          />
        </Paper>
        <Box sx={{ mb: 3 }}>
          <TextField
            type="file"
            fullWidth
            variant="outlined"
            margin="normal"
            onChange={(e) => handleFileChange(e, setArchivoStocks)}
            InputLabelProps={{
              shrink: true,
            }}
            label="Cargar Archivo Stocks (Opcional)"
            helperText={archivoStocks instanceof File ? `Archivo de Stocks Actual: ${archivoStocks.name}` : `Archivo de Stocks Actual: ${archivoStocks}`}
          />
        </Box>
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
        {loading && (
          <Backdrop sx={{ color: '#fff', zIndex: (theme) => theme.zIndex.drawer + 1 }} open={loading}>
            <CircularProgress color="inherit" />
          </Backdrop>
        )}
      </Box>
    </Container>
  );
}

export default App;
