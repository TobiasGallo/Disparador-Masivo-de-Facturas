# Disparador Masivo de Facturas  📧⚡  

**¡Automatiza el envío de facturas por correo directamente desde Outlook!**  
Una herramienta desarrollada en Python con interfaz gráfica para enviar automáticamente facturas a clientes vía Gmail, basándose en lecturas de datos de la propia factura y buscando su mail en un Excel con datos del cliente. 
Como encuentra el mail? Utiliza los datos leídos de la factura (DNI y Nombre), los busca en el Excel y extrae su mail de dicha columna.  

---

### 🔥 Funcionalidades estrella:  
✅ **Gestión multi-cuenta Outlook**:  
- Vincula/desvincula cuentas corporativas  
- Selección dinámica de remitente activo  
- Visualización de cuentas disponibles  

✅ **Procesamiento inteligente**:  
- Extracción automática de DNI y nombres desde PDF  
- Búsqueda cruzada en Excel (CUIT/DNI + Nombre)  
- Validación de datos integrada  

✅ **Envíos masivos profesionales**:  
- Plantillas preconfiguradas de correo  
- Adjuntado seguro de documentos  
- Registro detallado de cada operación  

---

### ⚙️ Tecnología bajo el capó:  
🐍 Desarrollado en Python con:  
- 🖼️ Interfaz gráfica en **Tkinter**  
- 📊 Procesamiento de datos con **Pandas**  
- 📑 Lectura avanzada de PDFs con **PyPDF2**  
- 📧 Integración Outlook via **pywin32**  

---
