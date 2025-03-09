# Convertir documentos a PDF

El script convierte archivos a PDF desde el explorador de archivos (funciona en Windows 10 y 11). Al hacer clic derecho sobre un archivo (como .docx, .xlsx, .pptx) y elegir "Convertir a PDF" en el menú contextual, se ejecuta un script de PowerShell (ConvertToPDF.ps1) que detecta la extensión: usa COM para convertir documentos de Office a PDF. 
El PDF se guarda en la misma carpeta con el mismo nombre. 
Este proceso es oculto gracias a un script VBScript (RunConvert.vbs) que ejecuta PowerShell sin mostrar una ventana de consola.

PROCESO DE INSTALACION:
1. Crear una carpeta en C:\ llamada Scripts
2. Dentro de Scripts crear otra y nombrarla ConvertToPDF. Debe quedar la ruta C:\Scripts\ConvertToPDF
3. Hacer doble click en ConvertToPDF.reg y seleccionar si
4. En caso de querer eliminar la funcion del Registro del sistema seleccionar RemoveConvertToPDF.reg
