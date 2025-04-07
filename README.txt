Actualizacion_mundial.py

¿Qué hace este script?

Este programa automatiza la actualización de una plantilla de Excel usando datos de varios archivos .xlsx. Para cada archivo:

1. Lee los datos desde todas las hojas (saltando la primera fila).
2. Abre una plantilla de Excel base.
3. Limpia el formato de las hojas donde se insertarán nuevos datos.
4. Copia los datos nuevos en la plantilla.
5. Guarda una versión nueva con el sufijo "_actualizado".

Requisitos

- Python 3.7 o superior
- Librerías necesarias:
  - pandas
  - openpyxl

Instalación

Puedes instalar las dependencias automáticamente ejecutando:

pip install -r requirements.txt

Rutas importantes del proyecto

El script usa rutas absolutas definidas directamente en el código. Asegúrate de modificar estas rutas si tu estructura de carpetas es diferente.

Ejemplo de rutas dentro del script:
- source_dir = r'C:/Users/rodri/Downloads/Datos_Extraidos'
- template_path = r'C:/Users/rodri/Downloads/estadisticas_macro_shared/Plantillas/Mercado mundial - Plantilla A.xlsx'
- output_dir = r'C:/Users/rodri/Downloads/estadisticas_macro_shared/estadisticas_macro_shared/Resultado'

Recomendaciones:
- Evita dejar espacios innecesarios en los nombres de archivos o carpetas (como en “Plantilla A.xlsx”).
- Usa rutas con doble diagonal invertida (`\`) o prefija la cadena con `r''` para evitar errores de escape en Windows.
- Si planeas compartir o mover el script, considera usar rutas relativas o variables de entorno para mayor portabilidad.

Cómo usarlo

1. Coloca los archivos fuente .xlsx en la carpeta definida como source_dir.
2. Verifica que la plantilla esté en la ruta especificada como template_path.
3. Ejecuta el script.
4. Los archivos actualizados se guardarán en la carpeta definida como output_dir.

