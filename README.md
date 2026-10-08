# Informes de no conformidades RETIE

Aplicación de escritorio para crear proyectos independientes, importar listas de verificación, registrar resultados por requisito y generar borradores sobre la hoja **Constructor** del F-GI-51. No contiene hallazgos predefinidos de CARACOLÍ ni de otra obra.

## Abrir en Windows

Extraer todos los archivos del paquete en una carpeta nueva y hacer doble clic en **INICIAR.bat**. Utiliza Python 3.14 o 3.13 ya instalado y prepara las dependencias si faltan. También puede ejecutarse con `python aplicacion.py` tras instalar `requirements.txt`.

1. Crear un proyecto y completar nombre, dirección, código, reglamento y carpeta de Drive.
2. Elegir el formato F-GI-51 suministrado, con la hoja Constructor. La aplicación verifica la estructura conocida. Otros formatos requieren un mapeo específico.
3. Importar por separado la lista de cada alcance: red subterránea, red aérea o subestación tipo poste. Confirmar que la lista y sus numerales corresponden al régimen normativo elegido.
4. Seleccionar un requisito y pulsar **Revisar requisito / registrar NC**. Registrar resultado, numeral, descripción, documento/página/enlace y una o varias fotos. Los resultados son Pendiente, Cumple, No cumple, No aplica y Evidencia insuficiente.
5. Solo **No cumple** pasa al informe de NC. Su descripción debe comenzar por **No se evidencia…**. La falta de evidencia no implica automáticamente incumplimiento. La validación de un numeral es una confirmación manual y se reinicia si se modifica el numeral o el reglamento.
6. Guardar el proyecto como `*.retie.json`. Sus fotos se copian a una carpeta propia de evidencias. El proyecto puede reabrirse y editarse. Para trasladarlo a otro equipo, copiar también su carpeta de evidencias y seleccionar otra vez el formato si su ubicación cambió.
7. Generar los Excel. Los datos, fotos y fórmulas de ejemplo de otras obras no se incorporan. Se mantienen los encabezados, columnas y estilos del Constructor; las filas se ajustan al número de NC. La primera foto se coloca en la fila y las demás se conservan en anexos, con trazabilidad.
8. Revisar los borradores y pulsar **Subir a Drive**. El destino es la carpeta del proyecto, dentro de **Informes de NC**. Cada exportación tiene nombre único, y las cargas crean archivos nuevos; no reemplazan versiones anteriores. Si una carga falla a mitad, consultar Drive antes de repetirla.

Los informes son preliminares: quedan pendientes firmas, subsanación y validación final. La aplicación no emite certificados ni determina automáticamente el reglamento o las NC. No lee por sí sola toda una carpeta de Drive ni analiza automáticamente documentos y fotografías. La primera versión sirve para que el inspector registre la revisión. Seleccionar «Nuevo proyecto» elimina del formulario los datos anteriores, tras avisar si hay cambios sin guardar.

## Google Drive

Reutiliza el token local de la conexión ya configurada en `~/.informes-nc-retie/token.json`. Si necesita autorización y no existe un cliente junto al programa, solicita el JSON OAuth de escritorio. Las credenciales no se guardan en proyectos ni se incluyen en el paquete.

Google Drive API debe estar habilitada y la cuenta debe ser usuario de prueba autorizado y tener permiso para añadir archivos. Se solicita el permiso `https://www.googleapis.com/auth/drive` para utilizar una carpeta existente. El alcance limitado `drive.file` requiere incorporar selección mediante Google Picker; no está integrado en esta versión.

Nunca subir `credentials.json`, `client_secret*.json`, `token.json`, proyectos ni evidencias al repositorio. La autorización de Google se ejecuta en el computador del usuario, no conecta al agente cloud con su cuenta.

## Desarrollo y validación

Usar el checkout existente. Las tareas cloud ya están aisladas y no requieren crear un worktree.

```sh
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
```

Las pruebas verifican aislamiento de proyectos, importación, resultados, conservación de fotos, protección de archivos originales, exportación Constructor, filas dinámicas y carga con API simulada. La autorización y carga real requieren el equipo y la cuenta del usuario. Las utilidades anteriores específicas de CARACOLÍ no se incluyen en esta aplicación general.
