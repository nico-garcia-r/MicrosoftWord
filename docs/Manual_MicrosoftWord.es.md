# Microsoft Word
  
Modulo para trabajar con Microsoft Word  

*Read this in other languages: [English](Manual_MicrosoftWord.md), [Portugues](Manual_MicrosoftWord.pr.md), [Español](Manual_MicrosoftWord.es.md).*
  
![banner](C:\Users\nicog\Desktop\Rocketbot\modules\MicrosoftWord\docs\imgs\banner.png)
## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de Rocketbot.  



## Descripción de los comandos

### Nuevo documento
  
Crea un nuevo documento word
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|

### Abrir Documento
  
Abre un documento de Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Archivo|Abre el documento especificado|archivo.docx|
|Sesión|Sesión del archivo|Word1|

### Leer documento
  
Extrae texto de documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Resultado|Almacena el resultado en una variable|Variable|
|Sesión|Sesión del archivo|Word1|
|Agregar Detalles|||

### Copiar y pegar texto
  
Copiar texto entre rangos del documento Word y pegarlo en otro documento.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Inicio del rango|Comienzo del rango|0|
|Fin del rango|Fin del rango|40|
|Sesión del archivo a copiar|Sesión del archivo|Word1|
|Archivo|Elige el documento|archivo.docx|

### Copiar texto
  
Copiar texto al portapapeles entre rangos del documento Word
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Inicio del rango|Comienzo del rango|0|
|Fin del rango|Fin del rango|40|
|Sesión|Sesión del archivo|Word1|

### Pegar texto
  
Pegar texto del portapapeles al documento Word
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|

### Contar caracteres
  
Contar caracteres de un párrafo específico
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Párrafo|Párrafo a contar caracteres|1|
|Resultado|Almacena el resultado en una variable|Variable|

### Agregar tabla
  
Agregar tabla en un documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Numero de filas|Numero de filas que tendrá la tabla|3 |
|Numero de columnas|Numero de columnas que tendrá la tabla|4 |
|Estilo de tabla|Estilo de los bordes de la tabla||
|Sesión|Sesión del archivo|Word1|
|Estilos del borde|||

### Leer tablas
  
Extrae los datos de las tablas en el documento
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Tabla a leer|Número de tabla de la cual se leerá el contenido|1|
|Sesión|Sesión del archivo|Word1|
|Resultado|Almacena el resultado en una variable|Variable|

### Editar tabla
  
Editar tabla de un documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Numero de tabla|Número de tabla que será editada|1|
|Sesión|Sesión del archivo|Word1|
|Ingrese el numero de fila a eliminar|Numero de fila a eliminar de la tabla| |
|Ingrese el numero de columna a eliminar|Numero de columna a eliminar de la tabla| |
|Insertar fila|Si se selecciona, agrega una fila al final de la tabla||
|Insertar columna|Si se selecciona, agrega una columna al final de la tabla||
|Ancho de columna|Ancho que tendrá cada columna de la tabla|140|
|Alto de fila|Alto que tendrá cada fila de la tabla|25|

### Guardar documento
  
Guarda el documento Word abierto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Guardar archivo|Guarda el archivo en la ruta especificada|archivo.docx|

### Escribir en documento
  
Escribe en un documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Escriba texto|Texto que se escribirá en el documento|Lorem ipsum |
|Tipo de texto|Selector del tipo de texto||
|Nivel||1-9|
|Tamaño de fuente||12|
|Alineación|||
|Negrita|||
|Cursiva|||
|Subrayar|||

### Cerrar documento
  
Cierra el documento que se está ejecutando
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|

### Insertar página
  
Inserta una nueva página al documento
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|

### Agregar imagen
  
Agrega una imagen al documento
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Ruta de la imagen|Ruta de imagen que sera agregada debajo del ultimo parrafo|imagen.jpg|

### Convertir a PDF
  
Convierte documento Word a PDF.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Guardar archivo|Ruta del archivo donde se creará el PDF|archivo.pdf|

### Buscar Texto en párrafo
  
Busca el párrafo donde se encuentra el texto indicado.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Texto a Buscar|Texto que sera usado para localizar el parrafo|Hola mundo|
|Nombre de la variable|Almacena el resultado en una variable|Varible|

### Contar párrafos
  
Cuenta la cantidad de párrafos del documento. Incluye los campos de tablas.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Nombre de la variable|Almacena el resultado en una variable|Varible|

### Remplazar texto en párrafo
  
Remplaza el texto de un párrafo.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Texto a Buscar||Hola mundo|
|Texto a Remplazar|Texto que sera reemplazado|Hola mundo|
|Lista de párrafo|Parrafos donde buscara el texto especificado|Separados por comas ',' ejemplo: 1,2|

### Borrar párrafo
  
Borra un párrafo del documento.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Número de párrafo|Numero de parrafo que será eliminado|1|
|Nombre de la variable donde se guardará el párrafo eliminado|Variable donde se guardará el texto que incluía el párrafo eliminado|Variable|

### Agregar texto a un bookmark
  
Agregar texto a un bookmark.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Texto a agregar||Hola mundo|
|Nombre del Marcador||Marcador 1|
