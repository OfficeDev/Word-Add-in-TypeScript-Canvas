# Ejemplo del complemento de Word de globos de imagen: cargar, editar e insertar imágenes

**Tabla de contenido**

* [Resumen](#summary)
* [Herramientas necesarias](#required-tools)
* [Cómo instalar certificados](#how-to-install-certificates)
* [Cómo configurar y ejecutar la aplicación](#how-to-set-up-and-run-the-app)
* [Cómo ejecutar el complemento en Word 2016 para Windows](#how-to-run-the-add-in-in-Word-2016-for-Windows)
* [Preguntas más frecuentes](#faq)
* [Preguntas y comentarios](#questions-and-comments)
* [Obtener más información](#learn-more)


## Resumen

Este ejemplo de complemento de Word muestra cómo:

1. Crear un complemento de Word con TypeScript.
2. Cargar imágenes del documento en el complemento.
3. Editar imágenes en el complemento mediante la API de lienzo HTML e insertar las imágenes en un documento de Word.
4. Implementar comandos de complemento que inicien un complemento de la cinta de opciones y ejecuten un script tanto desde la cinta de opciones como desde un menú contextual.
5. Usar Office UI Fabric para crear una experiencia nativa similar a Word para el complemento.

![](/readme-images/Word-Add-in-TypeScript-Canvas.gif)

Definición de **comando de complemento**: una extensión a la interfaz de usuario de Word que le permite iniciar el complemento en un panel de tareas o ejecutar un script desde la cinta de opciones o desde un menú contextual.

Si quiere ver esto en acción, vaya a [Configuración de Word 2016 para Windows](#word-2016-for-windows-set-up) y use este [manifiesto](https://github.com/OfficeDev/Word-Add-in-TypeScript-Canvas/blob/deploy2Azure/manifest-word-add-in-canvas.xml).

## Herramientas necesarias

Para usar el ejemplo del complemento de globos de imagen de Word, se requiere lo siguiente.

* Word 2016 16.0.6326.0000, una versión posterior o cualquier cliente compatible con la API de JavaScript de Word. En este ejemplo se realiza una comprobación de requisito para ver si se está ejecutando en un host compatible para las API de JavaScript.
* npm (https://www.npmjs.com/) para instalar las dependencias. Viene con [NodeJS](https://nodejs.org/en/).
* (Windows) [GIT Bash](http://www.git-scm.com/downloads).
* Clonar este repositorio en el equipo local.

> Nota: Word para Mac 2016 no admite comandos de complemento en este momento. Puede ejecutar este ejemplo en Mac sin los comandos de complemento.

## Cómo instalar certificados

Necesitará un certificado para ejecutar este ejemplo, ya que los comandos de complemento requieren HTTPS y, como los comandos de complementos no tienen interfaz de usuario, no puede aceptar certificados no válidos. Ejecute [./gen-cert.sh](#gen-cert.sh) para crear el certificado y, después, tendrá que instalar ca.crt en el almacén de entidades de certificación raíz de confianza (Windows).

## Cómo configurar y ejecutar la aplicación

1. Instale el Administrador de definición de TypeScript al escribir ```npm install typings -g``` en la línea de comandos.
2. Instale las definiciones de TypeScript identificadas en typings.json al ejecutar ```typings install``` en el directorio raíz del proyecto en la línea de comandos.
3. Instale las dependencias del proyecto identificadas en package.json al ejecutar ```npm install``` en el directorio raíz del proyecto.
4. Instale Gulp ```npm install -g gulp```.
5. Copie los archivos de Fabric y JQuery al ejecutar ```gulp copy:libs```. (Windows) Si tiene un problema aquí, asegúrese de que *%APPDATA%\npm* está en la variable de ruta de acceso.
6. Ejecute la tarea de Gulp predeterminada al ejecutar ```gulp``` desde el directorio raíz del proyecto. Si no se actualizan las definiciones de TypeScript, obtendrá un error aquí.

En este punto, ha implementado este complemento de ejemplo. Ahora debe indicarle a Word dónde encontrar el complemento.

### Configuración de Word 2016 para Windows

1. (Solo Windows) Descomprima y ejecute esta [clave del Registro](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Tools/AddInCommandsUndark) para activar la característica de comandos del complemento. Esto es necesario mientras los comandos de complementos son una **característica en vista previa**.
2. Cree un recurso compartido de red o [comparta una carpeta en la red](https://technet.microsoft.com/es-es/library/cc770880.aspx) y coloque el archivo de manifiesto [manifest-word-add-in-canvas.xml](manifest-word-add-in-canvas.xml) en él.
3. Inicie Word y abra un documento.
4. Seleccione la pestaña **Archivo** y haga clic en **Opciones**.
5. Haga clic en **Centro de confianza** y seleccione el botón **Configuración del Centro de confianza**.
6. Seleccione **Catálogos de complementos de confianza**.
7. En el cuadro **Dirección URL del catálogo**, escriba la ruta de red al recurso compartido de carpeta que contiene manifest-word-add-in-canvas.xml y después elija **Agregar catálogo**.
8. Active la casilla **Mostrar en menú** y haga clic en **Aceptar**.
9. Aparecerá un mensaje para informarle de que la configuración se aplicará la próxima vez que inicie Office. Cierre y vuelva a iniciar Word.

## Cómo ejecutar el complemento en Word 2016 para Windows

1. Abra un documento de Word.
2. En la pestaña **Insertar** de Word 2016, elija **Mis complementos**.
3. Seleccione la pestaña **Carpeta compartida**.
4. Elija **Image callout add-in** (Complemento Globos de imagen) y, después, seleccione **Insertar**.
5. Si su versión de Word admite los comandos de complemento, la interfaz de usuario le informará de que se ha cargado el complemento. Puede usar la pestaña **Callout add-in** (Complemento de globo) para cargar el complemento en la interfaz de usuario y para insertar una imagen en el documento. También puede usar el menú contextual del botón derecho para insertar una imagen en el documento.
6. Si los comandos del complemento no son compatibles con su versión de Word, el complemento se cargará en un panel de tareas. Debe insertar una imagen en el documento de Word para usar la funcionalidad del complemento.
7. Seleccione una imagen en el documento de Word y cárguela en el panel de tareas al seleccionar *Load image from doc* (Cargar imagen del documento). Ahora puede insertar globos en la imagen. Seleccione *Insert image into doc* (Insertar imagen en el documento) para colocar la imagen actualizada en el documento de Word. El complemento generará descripciones de marcador de posición para cada uno de los globos.

## Preguntas más frecuentes

* ¿Funcionarán los comandos de complemento en Mac y iPad? No, no funcionarán en Mac o iPad a partir de la publicación de este archivo Léame.
* ¿Por qué no aparece mi complemento en la ventana **Mis complementos**? El manifiesto del complemento puede tener un error. Le sugiero que valide el manifiesto en el [esquema del manifiesto](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Tools/XSD).
* ¿Por qué no se llama al archivo de función para los comandos de mis complementos? Los comandos de complemento requieren HTTPS. Ya que los comandos de complemento requieren TLS y no hay una interfaz de usuario, no puede ver si hay un problema de certificado. Si tiene que aceptar un certificado no válido en el panel de tareas, el comando de complemento no funcionará.
* ¿Por qué no responden los comandos de instalación npm? Probablemente sí que respondan. Tarda un poco en Windows.

## Preguntas y comentarios

Nos encantaría recibir sus comentarios sobre el ejemplo del complemento de Word Globos de imagen. Puede enviarnos sus preguntas y sugerencias a través de la sección [Problemas](https://github.com/OfficeDev/Word-Add-in-TypeScript-Canvas/issues) de este repositorio.

Las preguntas generales sobre desarrollo de complementos deben publicarse en [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Asegúrese de que sus preguntas o comentarios se etiquetan con [office-js], [word-addins] y [API]. Vemos esas etiquetas.

## Obtener más información

Aquí tiene más recursos para ayudarle a crear complementos basados en la API de JavaScript de Word:

* [Información general sobre la plataforma de complementos de Office](https://msdn.microsoft.com/es-es/library/office/jj220082.aspx)
* [Complementos de Word](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins.md)
* [Introducción a la programación de complementos de Word](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins-programming-guide.md)
* [Explorador de fragmentos de código para Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
* [Referencia de la API de JavaScript de complementos de Word](https://github.com/OfficeDev/office-js-docs/tree/master/word/word-add-ins-javascript-reference)
* [Ejemplo de SillyStories](https://github.com/OfficeDev/Word-Add-in-SillyStories): obtenga información sobre cómo cargar archivos docx desde un servicio e insertarlos en un documento de Word abierto.
* [Ejemplo de autenticación de servidor del complemento de Office](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth): obtenga información sobre cómo usar proveedores de Azure y Google OAuth para autenticar a usuarios de complemento.

## Copyright
Copyright (c) 2016 Microsoft. Todos los derechos reservados.
