# Automatización con python y selenium

Este código es un conjunto de pruebas automatizadas usando el framework de Python unittest, que se utiliza para verificar que los resultados obtenidos son los esperados en cada una de las etapas de un proceso automatizado.

En este caso, el script ejecuta una serie de pasos para extraer un informe de preinscritos desde una página web. Para esto, se hace uso del navegador Chrome a través de la librería selenium, la cual es capaz de interactuar con el navegador y emular las acciones que un usuario humano haría en la página.

En concreto, el script primero configura la ruta de descarga de un archivo, elimina cualquier archivo preexistente con el nombre "preinscritos" en dicha ruta, y luego inicia sesión en una página de autenticación mediante el envío de credenciales. Posteriormente, abre una nueva pestaña del navegador y accede a una página donde se encuentra el informe que se desea extraer.

En el método test_informe, se realizan las acciones para extraer el informe de preinscritos, como hacer clic en los botones necesarios y seleccionar la fecha para el rango de los datos requeridos. Una vez obtenido el informe, el siguiente paso sería realizar las verificaciones necesarias para asegurarse de que los datos extraídos son los esperados.

Este script se ejecuta mediante la línea de comandos y se puede incluir en un flujo de integración continua, tal como lo sería una acción en GitHub Actions. El objetivo es automatizar el proceso de extracción del informe de preinscritos y asegurarse de que se realiza correctamente sin intervención manual.
