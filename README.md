

# ExcelToSIE

ExcelToSIE sirve para exportar los asientos de cuentas bancarias desde una tabla de excel a un archivo de formato SIE, un formato estándar en los países nórdicos para programas de contabilidad. 

ExcelToSIE tiene fundamentalmente 2 funciones.

    1. Mediante un sistema de reglas, establecer predicciones sobre en qué cuenta contabilizar cada asiento.

    2. Exportar todos los asientos del excel a un archivo de formato SIE.

Estas dos funciones son independientes, aunque se ejecuten desde el mismo programa, y no se necesitan mutuamente.

ExcelToSIE funciona a través de dos componentes: un programa llamado ExcelToSIE.jar y un archivo de excel llamado verifikationer.xlsx.

* En el archivo de excel se deben de pegar todos los asientos de la cuenta bancaria, y escribir las cuentas de crédito y débito a la que van a ser asignadas en la contabilidad.

* El programa .jar se ejecuta una vez se ha rellenado el excel, para hacer las predicciones y generar el archivo en formato SIE. Primero pregunta si se desea usar las reglas para establecer predicciones, y a continuación si se desea exportar los asientos.

## Explicación de las pestañas en el archivo de excel

### Hoja "Verifikationer"

* Aquí se añaden todas las entradas. Se puede copiar y pegar directamente la tabla que ofrece la página web de Nordea.

* Para exportar una entrada son **obligatorios todos los campos menos Benämningar,** que es orientativo. 

* La columna "Exporterad" indica si una entrada ha sido ya exportada a un archivo SIE. Esta columna la rellena el programa. Si se quiere omitir temporalmente el análisis o la exportación de una entrada, se puede poner una X.

* Se pueden procesar como máximo 500 asientos al mismo tiempo.

### Hoja "Multipel"

* En esta pestaña se escriben las entradas que son más complejas, es decir, con multiples cuentas y cantidades en crédito y débito.

* Para crear una nueva entrada, se rellenan los datos básicos (fecha, número de verificación, mensaje) en la siguiente fila en blanco, y se van añadiendo las cuentas con las cantidades de débito o crédito empezando por esa fila y continuando en las siguientes (un par cuenta-cantidad) por cada nueva fila.

* No debe de dejarse ninguna fila de celdas en blanco entre dos entradas múltiples.

* Hay que asegurarse de que la **suma** de todos las cantidades de la columna **Debet** es igual a la suma de todas las cantidades de la cuenta **Kredit** para cada entrada.

* No hay problema en añadir varias filas con valores dirigidos a la misma cuenta. 

### Hoja "Regler"

* Aquí se añaden las reglas para clasificar automáticamente los asientos en cuentas concretas. El programa busca si hay coincidencias entre el texto que se escribe en los distintos parámetros de las reglas y cada uno de los asientos. Para los asientos que cumplan una regla, el programa escribe en el excel la cuenta de crédito y de débito correspondientes a las que van a ser asignados.

* No es obligatorio rellenar todos los parámetros para crear una regla, pero por lo menos deben estar rellenos **Debetkonto, Kreditkonto** y **al menos** uno de los siguientes tres: **nombre, mensaje o cantidad.**

* Los parámetros de las reglas son los siguientes:

    * **Nombre** : nombre del emisor o receptor de la transferencia. Se desaconseja poner un texto largo, como nombre y apellidos, pues el programa no encontrará coincidencias. En cambio es bueno poner una palabra corta, que al mismo tiempo sea única y no de pie a confusiones. También se puede poner el nombre de uno de los grupos de la parte derecha de la hoja de excel (nlark, nabr, inne), en cuyo caso el programa evaluará la regla para todos los nombres que se encuentran debajo (buscará las versiones cortas de los nombres por el motivo antes mencionado).

    * **Mensaje** el programa buscará coincidencias entre el texto aquí escrito y el del campo Meddelande (till/från) de la hoja Verifikationer. Si se quieren buscar varias palabras se pueden escribir separadas por `;` (semicólon).

    * **Cantidad** : Importe de la transferencia

    * **Margen** : cantidad por arriba y por debajo por la que se considera que un asiento cumple la regla. *Ejemplo:* si la regla dicta 7000 kr y el margen es de 200 kr, una entrada con 7250 kr no cumple la regla y una de 6900 kr sí.

    * **AND/OR** : Este parámetro va escrito entre el parámetro **nombre** y el parámetro **mensaje**, e indica si es necesario que estos dos parámetros se cumplan en una entrada al mismo tiempo para que esta sea válida (**AND**) o si basta con que uno de los dos lo cumpla para establecer la predicción (**OR**).

* **Es importante** escribir el nombre de la empresa (**företagsnamn**) en la casilla correspondiente del excel, para que el archivo SIE sea válido.

### Hoja "Arkiv"

Sirve para ir almacenando los asientos que ya han sido exportados. Se pueden cortar y pegar desde Verifikationer.

### Hoja "Konton"

Contiene las cuentas de la contabilidad junto a su descripción y el tipo de cuenta que es.

## Ejecución del programa, consejos prácticos.

* Para ejecutar el programa se necesita tener instalado la[ máquina virtual de Java](https://www.java.com/es/download/). Se abre un terminal en el directorio con el  archivo .jar y se escribe:  **java -jar ExcelToSIE.jar**. Alternativamente se puede hacer click en el archivo run.bat (solo Windows).

* El archivo Verifikationer.xlsx y el programa ExcelToSIE.jar deben de estar en la misma carpeta para funcionar correctamente.

* El programa exporta los asientos a un archivo en el mismo directorio, con nombre `verifikationer_nA_F.SI` dónde: A = [nº de asientos exportados] y B = [fecha]

* En la pestaña Verifikationer del excel se pueden ordenar los asientos según los distintos criterios (namn/meddelande/etc) para asignar las cuentas de crédito y débito más rápido si se hace manualmente. Para ello no hay más que clicar en la flechita de al lado de cada categoría.

