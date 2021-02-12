# refresh-excel-from-.net4

Programa que permite actualizar una lista definida de hojas excel en una carpeta predefinida.

Pasos para instalar la aplicación.

1. Copia el ejecutable junto con el fichero de configuracion config.ini en cualquier carpeta que el usuario elija

2. Rellenar los parametros del fichero config.ini:

pathlist = es el path en el que esta ubicado el excel que indica los excels a actualizar. Debe llamarse "Lista.xlxs"
pathexcels = es el path en el que debe estar ubicada la lista excel donde deben relacionarse los Excels a actualizar
timeout = el tiempo que requiere cada actualizacion en milisegundos.... incrementar este valor por encima del tiempo máximo que necesite cada excel para actualizarse.

Un ejemplo de contenido de dicho fichero:

pathlist = C:\proy\refresh-excel-from-.net4\Files\
pathexcels = C:\proy\refresh-excel-from-.net4\Files\Excel\
timeout = 3000
