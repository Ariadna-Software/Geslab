Attribute VB_Name = "Versiones"

'Este modulo se utilizará para guardar informacion
'sobre las modificaciones
'De momento solo recoge la fecha de inicio del
'proyecto:
'         Jueves, 4 de Octubre de 2001


'Modificacion del 'DATOSOK'


'###############################################################
'###############                                          ######
'------------------------------------------------
'Version 2.0.1         Lunes 5 de Octubre de 2001
'------------------------------------------------
'Hemos puesto el control ACTIVELOCK
'El formulario Main se muestra con los botones como imagenes
'El clase de configuracion tiene un metodo LEER


'###############################################################
'###############                                          ######
'--------------------------------------------------
'Version 2.0.2      Miercoles 21 de Octubre de 2001
'--------------------------------------------------
'Se ha preparado para que importe los ficheros desde lo que
'seria un reloj de produccion
'Otra cosa importante es que lo del reloj de produccion esta
'basado en la coop de alcira


'Se añade dentro del campo de control de un trabajador
'ahora tenemos
    '   0.- La aplicacion controla todo, incluso los execeso en
    '       las entradas y salidas generando las incidencias necesarias
    '       y comprobando las manuales
    '   1.- El trabj dispone de un horario flexible. Se recogen
    '       los fichajes pero no se generean incidencias
    '   2.- AÑADIDO en esta version
    '       Del horario del trabajador solo nos interesan las horas
    '       que le corresponde trabajar para poder ver asi las horas
    '       totales que hace. Eso lo podremos comprobar si los marcajes
    '       se producen de dos en dos(entrada salida)
    '       Si el num de horas que le corresponde trabajar se le aplicaran
    '       los valores de retraso y execeso de la empresa
    '       Devolvera 3 incidencias posibles
    '       -Error en marcaje-> Marcajes insuficientes(impares)
    '       -Hora extra
    '       -Retraso
    '       -Tambien podra devolver incidencia 0
    '           si ha trabajado sus horas laborables



'###############################################################
'###############                                          ######
'--------------------------------------------------
'Version 2.0.3      Miercoles 30 de Noviembre de 2001
'--------------------------------------------------
'
' Modificaciones ALZICOOP
' .-
'   Calendario laboral para los dias festivos
'   para contemplar los festivos. Modificaciones en base de Datos
'   y en frmHorarios, añadiendo frmDiasFest
'   tambien se ha modificado Leer horario para que si es festivo
'   de como horas a trabajar 0 y asi sean todas extras
'.-
'   Relizar la importacion del fichero sin que se solapen y ademas
'   poder pedirle que genere los marcajes desde hasta fecha
'.-
'   Añadimos secciones, que iran ligadas a empresas
'   ASi hace falta un mantenimiento como una revision en los trabajadores


'****************************************************************
'****************************************************************
'****************************************************************
'*
'*          Hasta que se unifiquen las versiones
'*          hay una control presencia que contempla
'*          los tikajes nocturnos y uno que no
'*          luego las modificaciones se haran por duplicado
'*          si son de consideracion(obviamente)
'*
'*
'****************************************************************
'****************************************************************
'****************************************************************


'###############################################################
'###############                                          ######
'--------------------------------------------------
'Version 2.0.4      Viernes 28 de Diciembre de 2001
'--------------------------------------------------
'
'
'       Vamos a introducir una opcion para contemplar los
'       descansos no ticados durante la jornada.
'       De momento solo hay dos casos donde quitar x-minutos
'       1ª En Alzicoop: Si trabaja horas extras por la tarde
'           le quitaremos 15'
'       2ª En Belgida: Cualquier persona que tenga un ticaje
'           mas alla de las 6 le quitaremos 15'
'           Parametrizaremos lo de las 6 dentro de empresa
'
'       Solucion:
'       ==============
'       -En HORARIO crearemos dos grupos mas:
'           Almuerzo:   DtoAlm-> Minutos(decimales) con dto
'                       HoraDtoAlm -> Hora a partir de la cual NO
'                                    contabilizaremos el dto
'           Merienda:   DtoMer-> Minutos(decimales) con dto
'                       HoraDtoMer -> Hora a partir de la cual SI
'                                    contabilizaremos el dto
'
'
'
'       -Para la opcion de Belgida (ProcesarMarcaje tipo 3) es
'       facil puesto que se aplicara en cualquier caso. Es decir
'       , a todos se les resta
'
'
'
'       ------------------------------------------------------------------
'       Otra modificación muy importante.
'           .- Se ha divido en DOS PROCESOS el anteriormente
'       denominado IMPORTAR MARCAJES
'       Ahora se descompone en:
'        1.- IMPORTAR FICHERO:  Leera el fichero y lo procesará
'           justo hasta empezar a generar los marcajes. Habrá
'           comprobado las tarjetas y que no existan pendientes
'           para ese trabajador en esa fecha.
'           Destruirá tb el archivo pasandolo a procesados
'        2.- PROCESAR MARCAJES:
'           A partir de la tabla EntradasFichajes genereremos
'           los marcajes. No hara falta( en el caso del TCP3) que
'           Este conectado.
'
'
'
'       ------------------------------------------------------------------
'       Otra modificación
'          .- Parametro dentro de traspaso a aplicaciones ariadna
'       para poder enviar las horas en formato decima o sexagesimal
'
'
'###############################################################
'###############                                          ######
'--------------------------------------------------
'Version 2.0.5      Lunes 14 de enero de 2002
'--------------------------------------------------
'
'       .- Cuando pasamos a procesados contemplamos que ya se haya
'       procesado durante este dia. Luego una funcion determinara
'       el nombre del fichero en procesados. En ella, buscaremos
'       el primer fichero libre de la forma
'           PR &  yymmdd & ". " & numero
'       con numero desde 0 hasta n
'
'
'
'###############################################################
'###############                                          ######
'--------------------------------------------------
'Version 2.0.6      Lunes 21 de Enero de 2002
'--------------------------------------------------
'
'       .-Modificamos la seccion para forzar que unas entradas de
'       fichajes pasen a tener un valor distinto
'       Esto es:
'           Si la produccion se detinene desde las 2 hasta las 3
'           , obviamente no podra haber una marcaje entre esas
'           horas( con la cortesia de por medio)
'           Luego todos los ticajes recogidos entre las 2y10
'           hasta las 2y50 pasaran a ser un ticjae de las 3
'       .-Cuando mostramos los marcajes y el dia mostrado er festivo
'       por horario, es decir, era fin de semana(por ejemplo) entonces
'       aparecera en el nombre del horario FESTIVO
'
'       .-Para no tener que introducir todas las veces el calendario
'       festivo puedo copiarlo:
'           desde un horario distinto
'           desde un año distinto
'           Con lo cual le hare las dos preguntas
'
'
'
'
'###############################################################
'###############                                          ######
'--------------------------------------------------
'Version 2.1.0      Lunes 21 de Enero de 2002
'--------------------------------------------------
' Añadir dos cierres de recordsets en Rectificacion de marcajes
'   Parece ser una version estable.
'
'
'
'###############################################################
'###############                                          ######
'--------------------------------------------------
'Version 2.1.1      Martes 12 de Febrero de 2002
'--------------------------------------------------
'   Informes nuevos para las horas trabajadas
'   Los nuevos informes son HTFecha y HTEmple
'   Solo sirven de momento para los informes en BELGIDA
'
'
'
'###############################################################
'###############                                          ######
'--------------------------------------------------
'Version 2.1.2      Martes 24 de Febrero de 2002
'--------------------------------------------------
'   Para VALSUR
'   Se pretende que las horas extras de los sabados sean festivas
'   Para ello en la configuracion habra una opcion para seleccionar
'   Horas especiales en Sabado
'   Y en frmUnix contemplaremos el mostrar esa informacion




'###############################################################
'###############                                          ######
'------------------------------------------------
'Version 3.0.1         Lunes 8 de Niviembre de 2002
'------------------------------------------------
'   La captacion de marcajes de modo ALZICOOP ha variado
'   las maquinas han sido cambiadas y generan nuevos ficheros





