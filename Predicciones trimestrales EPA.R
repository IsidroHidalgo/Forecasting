# -----------------------------------------------------------------------------
# (c) Observatorio Regional de Empleo. Junta de Comunidades de Castilla-La Mancha
# (c) Isidro Hidalgo Arellano (ihidalgo@jccm.es)
# -----------------------------------------------------------------------------

# Borramos los objetos en memoria
rm(list=ls(all.names = TRUE))

# Predicci�n series trimestrales:
inicial = 7
final = 12
nseries = final - inicial + 1

# Definimos variables globales
lista_modelos = c("HyndmanARIMA", "STL", "Hyndman", "UComp", "TramoSeats", "x13Seats")

#Cargamos las librer�as si no est�n disponibles
library(openxlsx) # extracci�n de datos desde Excel
library(forecast) # m�todos ETS ("Hyndman") y ARIMA (HyndmanARIMA)
library(UComp) # m�todo "UComp" (componentes no-observables)
library(RJDemetra) # m�todos "TRAMO-SEATS" y "X13-ARIMA-SEATS"

# Especificamos el directorio de trabajo
setwd("\\\\jclm.es/PROS/SC/OBSERVATORIOEMPLEO/An�lisis/Previsiones/Trimestrales - EPA")
ruta <- getwd()

# Creamos el fichero Excel de salida
ExcelSalida <- paste0(ruta,"/SalidaParoAfiliaci�n.xlsx")

# Cargamos las series desde el fichero EXCEL donde tenemos las series
excelfile <- read.xlsx("../Series.xlsx")
longitud <- max(excelfile[1, inicial:final])
datos <- excelfile[1:(longitud+6),inicial:final]

# Obtenemos el n�mero de series cargadas
n <- ncol(datos)

series_modelos <- function(nSerie, modelos = lista_modelos){
### Funci�n para seleccionar las series que queremos predecir y los modelos a utilizar
  
  print(names(datos)[nSerie])  
  
  # Creamos la serie eliminando los valores ausentes del final, si los hay
  columna <- na.omit(datos[, nSerie])
  filas <- length(columna)
  inicio<-columna[1]
  frecuencia<-columna[4]
  serie <- ts(columna[-c(1:5)],
              start = c(columna[2],columna[3]),
              frequency = frecuencia)
  nombre_serie <- colnames(datos)[nSerie]
  
  
  # Preparamos la matriz para guardar los resultados
  tandas <- matrix(NA, frecuencia, length(modelos))
  colnames(tandas) <- modelos
  
  for (modelo in modelos) {
    
    # Lanzamos la predicci�n...
    
    # ...con el m�todo STL
    if ("STL" == modelo){
      forecast <- stlf(serie)
      tandas[, modelo]<-forecast$mean[1:frecuencia]
    } # STL
    
    # ...con el forecast de Hyndman  
    if ("Hyndman" == modelo){
      forecast <- forecast(serie)
      tandas[, modelo]<-forecast$mean[1:frecuencia]
    } # Hyndman
    
    # ...con el TramoSeats
    if ("TramoSeats" == modelo){
      forecast <- tramoseats(serie, spec = "RSAfull")
      tandas[, modelo]<-forecast$regarima$forecast[1:frecuencia]
    } # TramoSeats
    
    # ...con el HyndmanARIMA
    if ("HyndmanARIMA" == modelo){
      forecast <- forecast(auto.arima(serie))
      tandas[, modelo]<-forecast$mean[1:frecuencia]
    } # HyndmanARIMA
    
    # ...con el UComp
    if ("UComp" == modelo){
      forecast <- UCmodel(serie, verbose = FALSE)
      tandas[, modelo]<-forecast$yFor[1:frecuencia]
    } # UComp
    
    # ...con el x13Seats
    if ("x13Seats" == modelo){
      forecast <- x13(serie, spec = "RSA5c")
      tandas[, modelo]<-forecast$regarima$forecast[1:frecuencia]
    } # x13Seats
    
  } # Cerramos modelo
  
  mediana <- apply(tandas, 1, median)
  
  return(mediana)
  
} # Cerramos funci�n

# Predicci�n:
predicciones <- sapply(1:nseries, series_modelos)
predicciones <- as.data.frame(predicciones)
names(predicciones) <- names(datos)

# Escribimos en el fichero Excel
write.xlsx(predicciones, file = ExcelSalida, sheet = "Predicciones")

print("Las predicciones se encuentran en el fichero:")
print("\\\\jclm.es/PROS/SC/OBSERVATORIOEMPLEO/An�lisis/Previsiones/Trimestrales - EPA/SalidaEPA.xlsx")
