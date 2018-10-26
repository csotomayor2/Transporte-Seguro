library(readxl)
    
    
    velocidad <- read_excel("C:/Users/User/Documents/Rstudio/TransSeguro/velocidad2.xlsx")
    velocidad<- as.data.frame(velocidad)
    velocidad$PROVINCIA<- as.factor(velocidad$PROVINCIA)
    
    velocidad$LONGITUDE<-as.numeric(velocidad$LONGITUDE)
    velocidad$LATITUDE<- as.numeric(velocidad$LATITUDE)
    
    
    library(xlsx)
    
    write.xlsx(velocidad, file="C:/Users/User/Desktop/automatico/velo.xlsx")
    #==========================================================
    
    exportar <- function(datos, archivo){        
      wb <- createWorkbook(type="xlsx")
      
      # Estilos de celdas
      # Estilos de titulos y subtitulos
      titulo <- CellStyle(wb)+ Font(wb,  heightInPoints=16, isBold=TRUE)
      
      subtitulo <- CellStyle(wb) + Font(wb,  heightInPoints=12,
                                        isItalic=TRUE, isBold=FALSE)
      
      # Estilo de tablas
      filas <- CellStyle(wb) + Font(wb, isBold=TRUE)
      
      columnas <- CellStyle(wb) + Font(wb, isBold=TRUE) +
        Alignment(vertical="VERTICAL_CENTER",wrapText=TRUE, horizontal="ALIGN_CENTER") +
        Border(color="black", position=c("TOP", "BOTTOM"), 
               pen=c("BORDER_THICK", "BORDER_THICK"))+Fill(foregroundColor = "lightblue", pattern = "SOLID_FOREGROUND")
      
      # Crear una hoja
      sheet <- createSheet(wb, sheetName = "Información - R Users Group - Ecuador")
      
      # Funcion linea (agregar texto)
      linea<-function(sheet, rowIndex, title, titleStyle){
        rows <- createRow(sheet, rowIndex=rowIndex)
        sheetTitle <- createCell(rows, colIndex=1)
        setCellValue(sheetTitle[[1,1]], title)
        setCellStyle(sheetTitle[[1,1]], titleStyle)
      }
      
      # Agregamos titulos,  subtitulos, etc.
      linea(sheet, rowIndex=8, 
            title=paste("Fecha:", format(Sys.Date(), format="%Y/%m/%d")),
            titleStyle = subtitulo)
      
      linea(sheet, rowIndex=9, 
            title="Elaborado por: R Users Group - Ecuador",
            titleStyle = subtitulo)
      
      linea(sheet, rowIndex=11, 
            paste("Información de prueba"),
            titleStyle = titulo)
      
      # Tablas
      addDataFrame(datos,
                   sheet, startRow=13, startColumn=1,
                   colnamesStyle = columnas,
                   rownamesStyle = filas,
                   row.names = F)
      
      # Ancho de columnas
      setColumnWidth(sheet, colIndex=c(1:ncol(datos)), colWidth=15)
      
      # Imagen    Cambia la ruta por la que de tu imagen
      addPicture("C:/Users/User/Desktop/logo-ant.png", sheet, scale=0.75, startRow = 1, startColumn = 9)
      addPicture("C:/Users/User/Desktop/logo2.jpeg", sheet, scale=0.75, startRow = 1, startColumn = 1)
      
      # Guardar
      saveWorkbook(wb, archivo)
    }
    
    exportar(mtcars, "R Users Group - Ecuador - mtcars.xlsx")  
    
    
   
    