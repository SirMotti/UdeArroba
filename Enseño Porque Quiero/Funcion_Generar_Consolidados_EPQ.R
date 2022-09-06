#Esteban Motta Ruiz - Auxiliar de Porgramación en Ude@
#Consolidados de informes de Zoom - 23 de marzo de 2022
#Consolidados de participación Enseño Porque Quiero
#esteban.motta@udea.edu.co


direccion_docs <- "D:/Proyectos/Programacion/RFiles/UdeArroba_Informes_EPQ/Reportes_Zoom_EPQ_2022"
nombre_archivo_final <- "D:/Proyectos/Programacion/RFiles/UdeArroba_Informes_EPQ/Consolidado_EPQ_07282022.xlsx"

consolidados_EPQ <- function(d1, nom){
  
  setwd(d1)
  nombres_arch <- list.files(pattern = "*.csv")
  
  reportes_zoom <- lapply(list.files(pattern = "*.csv"),
                          read.csv,
                          encoding = "UTF-8",
                          header = FALSE)
  
  lista_Ins_Vin <- list()
  lista_Ins_Und <- list()
  lista_Ast_Vin <- list()
  lista_Ast_Und <- list()
  
  for (p in 1:length(reportes_zoom)) {
       
       Rep_Ins_file_dt <- subset.data.frame(reportes_zoom[[p]][-1, -c(6, 8, 9, 10, 11, 14, 15)])
        
       library(stringr)
       for (i in 1:length(Rep_Ins_file_dt)) {
            Rep_Ins_file_dt[i] <- str_replace_all(Rep_Ins_file_dt[, i], " ", " ")
       }
        
       colnom <- c("Asist", "Reg.Nombre", "Nombre", "Apellido", "Correo", "Estado", "Vinculacion", "Unidad")
       names(Rep_Ins_file_dt) <- colnom
       
       Rep_Ins_file_dt$Asist[Rep_Ins_file_dt$Asist == "Sí"] <- 1
       Rep_Ins_file_dt$Asist[Rep_Ins_file_dt$Asist == "No"] <- 0
       
       Rep_Ins_file_fn <- Rep_Ins_file_dt[!duplicated(Rep_Ins_file_dt$Correo), ]
       
       Rep_Ins_file_fn$Asist <- as.numeric(Rep_Ins_file_fn$Asist)
       
       Ins_total <- subset(Rep_Ins_file_fn[-c(1, 2), ],
                           select = c(Asist, Vinculacion, Unidad))
       Ins_total$Asist <- 1
       
       lista_Ins_Vin[[p]] <- aggregate(Asist~Vinculacion,
                                       data = Ins_total,
                                       FUN = sum)
       lista_Ins_Und[[p]] <- aggregate(Asist~Unidad,
                                       data = Ins_total,
                                       FUN = sum)
       
       Rep_Ast_file_fn <- Rep_Ins_file_fn[Rep_Ins_file_fn$Asist > 0, ]
       
       lista_Ast_Vin[[p]] <- aggregate(Asist~Vinculacion,
                                       data = Rep_Ast_file_fn[-c(1, 2), ],
                                       FUN = sum)
       lista_Ast_Und[[p]] <- aggregate(Asist~Unidad,
                                       data = Rep_Ast_file_fn[-c(1, 2), ],
                                       FUN = sum)
    
  }
  
  lista_Und <- list(lista_Ins_Und,
                    lista_Ast_Und)
  
  lista_Vin <- list(lista_Ins_Vin,
                    lista_Ast_Vin)
  
  lista_Und_cons <- list()
  lista_Vin_cons <- list()
  
  for(g in 1:length(lista_Und)){
      
      lista_Und_cons[[g]]<- Reduce(function(x, y) merge(x,
                                                        y,
                                                        by = "Unidad",
                                                        all = TRUE),
                                   lista_Und[[g]])
    
  }
  
  nombres_col_con_Und <- c("Unidad", nombres_arch)
  
  names(lista_Und_cons[[1]]) <- nombres_col_con_Und
  names(lista_Und_cons[[2]]) <- nombres_col_con_Und
  
  lista_Und_cons<- rapply(lista_Und_cons,
                          f = function(x) ifelse(test = is.na(x),
                                                 yes = 0,
                                                 no = x),
                          how = "replace" )
  
  
  for(h in 1:length(lista_Vin)){
      
      lista_Vin_cons[[h]]<- Reduce(function(x, y) merge(x,
                                                        y,
                                                        by = "Vinculacion",
                                                        all = TRUE),
                                   lista_Vin[[h]])
      
  }
  
  nombres_col_con_Vin <- c("Vinculacion", nombres_arch)
  
  names(lista_Vin_cons[[1]]) <- nombres_col_con_Vin
  names(lista_Vin_cons[[2]]) <- nombres_col_con_Vin
  
  lista_Vin_cons<- rapply(lista_Vin_cons,
                          f = function(x) ifelse(test = is.na(x),
                                                 yes = 0,
                                                 no = x),
                          how = "replace" )
  
  
  library(xlsx)
  
  write.xlsx(lista_Und_cons[[1]],
             nom,
             sheetName = "Consolidado_Inscritos_Unidad",
             col.names = TRUE,
             row.names = FALSE,
             append = TRUE)
  
  write.xlsx(lista_Und_cons[[2]],
             nom,
             sheetName = "Consolidado_Asistentes_Unidad",
             col.names = TRUE,
             row.names = FALSE,
             append = TRUE)
  
  write.xlsx(lista_Vin_cons[[1]],
             nom,
             sheetName = "Consolidado_Inscritos_Vinculacion",
             col.names = TRUE,
             row.names = FALSE,
             append = TRUE)
  
  write.xlsx(lista_Vin_cons[[2]],
             nom,
             sheetName = "Consolidado_Asistentes_Vinculacion",
             col.names = TRUE,
             row.names = FALSE,
             append = TRUE)
  
}

consolidados_EPQ(d1 = direccion_docs,
                 nom = nombre_archivo_final)




