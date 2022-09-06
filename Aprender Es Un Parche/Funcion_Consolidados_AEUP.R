#Esteban Motta Ruiz - Auxiliar de Porgramación en Ude@
#Consolidados de informes de Zoom - 18 de mayo de 2022
#Consolidados de participación Enseño Porque Quiero
#esteban.motta@udea.edu.co


direccion_docs_reg <- "D:/Proyectos/Programacion/RFiles/UdeArroba_Informes_ADC/Archivos2022/reg"
direccion_docs_par <- "D:/Proyectos/Programacion/RFiles/UdeArroba_Informes_ADC/Archivos2022/par"
nombre_archivo_final <- "D:/Proyectos/Programacion/RFiles/UdeArroba_Informes_ADC/Archivos2022/Merge_ADC_2022.xlsx"

consolidados <- function(d1, d2, nom){
  
  setwd(d1)
  
  nombres_arch_reg <- list.files(pattern = "*.csv")
  
  reportes_zoom_reg <- lapply(list.files(pattern = "*.csv"),
                              FUN = read.csv,
                              encoding = "UTF-8",
                              header = TRUE)
  
  
  setwd(d2)
  
  nombres_arch_par <- list.files(pattern = "*.csv")
  
  reportes_zoom_par <- lapply(list.files(pattern = "*.csv"),
                              FUN = read.csv,
                              encoding = "UTF-8",
                              header = TRUE)
  
  lista_Ins_Prog <- list()
  lista_Ins_Dep <- list()
  lista_Ins_Vin <- list()
  lista_Ast_Prog <- list()
  lista_Ast_Dep <- list()
  lista_Ast_Vin <- list()
  
  library(stringr)
  
  for (g in 1:length(reportes_zoom_reg)) {
       
       for (i in 1:length(reportes_zoom_reg[[g]])) {
            reportes_zoom_reg[[g]][, i] <- str_replace_all(reportes_zoom_reg[[g]][, i], " ", " ")
       
       }
    
  }
  
  for (h in 1:length(reportes_zoom_par)) {
       
       for (f in 1:length(reportes_zoom_par[[h]])) {
            reportes_zoom_par[[h]][, f] <- str_replace_all(reportes_zoom_par[[h]][, f], " ", " ")
       
      }
      
  }
  
  ins_dt <- list()
  ast <- list()
  ast_dt <- list()
  colnom2 <- c("Nombre", "Correo")
  
  for (n in 1:length(reportes_zoom_par)) {
       
       ast[[n]] <- subset.data.frame(reportes_zoom_par[[n]][, c(1, 2)])
       colnames(ast[[n]]) <- colnom2
       
  }
  
  for (l in 1:length(reportes_zoom_reg)) {
       
       asist <- rep(1, times = length(reportes_zoom_reg[[l]][, 1]))
       ins <- subset.data.frame(reportes_zoom_reg[[l]][, c(1, 2, 3, 6, 7, 8)])
       ins_dt[[l]] <- cbind(ins, asist)
       
       colnom1 <- c("Nombre", "Apellido", "Correo", "Programa", "Dependencia",
                    "Vinculacion", "Asist")
       colnames(ins_dt[[l]]) <- colnom1
       
       ins_dt[[l]]$Correo <- tolower(ins_dt[[l]]$Correo)
       
       lista_Ins_Prog[[l]] <- aggregate(Asist~Programa,
                                        data = ins_dt[[l]],
                                        FUN = sum)
       lista_Ins_Dep[[l]] <- aggregate(Asist~Dependencia,
                                       data = ins_dt[[l]],
                                       FUN = sum)
       lista_Ins_Vin[[l]] <- aggregate(Asist~Vinculacion,
                                       data = ins_dt[[l]],
                                       FUN = sum)
       
  }
  
  for (k in 1:length(reportes_zoom_reg)) {
       
       
       ast_dt[[k]] <- merge(ins_dt[[k]], ast[[k]], by = "Correo")
       
       lista_Ast_Prog[[k]] <- aggregate(Asist~Programa,
                                        data = ast_dt[[k]],
                                        FUN = sum)
       lista_Ast_Dep[[k]] <- aggregate(Asist~Dependencia,
                                       data = ast_dt[[k]],
                                       FUN = sum)
       lista_Ast_Vin[[k]]  <- aggregate(Asist~Vinculacion,
                                        data = ast_dt[[k]],
                                        FUN = sum)
       
  }
  
  
  lista_Prog <- list(lista_Ins_Prog,
                     lista_Ast_Prog)
  
  lista_Dep <- list(lista_Ins_Dep,
                    lista_Ast_Dep)
  
  lista_Vin <- list(lista_Ins_Vin,
                    lista_Ast_Vin)
  
  lista_Prog_cons <- list()
  lista_Dep_cons <- list()
  lista_Vin_cons <- list()
  
  
  for(j in 1:length(lista_Prog)){
    
    lista_Prog_cons[[j]]<- Reduce(function(x, y) merge(x,
                                                       y,
                                                       by = "Programa",
                                                       all = TRUE),
                                  lista_Prog[[j]])
    
  }
  
  nombres_col_con_Prog <- c("Programa", nombres_arch_reg)
  
  names(lista_Prog_cons[[1]]) <- nombres_col_con_Prog
  names(lista_Prog_cons[[2]]) <- nombres_col_con_Prog
  
  lista_Prog_cons<- rapply(lista_Prog_cons,
                           f = function(x) ifelse(test = is.na(x),
                                                 yes = 0,
                                                 no = x),
                           how = "replace" )
  
  
  for(m in 1:length(lista_Dep)){
      
      lista_Dep_cons[[m]]<- Reduce(function(x, y) merge(x,
                                                        y,
                                                        by = "Dependencia",
                                                        all = TRUE),
                                   lista_Dep[[m]])
      
  }
  
  nombres_col_con_Dep <- c("Dependencia", nombres_arch_reg)
  
  names(lista_Dep_cons[[1]]) <- nombres_col_con_Dep
  names(lista_Dep_cons[[2]]) <- nombres_col_con_Dep
  
  lista_Dep_cons<- rapply(lista_Dep_cons,
                          f = function(x) ifelse(test = is.na(x),
                                                 yes = 0,
                                                 no = x),
                          how = "replace" )
  
  
  for(p in 1:length(lista_Dep)){
      
      lista_Vin_cons[[p]]<- Reduce(function(x, y) merge(x,
                                                        y,
                                                        by = "Vinculacion",
                                                        all = TRUE),
                                   lista_Vin[[p]])
      
  }
  
  nombres_col_con_Vin <- c("Vinculacion", nombres_arch_reg)
  
  names(lista_Vin_cons[[1]]) <- nombres_col_con_Vin
  names(lista_Vin_cons[[2]]) <- nombres_col_con_Vin
  
  lista_Vin_cons<- rapply(lista_Vin_cons,
                          f = function(x) ifelse(test = is.na(x),
                                                 yes = 0,
                                                 no = x),
                          how = "replace" )
  
  
  library(xlsx)
  
  write.xlsx(lista_Prog_cons[[1]],
             nom,
             sheetName = "Consolidado_Inscritos_Programa",
             col.names = TRUE,
             row.names = FALSE,
             append = TRUE)
  
  write.xlsx(lista_Prog_cons[[2]],
             nom,
             sheetName = "Consolidado_Asistentes_Programa",
             col.names = TRUE,
             row.names = FALSE,
             append = TRUE)
  
  write.xlsx(lista_Dep_cons[[1]],
             nom,
             sheetName = "Consolidado_Inscritos_Dependencia",
             col.names = TRUE,
             row.names = FALSE,
             append = TRUE)
  
  write.xlsx(lista_Dep_cons[[2]],
             nom,
             sheetName = "Consolidado_Asistentes_Dependencia",
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

consolidados(d1 = direccion_docs_reg,
             d2 = direccion_docs_par,
             nom = nombre_archivo_final)
