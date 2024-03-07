# Esteban Motta Ruiz - Auxiliar de Porgramación en Ude@
# Edición de informes de Zoom - 6 de Marzo de 2024
# Informes de Aprender desde Casa
# esteban.motta@duea.edu.co

direccion_reg <- "/reg.csv"
direccion_par <- "/par.csv"

nombre_arch <- "/ejemplo.xlsx"

Editar.Reporte.ADC <- function(d1, d2, nom){
  
  dir1 <- read.csv(d1,
                   encoding = "UTF-8")
  dir2 <- read.csv(d2,
                   encoding = "UTF-8")
  
  library(stringr)
  
  for (i in 1:length(dir1)) {
       dir1[i] <- str_replace_all(string = dir1[, i],
                                  pattern = " ",
                                  replacement = " ")
  }
  
  for (i in 1:length(dir2)) {
       dir2[i] <- str_replace_all(string = dir2[, i],
                                  pattern = " ",
                                  replacement = " ")
  }
  
  asist <- rep(x = 1,
               times = length(dir1[, 1]))
  ins <- subset.data.frame(dir1[, c(1, 2, 3, 6, 7, 8)])
  ins_dt <- cbind(ins, asist)
  
  colnom1 <- c("Nombre",
               "Apellido",
               "Correo",
               "Programa",
               "Dependencia",
               "Vinculacion",
               "Asist")
  colnames(ins_dt) <- colnom1
  
  ins_dt$Correo <- tolower(ins_dt$Correo)
  
  ast <- subset.data.frame(dir2[, c(1, 2)])
  
  colnom2 <- c("Nombre",
               "Correo")
  colnames(ast) <- colnom2
  
  ast_dt <- merge(x = ins_dt,
                  y = ast,
                  by = "Correo")

  ins_prog <- aggregate(Asist~Programa,
                        data = ins_dt,
                        FUN = sum)
  ins_dep <- aggregate(Asist~Dependencia,
                       data = ins_dt,
                       FUN = sum)
  ins_vin <- aggregate(Asist~Vinculacion,
                       data = ins_dt,
                       FUN = sum)
  
  ast_prog <- aggregate(Asist~Programa,
                        data = ast_dt,
                        FUN = sum)
  ast_dep <- aggregate(Asist~Dependencia,
                       data = ast_dt,
                       FUN = sum)
  ast_vin <- aggregate(Asist~Vinculacion,
                       data = ast_dt,
                       FUN = sum)
  
  ast_dt2 <- ast_dt[, 1:7]
  colnames(ast_dt2)[2] <- "Nombre"
  ast_dt2$Asist <- "1"
  
  no_ast <- data.frame(Correo = setdiff(ins_dt$Correo,
                                        ast_dt2$Correo))
  no_ast_dt <- merge(x = ins_dt,
                     y = no_ast,
                     by.x = "Correo",
                     by.y = "Correo")
  
  no_ast_dt$Asist <- "0"
  
  general <- rbind(ast_dt2, no_ast_dt)
  general <- general[c(7, 2, 3, 1, 4, 5, 6)]
  general$Asist <- as.numeric(general$Asist)
  
  library(openxlsx)
  
  finales <- list("Reporte General" = general,
                  "Ins Programa" = ins_prog,
                  "Ins Dependencia" = ins_dep,
                  "Ins Vinculacion" = ins_vin,
                  "Ast Programa" = ast_prog,
                  "Ast Dependencia" = ast_dep,
                  "Ast Vinculacion" = ast_vin)
  
  write.xlsx(finales,
             nom,
             colNames = TRUE,
             rowNames = FALSE,
             overwrite = TRUE)
  
}

Editar.Reporte.ADC(d1 = direccion_reg,
                   d2 = direccion_par,
                   nom = nombre_arch)
