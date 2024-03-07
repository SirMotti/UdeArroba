# Esteban Motta Ruiz - Auxiliar de Porgramación en Ude@
# Edición de informes de Zoom - 6 de Marzo de 2024
# Informes de Enseño Porque Quiero
# esteban.motta@udea.edu.co

direccion <- "/ast.csv"

nombre_arch <- "/reporte.xlsx"

Editar.Reporte.EPQ <- function(dir, nom){
  
  Rep_Ast <- read.csv(file = dir,
                      encoding = "UTF-8",
                      header = FALSE)
  
  Rep_Ast_dt <- subset.data.frame(Rep_Ast[-1, -c(6, 8, 9, 10, 11, 14, 15)])
  
  library(stringr)
  
  for (i in 1:length(Rep_Ast_dt)) {
       Rep_Ast_dt[i] <- str_replace_all(Rep_Ast_dt[, i], " ", " ")
  }
  
  colnom <- c("Asist",
              "Reg.Nombre",
              "Nombre",
              "Apellido",
              "Correo",
              "Estado",
              "Vinculacion",
              "Unidad")
  
  names(Rep_Ast_dt) <- colnom
  
  Rep_Ast_dt$Asist[Rep_Ast_dt$Asist == "Sí"] <- 1
  Rep_Ast_dt$Asist[Rep_Ast_dt$Asist == "No"] <- 0
  
  Rep_Ast_fn <- Rep_Ast_dt[!duplicated(Rep_Ast_dt$Correo), ]
  
  Rep_Ast_fn$Asist <- as.numeric(Rep_Ast_fn$Asist)
  
  Ins_total <- subset(Rep_Ast_fn,
                      select = c(Asist, Vinculacion, Unidad))
  
  Ins_total$Asist <- 1
  
  Ins_Vin <- aggregate(Asist~Vinculacion,
                       data = Ins_total,
                       FUN = sum)
  Ins_Und <- aggregate(Asist~Unidad,
                       data = Ins_total,
                       FUN = sum)
  
  Rep_Ast_fn_1 <- Rep_Ast_fn[Rep_Ast_fn$Asist > 0, ]
  
  Ast_Vin <- aggregate(Asist~Vinculacion,
                       data = Rep_Ast_fn_1,
                       FUN = sum)
  Ast_Und <- aggregate(Asist~Unidad,
                       data = Rep_Ast_fn_1,
                       FUN = sum)
  
  library(openxlsx)
  
  finales <- list("Reporte General" = Rep_Ast_fn,
                  "Inscritos por Unidad" = Ins_Und[-1, ],
                  "Inscritos por Vinculación" = Ins_Vin[-1, ],
                  "Asistentes por Unidad" = Ast_Und[-1, ],
                  "Asistentes por Vinculación" = Ast_Vin[-1, ])
  
  write.xlsx(finales,
             nom,
             colNames = TRUE,
             rowNames = FALSE,
             overwrite = TRUE)
  
  
}

Editar.Reporte.EPQ(dir = direccion,
                   nom = nombre_arch)
