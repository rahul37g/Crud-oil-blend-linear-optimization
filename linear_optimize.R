##########################################################################################################
#************************** Blending oil linear program optimization ************************************#
##########################################################################################################

library(Rglpk) ## Linear and Mixed Integer Programming Solver Using GLPK
library(tidyverse) ## R package for ggpolot2, dplyr, etc.
library(plotly) ## r package for hover effects
library(pracma) ## Practical Numerical Math Functions (like, ones, eye, zeros, etc)
library(xlsx) # read, write, format Excel.

rm(list = ls()) # clear gobal environment

gasoline <- read.xlsx("oil_blend_lp.xlsx",sheetName = "oil_blend_lp", startRow = 1, endRow = 30)
gasoline <- subset(gasoline, Select == "YES"); gasoline$Select <- NULL

raw_material <- read.xlsx("oil_blend_lp.xlsx",sheetName = "oil_blend_lp", startRow = 31, endRow = 69)
raw_material <- subset(raw_material, Select == "YES"); raw_material$Select <- NULL

obj  <- c(rep(0,length(raw_material$Octane_Rating)*length(gasoline$Gasoline)),
          raw_material$Avalibality,
          gasoline$Price)

## set constraint martix
## Constraint matrix is divided into 3 parts:

## 1. Part 1: matrix constraint for oil qualitiy.
## Part 1 (a):
fuel_quality1=c() 
for (row in 1:length(gasoline$Gasoline)) {
  for (col in 1:(length(gasoline$Gasoline)+1)) {
    if(row == col){
      fuel_quality1<-append(fuel_quality1, -raw_material$Octane_Rating)
    }
    else{
      fuel_quality1<-append(fuel_quality1, rep(0,length(raw_material$Octane_Rating)))
    }
  }
} 
fuel_quality1 <- matrix(fuel_quality1, nrow = length(gasoline$Gasoline), 
                        ncol = length(raw_material$Octane_Rating)*(length(gasoline$Gasoline)+1), 
                        byrow = TRUE)
## Part 1 (b):
fuel_quality2=c()
for (row in 1:length(gasoline$Gasoline)) {
  for (col in 1:length(gasoline$Gasoline)) {
    if(row == col){
      fuel_quality2<-append(fuel_quality2, gasoline$Octane_Rating[row])
    }
    else{
      fuel_quality2<-append(fuel_quality2, 0)
    }
  }
} 
fuel_quality2 <- matrix(fuel_quality2, nrow = length(gasoline$Gasoline), 
                        ncol = length(gasoline$Octane_Rating), byrow = TRUE)

## part 1 final matrix for fuel quality
constraint1 <- cbind(fuel_quality1,fuel_quality2)

## 2. Part 2: matrix constraint for oil availability.
constraint2 <- matrix(c(rep(eye(length(raw_material$Octane_Rating)),length(gasoline$Gasoline)+1), 
                        zeros(length(raw_material$Octane_Rating),length(gasoline$Octane_Rating))),
                      nrow = length(raw_material$Octane_Rating), 
                      ncol = (length(raw_material$Octane_Rating) * (length(gasoline$Gasoline)+1)) + 
                        length(gasoline$Gasoline))

## 3. Part 3: matrix constraint for oil quantity.
#* Part 3 (a):
fuel_quantity1=c()
for (row in 1:length(gasoline$Gasoline)) {
  for (col in 1:(length(gasoline$Gasoline)+1)) {
    if(row == col){
      fuel_quantity1<-append(fuel_quantity1,rep(1,length(raw_material$Octane_Rating)))
    }
    else{
      fuel_quantity1<-append(fuel_quantity1,rep(0,length(raw_material$Octane_Rating)))
    }
  }
} 
fuel_quantity1 <- matrix(fuel_quantity1, nrow = length(gasoline$Gasoline), 
                         ncol = length(raw_material$Octane_Rating)*(length(gasoline$Gasoline)+1), 
                         byrow = TRUE)

## Part 3 (b):
fuel_quantity2=c()
for (row in 1:length(gasoline$Gasoline)) {
  for (col in 1:length(gasoline$Gasoline)) {
    if(row == col){
      fuel_quantity2<-append(fuel_quantity2, -1)
    }
    else{
      fuel_quantity2<-append(fuel_quantity2, 0)
    }
  }
} 
fuel_quantity2 <- matrix(fuel_quantity2, nrow = length(gasoline$Gasoline), 
                         ncol = length(gasoline$Octane_Rating), byrow = TRUE)

## part 3 final matrix for fuel quantity
constraint3 <- cbind(fuel_quantity1, fuel_quantity2)

## Final matrix constraint required for optimization.
mat <- rbind(constraint1, 
             constraint2, 
             constraint3)

## RHS for the constraints
rhs <- c(zeros(1,length(gasoline$Price)), 
         raw_material$Avalibality, 
         zeros(1, length(gasoline$Octane_Rating)))

## Constraints direction
dir  <- c(rep(">=", length(gasoline$Price)),
          rep("==", length(raw_material$Avalibality)), 
          rep("==", length(gasoline$Octane_Rating)))

## solve using Rglpk package for linear programming solver
optimum <- Rglpk_solve_LP(obj, mat, dir, rhs, max = FALSE)

## Find the optimal solution
optimum$solution # Display the optimum values for x1 and x2 and others
optimum$optimum # Check the value of objective function at optimal point

# set fuel_used variable to store the result of gasoline used, obtain from optimum$solution
fuel_used <- data.frame(gasoline$Gasoline, tail(optimum$solution, length(gasoline$Gasoline)))
colnames(fuel_used) <- c("Fuel", "Used_Fuel")

# set revenue variable to store the result of objective function at optimal point
revenue <- data.frame("Revenue"=format(sum(tail(optimum$solution,length(gasoline$Gasoline))*gasoline$Price),
                                       big.mark=",",scientific=FALSE))
# revenue <- data.frame("Revenue" = format(optimum$optimum,big.mark=",",scientific=FALSE))

# set df blend_dist to save on .csv format, which contains the final result of blending process
blend_dist <- data.frame(matrix(
  c(raw_material$Octane_Rating, 
    as.numeric(format(round(optimum$solution[1:(length(raw_material$Octane_Rating)*length(gasoline$Gasoline))])))), 
  nrow = length(raw_material$Octane_Rating)))
if(length(gasoline$Gasoline) > 1){
  blend_dist <- cbind(blend_dist, rowSums(blend_dist[, -1]))
  colnames(blend_dist) <- c("Octane_Rating",paste("Fuel_",as.character(gasoline$Gasoline), sep = ""), "Total")
} else {
  colnames(blend_dist) <- c("Octane_Rating",paste("Fuel_",as.character(gasoline$Gasoline), sep = ""))
}

# Save/write data to excel sheet.
wb <- loadWorkbook("oil_blend_lp.xlsx")
sheets <- getSheets(wb)

if ("blend_result" %in% names(sheets)) {
  removeSheet(wb, sheetName="blend_result")
  mysheet <- createSheet(wb, sheetName="blend_result")
} else {
  mysheet <- createSheet(wb, sheetName="blend_result")
}

addDataFrame(blend_dist, mysheet, row.names = FALSE, col.names = TRUE)
addDataFrame(fuel_used, mysheet, row.names = FALSE, col.names = TRUE, startRow = length(blend_dist$Octane_Rating)+3)
addDataFrame(revenue, mysheet, row.names = FALSE, col.names = TRUE, 
             startRow = (length(blend_dist$Octane_Rating)+1) + (length(fuel_used$Fuel)+1)+3)
saveWorkbook(wb, "oil_blend_lp.xlsx")

# created df blend_graph according to ggplot2 requirements 
blend_graph <- data.frame(rep(gasoline$Gasoline, each = length(raw_material$Avalibality)), 
                          as.numeric(rep(raw_material$Octane_Rating,time=length(gasoline$Gasoline))),
                          as.numeric(format(round(optimum$solution[1:(length(raw_material$Octane_Rating)*length(gasoline$Gasoline))]))))
colnames(blend_graph) <- c("Fuel","Octane_Rating","Used_Fuel")

## plot ggplot2 graph for blending process
ggplotly(ggplot(blend_graph, aes(Octane_Rating, Used_Fuel, fill=Fuel)) +
  theme(panel.grid.major = element_line(color = "#a09d9a")) +
  # facet_wrap(~ Fuel) +
  geom_bar(stat='identity', position = 'dodge') +
  labs(x = "Ocatane Rating",y = "Fuel Used for Blending",
       title = "Blending") +
  scale_x_discrete(limits=raw_material$Octane_Rating)) 

# save plot in png format
ggsave("blend_graph.png")

# Display the final result of blending process along with the revenue
print(blend_dist, print.gap=2, row.names=F)
print(fuel_used, print.gap=2, row.names=F)
print(revenue, print.gap=2, row.names=F)
