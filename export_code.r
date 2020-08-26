#Rscript
#Date: 2020-08-26
#Author: Yunlong Ma
#@e-mail: glb-biotech@zju.edu.cn

#Note: this export R package is suitable for R.4.0.0

getwd()
# install export R package
if(!require("officer")){install.packages("officer")}
if(!require("rvg")){install.packages("rvg")}
if(!require("openxlsx")){install.packages("openxlsx")}
if(!require("ggplot2")){install.packages("ggplot2")}
if(!require("flextable")){install.packages("flextable")}
if(!require("xtable")){install.packages("xtable")}
if(!require("rgl")){install.packages("rgl")}
if(!require("stargazer")){install.packages("stargazer")}
if(!require("tikzDevice")){install.packages("tikzDevice")}
if(!require("xml2")){install.packages("xml2")}
if(!require("broom")){install.packages("broom")}
if(!require("devtools")){install.packages("devtools")}
#install.packages("officer")
#install.packages("rvg")
#install.packages("openxlsx")
#install.packages("ggplot2")
#install.packages("flextable")
#install.packages("xtable")
#install.packages("rgl")
#install.packages("stargazer")
#install.packages("tikzDevice")
#install.packages("xml2")
#install.packages("broom")
#install.packages("devtools")
#if(!require("export")){install.packages("export")}

#This is already installed.
#devtools::install_github("tomwenseleers/export")
library(export)



?graph2ppt
?graph2doc
?graph2svg
?graph2png
?table2ppt
?table2tex
?table2excel
?table2doc
?table2html

## export of ggplot2 plot
library(ggplot2)
qplot(Sepal.Length, Petal.Length, data = iris, color = Species, 
      size = Petal.Width, alpha = I(0.7))
# export to Powerpoint      
graph2ppt()      
graph2ppt(file="ggplot2_plot.pptx", aspectr=1.7)
# add 2nd slide with same graph 9 inches wide and A4 aspect ratio
graph2ppt(file="ggplot2_plot.pptx", width=9, aspectr=sqrt(2), append=TRUE) 
# add 3d slide with same graph with fixed width & height
graph2ppt(file="ggplot2_plot.pptx", width=6, height=5, append=TRUE) 



# export to Word
graph2doc()
# export to bitmap or vector formats
graph2svg()
graph2png()
graph2tif()
graph2jpg()



## export of aov Anova output
fit=aov(yield ~ block + N * P + K, npk)
x=summary(fit)
# export to Powerpoint
table2ppt(x=x)
table2ppt(x=x,file="table_aov.pptx")
table2ppt(x=x,file="table_aov.pptx",digits=4,append=TRUE)
table2ppt(x=x,file="table_aov.pptx",digits=4,digitspvals=1,
          font="Times New Roman",pointsize=16,append=TRUE)
# export to Word
table2doc(x=x)
# export to Excel
table2excel(x=x, file = "table_aov.xlsx",digits=4,digitspvals=1,
            sheetName = "Anova_table", add.rownames = TRUE)
# export to Latex
table2tex(x=x)
# export to HTML
table2html(x=x)



#Error in value[[3L]](cond) : pptx device only supports one page
#test for example
library(export)
x = c(1,3,5,6,7,10,23,34)
y=c(3,9,14,17,22,30,63,99)
opar1<-par(no.readonly=TRUE)
par(mfrow=c(2,2))   
plot(x,y)
boxplot(x,y)
dotchart(x,y)
hist(x)
par(opar1) 

# export to Powerpoint      
graph2ppt()      
graph2ppt(file="four_example_plot.pptx", aspectr=1.7) #pptx device only supports one page



# End

