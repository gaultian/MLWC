# Project - Calculating Normal Range Intervals
# Version 0.0
# Authour - Collin J. Arens - edited Jennifer Arens 14-06-2021
#NON-parametric step does percentile calculation only - no resampling - Present results using SR=100 (basically no outlier removal) and SR 3.5
# Last Modified - 22/Feb/2020
rm(list=ls()) #reset session
# Libraries----
library(openxlsx) #Open xlsx file
library(tidyverse) #data management
library(olsrr) #diagnostic figures
library(car) # needed for powerTransform function

# Import data----
NR.Data<-read.xlsx("data-raw/NR_Data5_add.xlsx",sheet= 4 ,startRow=1,colNames=TRUE,na.strings="") #import xlsx file


# Define Variables---
alpha<-0.05 #define alpha
tails<-2 #define number of tails
m<-1 #define m - needs to be one at this point
SW<-0.05 #define alpha for Shapiro Wilk
SR<-3.5 #define outlier studentized residual threshold

# Outlier Screening ----
for(i in 1:length(NR.Data[1,])){
  tmp<-rstudent(lm(NR.Data[,i]~1))>=SR|rstudent(lm(NR.Data[,i]~1))<=-SR #create a temp file identifying outliers
  NR.Data[,i][tmp]<-(NR.Data[,i][tmp]=NA) #remove outliers from each vector
}
# Parametric - Untransformed----
NR.Results = NULL
for(i in 1:length(NR.Data[1,])){
  NR.Results$Parameter[i]=colnames(NR.Data[i]) #parameter name
  NR.Results$n[i]=sum(!is.na(NR.Data[,i])) #parameter name sample size w/o outliers
  NR.Results$SW[i]=if((shapiro.test(NR.Data[,i])[["p.value"]])>SW){TRUE}else{FALSE} #normality test
  NR.Results$Upper[i]=mean(NR.Data[,i],na.rm=TRUE)+qt(1-alpha/tails,sum(!is.na(NR.Data[,i]))-1)*sd(NR.Data[,i],na.rm=TRUE)*sqrt(1/sum(!is.na(NR.Data[,i]))+1/m) #upper bound
  NR.Results$Lower[i]=mean(NR.Data[,i],na.rm=TRUE)-qt(1-alpha/tails,sum(!is.na(NR.Data[,i]))-1)*sd(NR.Data[,i],na.rm=TRUE)*sqrt(1/sum(!is.na(NR.Data[,i]))+1/m) #lower bound
}
# Parametric - Box Cox Transformation--------------------------
for(i in 1:length(NR.Data[1,])){
  NR.Results$lambda[i]=powerTransform(NR.Data[,i])[["lambda"]][["NR.Data[, i]"]] #calculate lambda for each vector
}
bc<-function (obs, lambda) {
  (obs^lambda-1)/lambda } #define Box Cox function
NR.BC<-mapply(bc,NR.Data,NR.Results$lambda) #apply Box Cox function to new data.frame
for(i in 1:length(NR.BC[1,])){
  NR.Results$SW_BC[i]=if((shapiro.test(NR.BC[,i])[["p.value"]])>SW){TRUE}else{FALSE} #normality test
  NR.Results$Upper_BC[i]=((mean(NR.BC[,i],na.rm=TRUE)+qt(1-alpha/tails,sum(!is.na(NR.BC[,i]))-1)*sd(NR.BC[,i],na.rm=TRUE)*sqrt(1/sum(!is.na(NR.BC[,i]))+1/m))*NR.Results$lambda[i]+1)^(1/NR.Results$lambda[i]) #upper bound
  NR.Results$Lower_BC[i]=((mean(NR.BC[,i],na.rm=TRUE)-qt(1-alpha/tails,sum(!is.na(NR.BC[,i]))-1)*sd(NR.BC[,i],na.rm=TRUE)*sqrt(1/sum(!is.na(NR.BC[,i]))+1/m))*NR.Results$lambda[i]+1)^(1/NR.Results$lambda[i]) #lower bound
}
# Non-Parametric----
NP.Iterations=NULL
for(i in 1:length(NR.Data[1,])){
  
  NR.Results$Upper_NP[i]=quantile(NR.Data[,i],1-alpha/tails, na.rm=TRUE)
  
  NR.Results$Lower_NP[i]=quantile(NR.Data[,i],alpha/tails, na.rm=TRUE)
  
}









# Compile and Export----
NR.Results=data.frame(NR.Results)
write.xlsx(NR.Results,"data/NR_Results_GW_OUTLIERSREMOVED_EHZ6_2023_06_15.xlsx") #export table
                            


#Split Data based on TSS and plot below using different symbols 


# Create figures for each parameter - **Add outliers**
for(i in 1:length(NR.Data[1,])){
Figure<-ggplot()+
 geom_rect(aes(ymin=(NR.Results[subset(NR.Results==colnames(NR.Data[i])),5]), ymax=(NR.Results[subset(NR.Results==colnames(NR.Data[i])),4]), xmin=-Inf, xmax=Inf),fill='grey90')+
 geom_boxplot(aes(y=NR.Data[,i]))+
 theme_bw()+
 labs(x=NULL,y=(colnames(NR.Data[i])))+
 scale_x_discrete(breaks=NULL)
ggsave(filename = paste("figures/NR Plots_EHZ6_2023_", colnames(NR.Data[i]), ".png", sep = ""), plot = Figure, width = 2, height = 3)
#Export Figure
}
warnings()
