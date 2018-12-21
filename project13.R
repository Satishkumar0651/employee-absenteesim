rm(list = ls())

#libraries
x <- c("ggplot2", "corrgram", "DMwR", "caret", "randomForest", "unbalanced", "C50", "dummies", "e1071", "Information",
       "MASS", "rpart", "gbm", "ROSE", 'sampling', 'DataCombine', 'inTrees','readxl')

#load packages(x)
lapply(x, require, character.only = TRUE)
rm(x)
setwd("E:/1ST SEM/eng/edwisor_assignments/7.project1")

#data
d1 <- read_excel("Absenteeism_at_work_Project.xls")

str(d1)

#missing value analysis
missing_value = data.frame(apply(d1,2,function(x){sum(is.na(x))}))
missing_value$Columns = row.names(missing_value)
names(missing_value)[2] =  "Variables"
row.names(missing_value)=NULL
names(missing_value)[1] =  "Missing_percentage"
missing_value$Missing_percentage = (missing_value$Missing_percentage/nrow(d1)) * 100
missing_value = missing_value[order(-missing_value$Missing_percentage),]
missing_value = missing_value[,c(2,1)]

names(d1)[1]="id"
names(d1)[2]="reason.of.absence"
names(d1)[3]="month.of.absence"
names(d1)[4]="day.of.the.week"
names(d1)[5]="seasons"
names(d1)[6]="transportation.expenses"
names(d1)[7]="distance.residence.work"
names(d1)[8] = "service.time"
names(d1)[9]="age"
names(d1)[10]="work.load.average"
names(d1)[11]="hit.target"
names(d1)[12]="disciplinary.failure"
names(d1)[13]="education"
names(d1)[14]="son"
names(d1)[15]="social.drinker"
names(d1)[16]="social.smoker"
names(d1)[17]="pet"
names(d1)[18]="weight"
names(d1)[19]="height"
names(d1)[20]="body.mass.index"
names(d1)[21]="absenteeism.hours"



#imputing missing values
d1$reason.of.absence[is.na(d1$reason.of.absence)] = median(d1$reason.of.absence, na.rm = T)
d1$month.of.absence[is.na(d1$month.of.absence)] = median(d1$month.of.absence, na.rm = T)
d1$disciplinary.failure[is.na(d1$disciplinary.failure)] = median(d1$disciplinary.failure, na.rm = T)
d1$education[is.na(d1$education)] = median(d1$education, na.rm = T)
d1$social.drinker[is.na(d1$social.drinker)] = median(d1$social.drinker, na.rm = T)
d1$social.smoker[is.na(d1$social.smoker)] = median(d1$social.smoker, na.rm = T)
d1$transportation.expenses[is.na(d1$transportation.expenses)] = median.default(d1$transportation.expenses, na.rm = T)
d1$distance.residence.work[is.na(d1$distance.residence.work)] = median(d1$distance.residence.work, na.rm = T)
d1$service.time[is.na(d1$service.time)] = median(d1$service.time, na.rm = T)
d1$age[is.na(d1$age)] = median(d1$age, na.rm = T)
d1$work.load.average[is.na(d1$work.load.average)] = median(d1$work.load.average, na.rm = T)
d1$hit.target[is.na(d1$hit.target)] = median(d1$hit.target, na.rm = T)
d1$son[is.na(d1$son)] = median(d1$son, na.rm = T)
d1$pet[is.na(d1$pet)] = median(d1$pet, na.rm = T)
d1$weight[is.na(d1$weight)] = median(d1$weight, na.rm = T)
d1$height[is.na(d1$height)] = median(d1$height, na.rm = T)
d1$body.mass.index[is.na(d1$body.mass.index)] = median(d1$body.mass.index, na.rm = T)
d1$absenteeism.hours[is.na(d1$absenteeism.hours)] = median(d1$absenteeism.hours, na.rm = T)


#converting variables to their types
d1$reason.of.absence=as.factor(d1$reason.of.absence)
d1$month.of.absence = as.factor(d1$month.of.absence)
d1$day.of.the.week = as.factor(d1$day.of.the.week)
d1$seasons = as.factor(d1$seasons)
d1$disciplinary.failure = as.factor(d1$disciplinary.failure)
d1$education = as.factor(d1$education)
d1$social.drinker = as.factor(d1$social.drinker)
d1$social.smoker= as.factor(d1$social.smoker)


df= d1
#d1 = df
numeric_index = sapply(d1,is.numeric) #selecting only numeric
numerical = d1[,numeric_index]
Numerical = colnames(numerical)


for (i in 1:length(Numerical))
{
  assign(paste0("gn",i), ggplot(aes_string(y = (Numerical[i]), x = "absenteeism.hours"), data = subset(d1))+ 
           stat_boxplot(geom = "errorbar", width = 0.5) +
           geom_boxplot(outlier.colour="red", fill = "blue" ,outlier.shape=18,
                        outlier.size=1, notch=FALSE) +
           theme(legend.position="bottom")+
           labs(y=Numerical[i],x="absenteeism.hours")+
           ggtitle(paste("Box plot for",Numerical[i])))
}

gridExtra::grid.arrange(gn1,gn2,gn3,ncol=3)
gridExtra::grid.arrange(gn4,gn5,gn6,ncol=3)
gridExtra::grid.arrange(gn7,gn8,gn9,ncol=3)
gridExtra::grid.arrange(gn10,gn11,ncol=2)


#outlier impute using NA technique
Out = d1$transportation.expenses[d1$transportation.expenses %in% boxplot.stats(d1$transportation.expenses)$out]
d1$transportation.expenses[(d1$transportation.expenses %in% Out)] = NA
d1$transportation.expenses[is.na(d1$transportation.expenses)] = median.default(d1$transportation.expenses, na.rm = T)

Out = d1$distance.residence.work[d1$distance.residence.work %in% boxplot.stats(d1$distance.residence.work)$out]
d1$distance.residence.work[(d1$distance.residence.work %in% Out)] = NA
d1$distance.residence.work[is.na(d1$distance.residence.work)] = median(d1$distance.residence.work, na.rm = T)

Out = d1$age[d1$age %in% boxplot.stats(d1$age)$out]
d1$age[(d1$age %in% Out)] = NA
d1$age[is.na(d1$age)] = median(d1$age, na.rm = T)

Out = d1$ work.load.average [d1$ work.load.average  %in% boxplot.stats(d1$ work.load.average )$out]
d1$ work.load.average [(d1$ work.load.average  %in% Out)] = NA
d1$ work.load.average [is.na(d1$ work.load.average )] = median(d1$ work.load.average , na.rm = T)

Out = d1$hit.target[d1$hit.target %in% boxplot.stats(d1$hit.target)$out]
d1$hit.target[(d1$hit.target %in% Out)] = NA
d1$hit.target[is.na(d1$hit.target)] = median(d1$hit.target, na.rm = T)

Out = d1$pet[d1$pet %in% boxplot.stats(d1$pet)$out]
d1$pet[(d1$pet %in% Out)] = NA
d1$pet[is.na(d1$pet)] = median(d1$pet, na.rm = T)

Out = d1$height[d1$height %in% boxplot.stats(d1$height)$out]
d1$height[(d1$height %in% Out)] = NA
d1$height[is.na(d1$height)] = median(d1$height, na.rm = T)

Out = d1$absenteeism.hours[d1$absenteeism.hours %in% boxplot.stats(d1$absenteeism.hours)$out]
d1$absenteeism.hours[(d1$absenteeism.hours %in% Out)] = NA
d1$absenteeism.hours[is.na(d1$absenteeism.hours)] = median(d1$absenteeism.hours, na.rm = T)

rm(Out)
#rm(df)

#correlation plot
corrgram(d1, order = F,
         upper.panel=panel.pie, text.panel=panel.txt, main = "Correlation Plot")



str(d1)

d2 = d1
d1 = d2




df = subset(d2, select = c(month.of.absence,absenteeism.hours,service.time,work.load.average))


library(randomForest)

random_model = randomForest(absenteeism.hours~.,data = d1 , importance= TRUE,ntree = 500)
print(random_model)
attributes(random_model)
varUsed(random_model) # to find which variables used in random forest
varImpPlot(random_model,sort = TRUE,n.var = 10,main=" Importance graph")
varImp(random_model)

# anova for categorical data
str(d2)
library(ANOVA.TFNs)
library(ANOVAreplication)
result = aov(formula=absenteeism.hours~reason.of.absence+month.of.absence+day.of.the.week+seasons+disciplinary.failure
             +education+social.smoker+social.drinker, data = d1)
summary(result)


#library(rpart.plot)
#tree <- rpart(absenteeism.hours ~ . , method='class', data = d1)
#printcp(tree)
#plot(tree, uniform=TRUE, main="Main Title")
#text(tree, use.n=TRUE, all=TRUE)
#prp(tree)


d1 = subset(d1, select = -c(id,education,day.of.the.week, pet,hit.target, seasons,social.smoker,
                            social.drinker,weight))
str(d1)

#qqnorm(d1$transportation.expenses)
#hist(d1$transportation.expenses)
#qqnorm(d1$distance.residence.work)
#hist(d1$distance.residence.work)
#qqnorm(d1$service.time)
#hist(d1$service.time)
#qqnorm(d1$work.load.average)
#hist(d1$work.load.average)
#qqnorm(d1$height)
#hist(d1$height)
#qqnorm(d1$age)
#hist(d1$age)
#qnorm(d1$body.mass.index)
#hist(d1$body.mass.index)
#qqnorm(d1$absenteeism.hours)
#hist(d1$absenteeism.hours)


qqnorm(d1$absenteeism.hours)
hist(d1$absenteeism.hours)



Numerical_Col = c("transportation.expenses","age","son","height","body.mass.index","distance.residence.work","work.load.average",
                  "distance.residence.work")

for(i in Numerical_Col){
  print(i)
  d1[,i] = (d1[,i] - min(d1[,i]))/
    (max(d1[,i] - min(d1[,i])))
}

rmExcept(c("d2","d1"))

sample = sample(1:nrow(d1), 0.8 * nrow(d1))
train = d1[sample,]
test = d1[-sample,]

library(caTools)
library(mltools)
RMSE <- function(Error)
{
  sqrt(mean(Error^2))
}

dtree = rpart(absenteeism.hours~.,data = train, method = "anova")
dtree.plt = rpart.plot(dtree,type = 3,digits = 3,fallen.leaves = TRUE)
prediction_dtree = predict(dtree,test[,-12])
actual = test[,12]
predicted_dtree = data.frame(prediction_dtree)
Error = actual - predicted_dtree
RMSE(Error$absenteeism.hours)


# Random forest
random_model = randomForest(absenteeism.hours~.,train,importance = TRUE,ntree = 100)
rand_pred = predict(random_model,test[-12])
actual = test[,12]
predicted_rand = data.frame(rand_pred)
Error = actual - predicted_rand
RMSE(Error$absenteeism.hours)


d1$reason.of.absence=as.numeric(d1$reason.of.absence)
d1$month.of.absence = as.numeric(d1$month.of.absence)
d1$disciplinary.failure= as.numeric(d1$disciplinary.failure)
str(d1)

#droplevels(d1$Reason.for.absence)
sample_lm = sample(1:nrow(d1),0.8*nrow(d1))
train_lm = d1[sample_lm,]
test_lm = d1[-sample_lm,]
linear_model = lm(absenteeism.hours~.,data = train_lm)
summary(linear_model)

vif(linear_model)
predictions_lm = predict(linear_model,test_lm[,1:11])
Predicted_LM = data.frame(predictions_lm)
Actual = test_lm[,12]
Error = Actual - Predicted_LM
RMSE(Error$absenteeism.hours)



write.csv(d2,"dk.csv",row.names=F)

#part 2

DFF = subset(d2, select = c(month.of.absence,service.time,absenteeism.hours, 
                                    work.load.average))

DFF["Loss"]=with(DFF,((DFF[,4]*DFF[,3])/DFF[,2]))

for(i in 1:12)
{
  LOSS=DFF[which(DFF["month.of.absence"]==i),]
  print(data.frame(sum(LOSS$Loss)))
  
}

