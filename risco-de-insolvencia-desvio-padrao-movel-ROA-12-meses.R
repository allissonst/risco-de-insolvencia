options(scipen = 9999)
library(readxl)
library(Hmisc)
library(dplyr)
library(ggplot2)
library(ggfortify)
library(outliers)
library(tidyverse)
library(plm)
library(psych)
library(caTools)
library(class)
library(caret)
library(ROCR)
library(klaR)
library(gains)
library(pROC)
library(e1071)
library(randomForest)
library(gbm)
library(ipred)
library(xlsx)
library(tm)
library(SnowballC)
library(wordcloud)
library(wesanderson)

#-----------------------------------------------
#incluindo a base de dados
#-----------------------------------------------

dados <- read_excel("C:/Users/Admin/Desktop/dados.xlsx")


dadosROA <- read_excel("C:/Users/Admin/Desktop/dados_ROA.xlsx")

#-----------------------------------------------
#tornando os dados não numéricos em numéricos
#-----------------------------------------------

dados[,c(3:16)]<- sapply(dados[, c(3:16)], as.numeric)

#-----------------------------------------------
#transformando dados dos denominadores que são 0 em NA, para evitar problemas de (inf;-inf)
#-----------------------------------------------

dados$ATIVTOT[dados$ATIVTOT==0]=NA
dados$PASTOT[dados$PASTOT==0]=NA
dados$PESSOAL[dados$PESSOAL==0]=NA
dados$PATLIQ[dados$PATLIQ==0]=NA
dados$REC_LIQ[dados$REC_LIQ==0]=NA
dados$PASSC[dados$PASSC==0]=NA
dadosROA$ATIVTOT[dadosROA$ATIVTOT==0]=NA

#-----------------------------------------------
#calculando variáveis e adicionando às matrizes
#-----------------------------------------------

#-----------------------------------------------
# Incluindo o Z-Score
#-----------------------------------------------

Xum=double(); Xdois=double();Xtres=double();Xquatro=double();ZSCORE=double()
Xum=(dados$ATIVC-dados$PASSC)/dados$ATIVTOT
Xdois=dados$LUCACUM/dados$ATIVTOT
Xtres=dados$EBIT/dados$ATIVTOT
Xquatro= dados$PATLIQ/dados$PASTOT
dados["ZSCORE"]=3.25+6.56*Xum+3.26*Xdois+6.72*Xtres+1.05*Xquatro

#-----------------------------------------------
# Incluindo o VAIC
#-----------------------------------------------

VA=dados$EBIT+dados$PESSOAL+dados$DEP_AMOR	
HCE=VA/dados$PESSOAL
SCE=(VA-dados$PESSOAL)/VA
CEE=VA/dados$PATLIQ
dados["VAIC"]=HCE+SCE+CEE

#-----------------------------------------------
# Incluindo as demais variáveis
#-----------------------------------------------

dados["TAM"]=log(dados$ATIVTOT)
dados["ROA"]=dados$LUC_LIQ/dados$ATIVTOT
dados["ALAV"]=(dados$PASSONER_CUR+dados$PASSONER_LON)/dados$ATIVTOT
dados["LIQC"]=dados$ATIVC/dados$PASSC
dados["GIRO"]=dados$REC_LIQ/dados$ATIVTOT
dados["CGAT"]=(dados$ATIVC-dados$PASSC)/dados$ATIVTOT
dados["FCDT"]=dados$FLU_CX/dados$PASTOT
dados["FCV"]=dados$FLU_CX/dados$REC_LIQ
dadosROA["ROA"]=dadosROA$LUCLIQ/dadosROA$ATIVTOT

#-----------------------------------------------  
#selecionando as variáveis de interesse em uma nova tabela
#-----------------------------------------------

mydados=dplyr::select(dados, c(1:2,29:37,17:28))
mydadosROA=dplyr::select(dadosROA, c(1:2,5))

#-----------------------------------------------
#transformando as datas em Date
#-----------------------------------------------

mydados$TRIM=as.Date(as.POSIXct(mydados$TRIM))
mydadosROA$TRIM=as.Date(as.POSIXct(mydadosROA$TRIM))

#-----------------------------------------------
#ordenando os dados a partir de PREFIXO & TRIMESTRE
#-----------------------------------------------

mydados=mydados[order(mydados$PREFIXO, mydados$TRIM, decreasing=c(F, F)),]
mydadosROA=mydadosROA[order(mydadosROA$PREFIXO, mydadosROA$TRIM, decreasing=c(F, F)),]

#-----------------------------------------------
#trabalhando com o desvio padrão móvel do ROA (12 trimestres) como medida de risco.
#-----------------------------------------------

dadosROA$ROA[is.na(dadosROA$ROA)==T]=0

dadosROA$DROA[1:nrow(dadosROA)]=c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA)
i=1
while (i<(nrow(dadosROA)-11)){
  if(dadosROA$PREFIXO[i]==dadosROA$PREFIXO[(i+11)]){
    dadosROA$DROA[(i+11)]=sd(dadosROA$ROA[i:(i+11)])}
  else{
    dadosROA$DROA[(i+11)]=NA
  }
  i=i+1
  
}

while (i<=nrow(dadosROA)){
  if(dadosROA$PREFIXO[i]==dadosROA$PREFIXO[i-11]){
    dadosROA$DROA[i]=sd(dadosROA$ROA[i-11:i])}
  else{
    dadosROA$DROA[i]=NA
  }
  i=i+1
}

#-----------------------------------------------
#passando o DP do ROA para PREF e TRIM de destino
#-----------------------------------------------

droa=dadosROA[,c(1:2,6)]
mydata=left_join(mydados, droa, by=c("PREFIXO","TRIM"))

#-----------------------------------------------
#Omitindo observações com dados faltantes
#-----------------------------------------------

mydata=na.omit(mydata)

#-----------------------------------------------
#Winsorização
#-----------------------------------------------

par(mfrow=c(2,5))
hist(mydata$VAIC,xlab="VAIC",main="");hist(mydata$TAM,xlab="TAM",main="");hist(mydata$ROA,xlab="ROA",main="");hist(mydata$ALAV,xlab="ALAV",main="");hist(mydata$LIQC,xlab="LIQC",main="");hist(mydata$GIRO,xlab="GIRO",main="");hist(mydata$CGAT,xlab="CGAT",main="");hist(mydata$FCDT,xlab="FCDT",main="");hist(mydata$FCV,xlab="FCV",main="")

mydata5 = mydata%>% mutate(VAICw=winsor(mydata$VAIC, trim=0.05)) %>% mutate(TAMw=winsor(mydata$TAM, trim=0.05)) %>% mutate(ROAw=winsor(mydata$ROA, trim=0.05))%>% mutate(ALAVw=winsor(mydata$ALAV, trim=0.05))%>%mutate(LIQCw=winsor(mydata$LIQC, trim=0.05))%>%mutate(GIROw=winsor(mydata$GIRO, trim=0.05))%>%mutate(CGATw=winsor(mydata$CGAT, trim=0.05))%>%mutate(FCDTw=winsor(mydata$FCDT, trim=0.05))%>%mutate(FCVw=winsor(mydata$FCV, trim=0.05))%>%mutate(ZSCOREw=winsor(mydata$ZSCORE, trim=0.05))%>%mutate(DROAw=winsor(mydata$DROA, trim=0.05))

par(mfrow=c(2,5))
hist(mydata5$VAICw, xlab="VAIC",main="");hist(mydata5$TAMw, xlab="TAM",main="");hist(mydata5$ROAw, xlab="ROA",main="");hist(mydata5$ALAVw, xlab="ALAV",main="");hist(mydata5$LIQCw,xlab="LIQC",main="");hist(mydata5$GIROw,xlab="GIRO",main="");hist(mydata5$CGATw,xlab="CGAT",main="");hist(mydata5$FCDTw,xlab="FCDT",main="");hist(mydata5$FCVw,xlab="FCV",main="")

mydata11 = mydata%>% mutate(VAICw=winsor(mydata$VAIC, trim=0.11)) %>% mutate(TAMw=winsor(mydata$TAM, trim=0.11)) %>% mutate(ROAw=winsor(mydata$ROA, trim=0.11))%>% mutate(ALAVw=winsor(mydata$ALAV, trim=0.11))%>%mutate(LIQCw=winsor(mydata$LIQC, trim=0.11))%>%mutate(GIROw=winsor(mydata$GIRO, trim=0.11))%>%mutate(CGATw=winsor(mydata$CGAT, trim=0.11))%>%mutate(FCDTw=winsor(mydata$FCDT, trim=0.11))%>%mutate(FCVw=winsor(mydata$FCV, trim=0.11))%>%mutate(ZSCOREw=winsor(mydata$ZSCORE, trim=0.11))%>%mutate(DROAw=winsor(mydata$DROA, trim=0.11))

hist(mydata11$VAICw,xlab="VAIC",main="");hist(mydata11$TAMw,xlab="TAM",main="");hist(mydata11$ROAw,xlab="ROA",main="");hist(mydata11$ALAVw,xlab="ALAV",main="");hist(mydata11$LIQCw,xlab="LIQC",main="");hist(mydata11$GIROw,xlab="GIRO",main="");hist(mydata11$CGATw,xlab="CGAT",main="");hist(mydata11$FCDTw,xlab="FCDT",main="");hist(mydata11$FCVw,xlab="FCV",main="")

#----------------------------------------------------------------------------------------------
# Winsorização escolhida de 5%, por ser o mais tolerado na literatura
#----------------------------------------------------------------------------------------------

mydata=mydata5

#----------------------------------------------------------------------------------------------
# definindo (1) maior e (0) menor risco de insolvência com base no Z-Score
# O ponto de corte é dado pelo próprio modelo de Altman (2005)
#----------------------------------------------------------------------------------------------

mydata$dZSCORE[mydata$ZSCOREw>=4.15]=0 #menor risco de insolvência
mydata$dZSCORE[mydata$ZSCOREw<4.15]=1 #maior risco de insolvência

#-----------------------------------------------
#selecionando as variáveis winsorizadas e dummy ZSCORE
#-----------------------------------------------

mydatafin=dplyr::select(mydata, c(1:2, 25:36))

#-----------------------------------------------
#Estatística descritiva
#----------------------------------------------
unique(mydatafin$PREFIXO) #quantidade de empresas na amostra

summary(mydatafin)
sd(mydatafin$VAICw);sd(mydatafin$TAMw);sd(mydatafin$ROAw);sd(mydatafin$ALAVw);sd(mydatafin$LIQCw);sd(mydatafin$GIROw);sd(mydatafin$CGATw);sd(mydatafin$FCDTw);sd(mydatafin$FCVw)

#-----------------------------------------------
#Correlação entre as variáveis
#-----------------------------------------------
cormydatafin=mydatafin[,3:11]
p=rcorr(as.matrix(cormydatafin));p

#-----------------------------------------------
#Modelo Kmeans
#-----------------------------------------------
Kms=scale(mydatafin$DROAw)
set.seed(1)
KMM=kmeans(Kms,2, iter.max = 20)
SampleKMM=KMM$cluster
KMM$centers
if(KMM$centers[1,]>KMM$centers[2,]){
  SampleKMM[SampleKMM==2]=0;
} else{
  SampleKMM[SampleKMM==1]=0;SampleKMM[SampleKMM==2]=1;
}

matchsaudavel=0; matchinsolv=0
i=1;

for (i in i:nrow(mydatafin)){
  
  if(mydatafin$dZSCORE[i]==0 && SampleKMM[i]==0){
    matchsaudavel=matchsaudavel+1
  }
  if(mydatafin$dZSCORE[i]==1 && SampleKMM[i]==1){
    matchinsolv=matchinsolv+1}
  
}

mydatafin$ROAKM=as.numeric(SampleKMM)

#-----------------------------------------------
#EXIBINDO RESULTADO DESCRITIVO DO KMEANS
#-----------------------------------------------

df.kmeans=data.frame(Risco = c("1 - maior risco", "0 - menor risco"),
                     Zscore= c(sum(as.numeric(mydatafin$dZSCORE)),nrow(mydatafin)-sum(as.numeric(mydatafin$dZSCORE))),
                     DPROA=c(sum(as.numeric(SampleKMM)),nrow(mydatafin)-sum(as.numeric(SampleKMM))),
                     
                     Comum=c(matchinsolv,matchsaudavel),
                     Similar=c(matchinsolv/sum(as.numeric(mydatafin$dZSCORE)),matchsaudavel/((nrow(mydatafin))-sum(as.numeric(mydatafin$dZSCORE)))))
df.kmeans
  
set.seed(1)
#-----------------------------------------------
#ZSCORE EMPRESAS COM MAIOR RISCO DE INSOLVÊNCIA
#-----------------------------------------------

ZSCORE=mydatafin%>%filter(mydatafin$dZSCORE==1)

#-----------------------------------------------
#Download de planilha excel, quando necessário
#write.xlsx(as.data.frame(ZSCORE), file="zscore.xlsx")
#-----------------------------------------------

#-----------------------------------------------
#GERAÇÃO DA NUVEM DE PALAVRAS - Z-SCORE
#-----------------------------------------------

aux=ZSCORE$PREFIXO
auxCorpus <- Corpus(VectorSource(aux))
wordcloud(auxCorpus,max.words=110,scale = c(4, 0.3),random.order=F, color=wes_palette("Darjeeling1"), rot.per = 0.7)

#-----------------------------------------------
#KMEANS EMPRESAS COM MAIOR RISCO DE INSOLVÊNCIA
#-----------------------------------------------

KMEANS=mydatafin%>%filter(mydatafin$ROAKM==1)

#-----------------------------------------------
#GERAÇÃO DA NUVEM DE PALAVRAS - KMEANS
#-----------------------------------------------

aux=KMEANS$PREFIXO
auxCorpus <- Corpus(VectorSource(aux))
wordcloud(auxCorpus,max.words=110,scale = c(4, 0.3),random.order=F, color=wes_palette("Darjeeling1"), rot.per = 0.7)

#-----------------------------------------------
#Download de planilha excel, quando necessário
#write.xlsx(as.data.frame(KMEANS), file="KMEANS.xlsx")
#-----------------------------------------------

#-----------------------------------------------
#   SEPARAÇÃO - TREINAMENTO/TESTE DPROA
#-----------------------------------------------
set.seed(1)
divisao=sample.split(mydatafin$ROAKM,SplitRatio = 0.75)
treinamento=subset(mydatafin,divisao==T)
teste=subset(mydatafin,divisao==F)
treinamento$ROAKM=as.factor(treinamento$ROAKM);teste$ROAKM=as.factor(teste$ROAKM)
mettreinamento=treinamento[,c(3:11,15)];metteste=teste[,c(3:11,15)];

#-----------------------------------------------
#APLICAÇÃO KNN (K-Nearest Neighbor, ou K Vizinhos Mais Próximos)
#-----------------------------------------------

ctrl <- trainControl(method="repeatedcv",
                     repeats = 3) 

knnFit <- train(ROAKM ~ ., 
                data = mettreinamento, 
                method = "knn", trControl = ctrl, 
                preProcess = c("scale"), 
                tuneLength = 20)


knnclass=predict(knnFit, newdata=metteste)


confusionMatrix(knnclass,metteste$ROAKM, positive = "1")
knnclassprob=predict(knnFit, newdata=metteste, type="prob")
rocobject=roc(metteste$ROAKM, knnclassprob[,2])
plot.roc(rocobject)

auc(rocobject)

#-----------------------------------------------
#     OUTRAS MÉTRICAS DE AVALIAÇÃO - KNN
#-----------------------------------------------

predicttest=as.numeric(knnclass);
original=as.numeric(metteste$ROAKM);

MAE(predicttest,original)
RMSE(predicttest,original)

#-----------------------------------------------
#                 APLICAÇÃO - Naive bayes
#-----------------------------------------------

set.seed(1)
ctrl=trainControl(method="cv",
                  number=10)
nbFit= train(
  ROAKM ~ .,
  method="nb",
  preProcess=c("scale"),
  trControl=ctrl,
  data=mettreinamento
)
nbFit2=naiveBayes(ROAKM ~ ., data=mettreinamento)

nbclass=predict(nbFit, newdata=metteste)
confusionMatrix(nbclass,metteste$ROAKM)
nbclassprob=predict(nbFit, newdata=metteste, type="prob")
rocobject=roc(metteste$ROAKM, nbclassprob[,2])
plot.roc(rocobject)

#-----------------------------------------------
#      OUTRAS MÉTRICAS DE AVALIAÇÃO - Naive Bayes
#-----------------------------------------------

auc(rocobject)

predicttest=as.numeric(nbclass);
original=as.numeric(metteste$ROAKM);

MAE(predicttest,original)
RMSE(predicttest,original)

#-----------------------------------------------
#                 APLICAÇÃO - Logit
#-----------------------------------------------
set.seed(1)
fitlogit=glm(ROAKM ~ ., data=mettreinamento, family="binomial")
Fitlogit= train(
  ROAKM ~ .,
  method="glm",
  preProcess=c("scale"),
  family="binomial",
  data=mettreinamento
)
logitclass=predict(Fitlogit, newdata=metteste)
confusionMatrix(logitclass,metteste$ROAKM,positive = "1")
logitclassprob=predict(Fitlogit, newdata=metteste, type="prob")
logitrocobject=roc(metteste$ROAKM, logitclassprob[,2])
plot.roc(logitrocobject)

#-----------------------------------------------
#      OUTRAS MÉTRICAS DE AVALIAÇÃO - Logit
#-----------------------------------------------

plot(varImp(Fitlogit))
auc(logitrocobject)

predicttest=as.numeric(logitclass);
original=as.numeric(metteste$ROAKM);

MAE(predicttest,original)
RMSE(predicttest,original)

#-----------------------------------------------
#                 APLICAÇÃO - Random Forest
#-----------------------------------------------
set.seed(1)
ctrl=trainControl(method="cv",
                  number=10)
FitRF= train(
  ROAKM ~ .,
  method="rf",
  preProcess=c("scale"),
  trControl=ctrl,
  data=mettreinamento,
  importance=T
)
RFclass=predict(FitRF, newdata=metteste)
confusionMatrix(RFclass,metteste$ROAKM,positive = "1")
RFclassprob=predict(FitRF, newdata=metteste, type="prob")
RFrocobject=roc(metteste$ROAKM, RFclassprob[,2])
plot.roc(RFrocobject)
plot(varImp(FitRF))

#-----------------------------------------------
#      OUTRAS MÉTRICAS DE AVALIAÇÃO - Random Forest
#-----------------------------------------------

auc(RFrocobject)

predicttest=as.numeric(RFclass);
original=as.numeric(metteste$ROAKM);

MAE(predicttest,original)
RMSE(predicttest,original)

#-----------------------------------------------
#           APLICAÇÃO - Support Vector Machines
#-----------------------------------------------

#-----------------------------------------------
#           RADIAL BASIS
#-----------------------------------------------
set.seed(1)

svmfit=svm(ROAKM ~ ., data=mettreinamento,scale=T, kernel="radial")
SVMclass=predict(svmfit, newdata=metteste)
confusionMatrix(SVMclass,metteste$ROAKM, positive="1")
SVMclassprobradial=predict(svmfit, newdata=metteste, type="prob")
SVMrocobject=roc(metteste$ROAKM,as.numeric(SVMclassprobradial))
plot.roc(SVMrocobject)

#-----------------------------------------------
#      OUTRAS MÉTRICAS DE AVALIAÇÃO - SVM Radial
#-----------------------------------------------

auc(SVMrocobject)

predicttest=as.numeric(SVMclass);
original=as.numeric(metteste$ROAKM);

MAE(predicttest,original)
RMSE(predicttest,original)


#-----------------------------------------------
#           POLYNOMIAL
#-----------------------------------------------

set.seed(1)

svmfit=svm(ROAKM ~ ., data=mettreinamento,scale=T, type= "C-classification", kernel="polynomial")
SVMclass=predict(svmfit, newdata=metteste)
confusionMatrix(SVMclass,metteste$ROAKM)
SVMclassprobpolinomial=predict(svmfit, newdata=metteste, type="prob")
SVMrocobject=roc(metteste$ROAKM,as.numeric(SVMclassprobpolinomial))
plot.roc(SVMrocobject)

#-----------------------------------------------
#      OUTRAS MÉTRICAS DE AVALIAÇÃO - SVM Polinomial
#-----------------------------------------------

auc(SVMrocobject)

predicttest=as.numeric(SVMclass);
original=as.numeric(metteste$ROAKM);

MAE(predicttest,original)
RMSE(predicttest,original)

#-----------------------------------------------
#           LINEAR
#-----------------------------------------------
set.seed(1)
svmfit=svm(ROAKM ~ ., data=mettreinamento, scale=T, type= "C-classification", kernel="linear")
SVMclass=predict(svmfit, newdata=metteste)
confusionMatrix(SVMclass,metteste$ROAKM)
SVMclassproblinear=predict(svmfit, newdata=metteste, type="prob")
SVMrocobject=roc(metteste$ROAKM,as.numeric(SVMclassproblinear))
plot.roc(SVMrocobject)

#-----------------------------------------------
#      OUTRAS MÉTRICAS DE AVALIAÇÃO - SVM Linear
#-----------------------------------------------

auc(SVMrocobject)

predicttest=as.numeric(SVMclass);
original=as.numeric(metteste$ROAKM);

MAE(predicttest,original)
RMSE(predicttest,original)

#-----------------------------------------------
#           APLICAÇÃO - Bagged Model
#-----------------------------------------------

set.seed(1)

baggedfit <- bagging(formula = ROAKM ~ ., 
                     data = mettreinamento,
                     coob = TRUE,
                     scale=T)

baggedclass <- predict(object = baggedfit,    
                       newdata = metteste,  
                       type = "class")

confusionMatrix(data = baggedclass,       
                reference = metteste$ROAKM, positive = "1")

baggedclassprob <- predict(object = baggedfit,    
                           newdata = metteste,  
                           type = "prob")

baggedrocobject=roc(metteste$ROAKM, baggedclassprob[,2])
plot.roc(baggedrocobject)

#-----------------------------------------------
#      OUTRAS MÉTRICAS DE AVALIAÇÃO - Bagged Model
#-----------------------------------------------

auc(baggedrocobject)

predicttest=as.numeric(baggedclass);
original=as.numeric(metteste$ROAKM);

MAE(predicttest,original)
RMSE(predicttest,original)

set.seed(1)
#-----------------------------------------------
#           APLICAÇÃO - boosted trees
#-----------------------------------------------

tc = trainControl(method = "cv", 
                  number=10)
fitboosted = train(ROAKM ~., 
                   data=mettreinamento, 
                   method="gbm", 
                   trControl=tc,
                   preProcess = c("scale"))

boostedclass=predict(fitboosted, newdata=metteste)
confusionMatrix(boostedclass,metteste$ROAKM, positive = "1")
boostedclassprob=predict(fitboosted, newdata=metteste, type="prob")
boostedrocobject=roc(metteste$ROAKM, boostedclassprob[,2])
plot.roc(boostedrocobject)

#-----------------------------------------------
#      OUTRAS MÉTRICAS DE AVALIAÇÃO - boosted trees
#-----------------------------------------------

auc(boostedrocobject)

predicttest=as.numeric(boostedclass);
original=as.numeric(metteste$ROAKM);

MAE(predicttest,original)
RMSE(predicttest,original)

#-----------------------------------------------
#         COMPARAÇÃO DE MODELOS COM O ROC
#-----------------------------------------------

# List of predictions
preds_list <- list(knnclassprob[,2], nbclassprob[,2], RFclassprob[,2], logitclassprob[,2], SVMclassprobradial, SVMclassprobpolinomial, SVMclassproblinear, boostedclassprob[,2], baggedclassprob[,2])

# List of actual values (same for all)
m <- length(preds_list)
actuals_list <- rep(list(as.numeric(metteste$ROAKM)), m)
# Plot the ROC curves
pred <- prediction(predictions=preds_list, labels=actuals_list)
rocs <- performance(pred, "tpr", "fpr")
plot(rocs, col = as.list(1:m), main = "Test Set ROC Curves")
legend(x = "bottomright", 
       legend = c("KNN", "Naive Bayes","Random Forest","Logit","SVM Radial","SVM Polinomial","SVM linear", "Boosting","Bagging"),
       fill = 1:m)
