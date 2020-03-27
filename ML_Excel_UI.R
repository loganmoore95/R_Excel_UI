#Collect the arguments that the excel file has supplied to the command line
args <- commandArgs(trailingOnly = T)

#To avoid unnecessary output on the excel end, invisible() is used,

#Read in necessary packages. 
if (!require("readxl")) {
  install.packages('readxl',
                   repos="https://cran.rstudio.com/",
                   quiet = TRUE)
  require('readxl')
}

if (!require("AUC")) {
  install.packages('AUC',
                   repos="https://cran.rstudio.com/",
                   quiet = TRUE)
  require('AUC')
}

if (!require("e1071")) {
  install.packages('e1071',
                   repos="https://cran.rstudio.com/",
                   quiet = TRUE)
  require('e1071')
}

if (!require("glmnet")) {
  install.packages('glmnet',
                   repos="https://cran.rstudio.com/",
                   quiet = TRUE)
  require('glmnet')
}

if (!require("nnet")) {
  install.packages('nnet',
                   repos="https://cran.rstudio.com/",
                   quiet = TRUE)
  require('nnet')
}

if (!require("FNN")) {
  install.packages('FNN',
                   repos="https://cran.rstudio.com/",
                   quiet = TRUE)
  require('FNN')
}

if (!require("rpart")) {
  install.packages('rpart',
                   repos="https://cran.rstudio.com/",
                   quiet = TRUE)
  require('rpart')
}

if (!require("randomForest")) {
  install.packages('randomForest',
                   repos="https://cran.rstudio.com/",
                   quiet = TRUE)
  require('randomForest')
}

if (!require("ada")) {
  install.packages('ada',
                   repos="https://cran.rstudio.com/",
                   quiet = TRUE)
  require('ada')
}

if (!require("rotationForest")) {
  install.packages('rotationForest',
                   repos="https://cran.rstudio.com/",
                   quiet = TRUE)
  require('rotationForest')
}

#Load the arguments from the command line
filepath <- as.character(args[1])
datasheetname <- as.character(args[2])
newdatasheetname <- as.character(args[3])
yvalue <- as.character(args[4])
xvalues <- as.character(args[5])
NaiveBayes <- as.numeric(args[6])
SLogistic <- as.numeric(args[7])
RLogistic <- as.numeric(args[8])
NNet <- as.numeric(args[9])
KNN <- as.numeric(args[10])
DT <- as.numeric(args[11])
BT <- as.numeric(args[12])
RF <- as.numeric(args[13])
AB <- as.numeric(args[14])
RotF <- as.numeric(args[15])
AUCR <- as.numeric(args[16])
PR <- as.numeric(args[17])

#Read in the data
DATA <- read_excel(filepath, sheet=datasheetname)
DATA <- as.data.frame(DATA)

NEWDATA <- read_excel(filepath, sheet = newdatasheetname)
NEWDATA <- as.data.frame(NEWDATA)

#Remove all spaces after commas and separate the xvalues into rows of a data frame
xvalues <- gsub(", ",",",xvalues)
xvalues <- unlist(strsplit(xvalues,","))
  
#Remove any columns not used in the analysis
neededcols <- c(xvalues,yvalue)
DATA <- DATA[neededcols]
NEWDATA <- NEWDATA[neededcols]
  
#Data needs to be split into train, validation, and test sets.
#create randomized indicators for each row of data
allind <- sample(x=1:nrow(DATA), size = nrow(DATA))
  
#split each row randomly into three parts
trainind <- allind[1:round(length(allind)/3)]
valind <- allind[(round(length(allind)/3)+1):round(length(allind)*(2/3))]
testind <- allind[round(length(allind)*(2/3)+1):length(allind)]
  
#Now we can simply create three different data frames for train, val, and test.
TRAINDATA <- DATA[trainind,]
VALDATA <- DATA[valind,]
TESTDATA <- DATA[testind,]
  
#Because some algorithms will not reqiure tuning, we need a data fram of the combined train and validation set
TRAINBIGDATA <- rbind(TRAINDATA,VALDATA)
  
#Assign the Y values as factors
yTRAIN <- as.data.frame(TRAINDATA[yvalue])
yTRAIN <- as.factor(yTRAIN[,1])

yVAL <- as.data.frame(VALDATA[yvalue])
yVAL <- as.factor(yVAL[,1])

yTEST <- as.data.frame(TESTDATA[yvalue])
yTEST <- as.factor(yTEST[,1])

yTRAINBIG <- as.data.frame(TRAINBIGDATA[yvalue])
yTRAINBIG <- as.factor(yTRAINBIG[,1])
  
#Remove the Y value from the other tables, as it is no longer needed
TRAINDATA <- TRAINDATA[xvalues]
VALDATA <- VALDATA[xvalues]
TESTDATA <- TESTDATA[xvalues]
TRAINBIGDATA <- TRAINBIGDATA[xvalues]
NEWDATA <- NEWDATA[xvalues]

#Create a dataframe to hold the results of our models
RESULTS <- data.frame(Algorithm = as.character(c("Naive Bayes","Stepwise Logistic Regression","Lasso Logistic Regression",
                                                 "Neural Network", "K-Nearest Neighbors", "Decision Tree", "Bagged Trees",
                                                 "Random Forest", "Adaptive Boosting", "Rotation Forest")), 
                      AUC_ROC = as.numeric(NA))

#Create a dataframe to hold the predictions of our models
PREDICTIONS <- data.frame(a = as.numeric(rep(0,nrow(NEWDATA))), b = as.numeric(0), c = as.numeric(0), 
                           d = as.numeric(0), e = as.numeric(0), f = as.numeric(0),
                           g = as.numeric(0), h = as.numeric(0), i = as.numeric(0),
                           j = as.numeric(0))

#Create column names based on algorithm
for(i in 1:nrow(RESULTS)){
  colnames(PREDICTIONS)[i] <- as.character(RESULTS[i,1])
}

if(NaiveBayes == 1){
  print("Beginning Naive Bayes estimation.")
  
  #Fit the NaiveBayes model
  NB <-naiveBayes(x=TRAINBIGDATA, y=yTRAINBIG)
  
  #Use the model to predict on our test data
  #note that type can be changed to "class" to get labels
  predNB <- predict(NB, TESTDATA, type = "raw")[,2]
  
  #Store the accuracy of the model in the results dataframe
  
  RESULTS$AUC_ROC[1] <-  round(AUC::auc(roc(predNB,yTEST)),4)
  
  #Predict the new data with this model
  predNBnew <- predict(NB, NEWDATA, type = "raw")[,2]
  
  #Place the predictions in the predictions data frame
  PREDICTIONS$`Naive Bayes` <- round(predNBnew,5)
  
  #Remove unneeded items from memory
  rm(list = c('NB','predNB','predNBnew'))
}

if(SLogistic == 1){
  print("Beginning Bi-Directional Stepwise Logistic Regression estimation.")
  
  #Fit the Logistic Regression model
  SLR <- glm(yTRAINBIG ~ .,
            data=TRAINBIGDATA,
            family=binomial("logit"))
  
  #Run a stepwise bi-directional variable selection process
  LRstep <- step(SLR, direction = "both", trace = FALSE)
  
  #Use this model to make a prediction on our test data.
  predLRstep <- predict(LRstep,
                        newdata = TESTDATA,
                        type = "response")
  
  #And let's assess the performance of the model
  RESULTS$AUC_ROC[2] <- round(AUC::auc(roc(predLRstep,yTEST)),4)
  
  #Let's use the estimated model to predict on the new data
  predLRstepnew <- predict(LRstep,
                           newdata = NEWDATA,
                           type = "response")
  
  #Place predictions in the predictions dataframe
  PREDICTIONS$`Stepwise Logistic Regression` <- round(predLRstepnew,5)
  #Clean up memory
  rm(list = c('SLR','LRstep','predLRstep','predLRstepnew'))
}

if(RLogistic == 1){
  print("Beginning LASSO Logistic Regression estimation.")
  
  #Fit the Regularized Logistic Regression model
  RLR <- glmnet(x = data.matrix(TRAINDATA),
               y = yTRAIN,
               family = 'binomial')
  
  aucstore <- numeric()
  
  for(i in 1:length(RLR$lambda)){
    predglmnet <- predict(RLR,
                          newx = data.matrix(VALDATA),
                          type = "response",
                          s = RLR$lambda[i])
    aucstore[i] <- AUC::auc(roc(as.numeric(predglmnet),yVAL))
  }
  
  RLR.lambda <- RLR$lambda[which.max(aucstore)]
  
  RLR <- glmnet(x = data.matrix(TRAINBIGDATA),
               y = yTRAINBIG,
               family = "binomial")
  
  predRLR <- as.numeric(predict(RLR,
                                  newx = data.matrix(TESTDATA),
                                  type = "response",
                                  s=RLR.lambda))
  
  RESULTS$AUC_ROC[3] <- round(AUC::auc(roc(predRLR,yTEST)),4)
  
  #Now use the estimated model to predict on the new data
  predRLRnew <- as.numeric(predict(RLR,
                                   newx = data.matrix(NEWDATA),
                                   type = "response",
                                   s = RLR.lambda))
  
  #Place predictions in the predictions data frame
  PREDICTIONS$`Lasso Logistic Regression` <- round(predRLRnew, 5)
  #Clean up memory
  rm(list = c('RLR','aucstore','predglmnet','RLR.lambda','predRLR','predRLRnew'))
}

if(NNet == 1){
  print("Beginning Neural Network estimation.")
  #Load an external function to tune the neural net
  source("http://ballings.co/hidden/aCRM/code/chapter2/tuneMember.R")
  
  #First, we need to scale the data to range [0,1] to avoid scaling problems.
  TRAINDATAnumID <- sapply(TRAINDATA, is.numeric)
  TRAINDATAnum <- TRAINDATA[, TRAINDATAnumID]
  
  minima <- sapply(TRAINDATAnum, min)
  scaling <- sapply(TRAINDATAnum, max)-minima
  
  #Center is subtracted from each column. Because we use the minima this sets the minimum to zero. 
  #Each column is then divided by scale. Because we use the range this sets the maximum to one.
  TRAINDATAscaled <- data.frame(base::scale(TRAINDATAnum,
                                                 center = minima,
                                                 scale = scaling),
                                     TRAINDATA[,!TRAINDATAnumID])
  
  colnames(TRAINDATAscaled) <- c(colnames(TRAINDATA)[TRAINDATAnumID],
                                      colnames(TRAINDATA)[!TRAINDATAnumID])
  
  NN.rang <- 0.5 #the range of the initial random weights parameter
  NN.maxit <- 10000 #set high as to not run into early stopping
  NN.size <- c(5,10,20) #number of units in the hidden layer
  NN.decay <- c(0,0.001,0.01,0.1) #weight decay
  
  #Create the neural network.
  call<-call("nnet",
             formula = yTRAIN ~ .,
             data = TRAINDATAscaled,
             rang = NN.rang, maxit = NN.maxit,
             trace=FALSE, MaxNWts = Inf)
  
  tuning <- list(size=NN.size, decay=NN.decay)
  
  #Now to tune the neural net.
  #Let's scale the validation data
  VALnum <- VALDATA[, TRAINDATAnumID]
  VALDATAscaled <- data.frame(base::scale(VALnum,
                                                    center = minima,
                                                    scale = scaling),
                                        VALDATA[, !TRAINDATAnumID])
  colnames(VALDATAscaled) <- colnames(TRAINDATAscaled)
  
  result <- tuneMember(call = call,
                        tuning = tuning,
                        xtest = VALDATAscaled,
                        ytest = yVAL,
                        predicttype = "raw")
  
  #Now let's create the final model
  TRAINBIGDATAnum <- TRAINBIGDATA[, TRAINDATAnumID]
  TRAINBIGDATAscaled <- data.frame(base::scale(TRAINBIGDATAnum,
                                                    center=minima,
                                                    scale=scaling),
                                        TRAINBIGDATA[,!TRAINDATAnumID])
  colnames(TRAINBIGDATAscaled) <- c(colnames(TRAINBIGDATA)[TRAINDATAnumID],
                                         colnames(TRAINBIGDATA)[!TRAINDATAnumID])
  
invisible(capture.output(NN <- nnet(yTRAINBIG ~ .,
             TRAINBIGDATAscaled,
             size = result$size,
             rang = NN.rang,
             decay = result$decay,
             maxit = NN.maxit,
             trace = TRUE,
             MaxNWts = Inf)))  
  
  TESTBIGDATAnum <- TESTDATA[, TRAINDATAnumID]
  TESTDATAscaled <- data.frame(base::scale(TESTBIGDATAnum,
                                                center=minima,
                                                scale=scaling),
                                    TESTBIGDATAnum[,!TRAINDATAnumID])
  
  colnames(TESTDATAscaled) <- c(colnames(TESTBIGDATAnum)[TRAINDATAnumID],
                                     colnames(TESTBIGDATAnum)[!TRAINDATAnumID])
  
  predNN <- as.numeric(predict(NN, TESTDATAscaled, type="raw"))
  
  RESULTS$AUC_ROC[4] <- round(AUC::auc(roc(predNN,yTEST)),4)
  
  #Now let's scale the new data.
  NEWDATAnumID <- sapply(NEWDATA, is.numeric)
  NEWDATAnum <- NEWDATA[, NEWDATAnumID]
  
  minima <- sapply(NEWDATAnum, min)
  scaling <- sapply(NEWDATAnum, max)-minima
  
  #Set the minimum to 0 and the max to 1
  NEWDATAscaled <- data.frame(base::scale(NEWDATAnum,
                                            center = minima,
                                            scale = scaling),
                                NEWDATA[,!NEWDATAnumID])
  
  colnames(NEWDATAscaled) <- c(colnames(NEWDATA)[NEWDATAnumID],
                                 colnames(NEWDATA)[!NEWDATAnumID])
  
  #Predict the new data
  predNNnew <- as.numeric(predict(NN, NEWDATAscaled, type="raw"))
  
  #Place predictions in the predictions dataframe
  PREDICTIONS$`Neural Network` <- round(predNNnew, 5)
  
  #Clean up memory
  rm(list = c('TRAINDATAnumID','TRAINDATAnum','minima','scaling','TRAINDATAscaled',
              'NN.rang','NN.maxit','NN.size','NN.decay','call','tuning','VALnum',
              'VALDATAscaled','result','TRAINBIGDATAnum','TRAINBIGDATAscaled','NN',
              'TESTBIGDATAnum','TESTDATAscaled','predNN','NEWDATAnumID','NEWDATAnum',
              'NEWDATAscaled','predNNnew'))
}

if(KNN == 1){
  print("Beginning K-Nearest Neighbors estimation.")
  #The KNNX function requires all indicators to be numeric, so let's convert all data to the appropriate form.
  trainKNN <- data.frame(sapply(TRAINDATA, function(x) as.numeric(as.character(x))))
  
  trainKNNbig <- data.frame(sapply(TRAINBIGDATA, function(x) as.numeric(as.character(x))))
  
  valKNN <- data.frame(sapply(VALDATA, function(x) as.numeric(as.character(x))))
  
  testKNN <- data.frame(sapply(TESTDATA, function(x) as.numeric(as.character(x))))
  
  newdataKNN <- data.frame(sapply(NEWDATA, function(x) as.numeric(as.character(x))))
  
  #The distance function (Euclidean distance) is sensitive to the scale of the variables. 
  #Therefore we need to standardize the variables first.
  stdev <- sapply(trainKNN, sd)
  means <- sapply(trainKNN, mean)
  
  trainKNNbig <- data.frame(t((t(trainKNNbig)-means)/stdev))
  trainKNN <- data.frame(t((t(trainKNN)-means)/stdev))
  valKNN <- data.frame(t((t(valKNN)-means)/stdev))
  testKNN <- data.frame(t((t(testKNN)-means)/stdev))
  newdataKNN <- data.frame(t((t(newdataKNN)-means)/stdev))
  
  #We will start by evaluating the 10 nearest neighbors
  k <- 10
  
  #If we want to tune, it comes down to evaluating which value for k is best
  auc <- numeric()
  for(k in 1:nrow(trainKNN)){
    #retrieve the indicators of the k nearest neighbors of the query data
    indicatorsKNN <- as.integer(knnx.index(data=trainKNN,
                                           query=valKNN,
                                           k=k))
    #retrieve the actual y from the training set
    predKNN <- as.integer(as.character(yTRAIN[indicatorsKNN]))
    
    #if k > 1 then we take the proportion of 1s
    predKNN <- rowMeans(data.frame(matrix(data = predKNN,
                                          ncol = k,
                                          nrow = nrow(valKNN))))
    #Compute AUC
    auc[k] <- AUC::auc(roc(predKNN, yVAL))
  }
  #Retrieve the k that resulted in the maximum AUC
  k <- which.max(auc)
  
  #retrieve the indicators of the k nearest neighbors of the query data
  indicatorsKNN <- as.integer(knnx.index(data = trainKNNbig,
                                         query = testKNN,
                                         k=k))
  
  #retrieve the actual y from the training set
  predKNNoptimal <- as.integer(as.character(yTRAINBIG[indicatorsKNN]))
  
  #if k>1 then we take the proportion of 1s
  predKNNoptimal <- rowMeans(data.frame(matrix(data=predKNNoptimal,
                                               ncol = k,
                                               nrow = nrow(testKNN))))
  
  RESULTS$AUC_ROC[5] <- round(AUC::auc(roc(predKNNoptimal, yTEST)),4)
  
  #Now for the new data
  newindicatorsKNN <- as.integer(knnx.index(data = trainKNNbig,
                                         query = newdataKNN,
                                         k=k))
  
  #retrieve the actual y from the training set
  prednewKNNoptimal <- as.integer(as.character(yTRAINBIG[newindicatorsKNN]))
  
  #if k>1 then we take the proportion of 1s
  prednewKNNoptimal <- rowMeans(data.frame(matrix(data=prednewKNNoptimal,
                                               ncol = k,
                                               nrow = nrow(newdataKNN))))
  
  #Place predictions in predictions dataframe
  PREDICTIONS$`K-Nearest Neighbors` <- round(prednewKNNoptimal,5)
  
  #Clean up memory
  rm(list = c('trainKNN','trainKNNbig','valKNN','testKNN','newdataKNN','stdev', 
              'means','k','auc','indicatorsKNN','predKNN','predKNNoptimal',
              'newindicatorsKNN','prednewKNNoptimal'))
}

if(DT == 1){
  print("Beginning Decision Tree estimation.")
  
  #Estimate a tree model where we cross validate the cp parameter
  #cp sets the minimum amount the model needs to be improved for a split to be made
  candidates <- seq(0.00001, 0.2, by=0.001)
  aucstore <- numeric(length(candidates))
  j <- 0
  
  for(i in candidates) {
    j <- j+1
    tree <- rpart(yTRAIN ~ ., 
                  control = rpart.control(cp=i),
                  TRAINDATA)
    predTree <- predict(tree, VALDATA)[,2]
    aucstore[j] <- AUC::auc(roc(predTree,yVAL))
  }
  
  #Next we train the model on TRAINbig with the optimal cp and confirm the final performance on the test set
  
  tree <- rpart(yTRAINBIG ~ .,
                control = rpart.control(cp=candidates[which.max(aucstore)]),
                TRAINBIGDATA)
  predTREE <- predict(tree, TESTDATA)[,2]
  
  RESULTS$AUC_ROC[6] <- round(AUC::auc(roc(predTREE,yTEST)),4)
  
  #Classify the new data with the tree
  predTREEnew <- predict(tree, NEWDATA)[,2]
  
  #Place predictions in predictions dataframe
  PREDICTIONS$`Decision Tree`<-round(predTREEnew,5)
  
  #Clean up memory
  rm(list = c('candidates','aucstore','j','tree','predTREE','predTREEnew'))
}

if(BT == 1){
  print("Beginning Bagged Trees estimation.")
  
  #100 bagged trees is set as it is a typical default value
  ensemblesize <- 100
  ensembleoftrees <- vector(mode = 'list', length = ensemblesize)
  
  for(i in 1:ensemblesize){
    bootstrapsampleindicators <- sample.int(n=nrow(TRAINDATA),
                                            size=nrow(TRAINDATA),
                                            replace=TRUE)
    ensembleoftrees[[i]] <- rpart(yTRAIN[bootstrapsampleindicators]~.,
                                  TRAINDATA[bootstrapsampleindicators,])
  }
  
  baggedpredictions <- data.frame(matrix(NA, ncol = ensemblesize,
                                         nrow = nrow(TESTDATA)))
  
  for (i in 1:ensemblesize){
    
    baggedpredictions[,i] <- as.numeric(predict(ensembleoftrees[[i]],
                                                TESTDATA)[,2])
  }
  
  finalprediction <- rowMeans(baggedpredictions)
  
  RESULTS$AUC_ROC[7] <- round(AUC::auc(roc(finalprediction, yTEST)),4)
  
  #Predict new data
  baggedpredictionsnew <- data.frame(matrix(NA, ncol = ensemblesize,
                                         nrow = nrow(NEWDATA)))
  
  for (i in 1:ensemblesize){
    
    baggedpredictionsnew[,i] <- as.numeric(predict(ensembleoftrees[[i]],
                                                NEWDATA)[,2])
  }
  
  finalpredictionnew <- rowMeans(baggedpredictionsnew)
  
  #Place predictions in the predictions data frame
  PREDICTIONS$`Bagged Trees` <- round(finalpredictionnew,5)
  
  #Clean up memory
  rm(list = c('ensemblesize','ensembleoftrees','bootstrapsampleindicators','baggedpredictions',
              'finalprediction','finalpredictionnew'))
}

if(RF == 1){
  print("Beginning Random Forest estimation.")
  
  #When estimating a random forest model the number of trees must be specified
  #Inherent randomization between each tree make tuning the tree parameter infeasible.
  #Therefore, two random forests are fit with 500 and 1000 trees and the maximum AUC is selected.
  aucstore<-numeric()
  #Fit Random forest with 1000 trees
  rFmodelTHOU <- randomForest(x=TRAINBIGDATA,
                          y=yTRAINBIG,
                          ntree=1000,
                          importance=TRUE)
  #Predict test data with 1000 trees
  predrFTHOU <- predict(rFmodelTHOU, TESTDATA, type="prob")[,2]
  #Store the AUC
  aucstore[1] <- AUC::auc(roc(predrFTHOU,yTEST))
  #Fit with 500 trees
  rFmodelFIVE <- randomForest(x=TRAINBIGDATA,
                              y=yTRAINBIG,
                              ntree=500,
                              importance=TRUE)
  #Predict on test
  predrFFIVE <- predict(rFmodelFIVE, TESTDATA, type="prob")[,2]
  
  aucstore[2] <- AUC::auc(roc(predrFFIVE,yTEST))
  
  RESULTS$AUC_ROC[8] <- round(aucstore[which.max(aucstore)],4)
  
  #Predict new data based on the random forest with the highest AUC
  if(aucstore[which.max(aucstore)] == AUC::auc(roc(predrFTHOU,yTEST))){
    
    predrFnew <- predict(rFmodelTHOU, NEWDATA, type="prob")[,2]
    
  } else {predrFnew <- predict(rFmodelFIVE, NEWDATA, type="prob")[,2]}
  
  #Place predictions in predictions dataframe
  PREDICTIONS$`Random Forest` <- round(predrFnew,5)
  
  #Clean up memory
  rm(list = c('aucstore','rFmodelTHOU','predrFTHOU','rFmodelFIVE','predrFFIVE','predrFnew'))
}

if(AB == 1){
  print("Beginning Adaptive Boost estimation")
  #Fit the Adaptive Boost model
  ABmodel <- ada(x=TRAINDATA,
                 y=yTRAIN,
                 test.x = VALDATA,
                 test.y = yVAL,
                 iter = 150)
  #Pick the model with the lowest test error
  ABmodel <- ada(x=TRAINDATA,
                 y=yTRAIN,
                 iter=which.min(ABmodel$model$errs[,"test.err"]))
  #Predict on the test data
  predAB <- as.numeric(predict(ABmodel,
                               TESTDATA,
                               type="probs")[,2])
  
  RESULTS$AUC_ROC[9] <- round(AUC::auc(roc(predAB, yTEST)),4)
  
  #Predict the new data with the AB model
  predABnew <- as.numeric(predict(ABmodel,
                               NEWDATA,
                               type="probs")[,2])
  
  #Place predictions in predictions dataframe
  PREDICTIONS$`Adaptive Boosting` <- round(predABnew,5)
  
  #Clean up memory
  rm(list = c('ABmodel','predAB','predABnew'))
}

if(RotF == 1){
  print("Beginning Rotation Forest estimation.")
  #Train the rotation forest
  RoF <- rotationForest(x=TRAINBIGDATA,
                        y=as.factor(yTRAINBIG),
                        L=10)
  #Predict on test data
  predRoF <- predict(RoF, TESTDATA)
  
  #Assess performance
  RESULTS$AUC_ROC[10] <- round(AUC::auc(roc(predRoF,yTEST)),4)
  
  #Predict new data with Rotation Forest
  predRoFnew <- predict(RoF, NEWDATA)
  
  #Place predictions in predictions dataframe
  PREDICTIONS$`Rotation Forest` <- round(predRoFnew,5)
  
  #Clean up memory
  rm(list = c('RoF','predRoF','predRoFnew'))
}


#Remove the models not tested from the results data frame
RESULTS <- RESULTS[complete.cases(RESULTS),]

#The last arguments provided in the Macro Workbook specify the number of model AUC's
#and distinct prediction sets to return.

#Order the rows in RESULTS in decending order by AUC
RESULTS <- RESULTS[order(-RESULTS$AUC_ROC),]
rownames(RESULTS)<-c()

#Save only the results requested
RESULTS <- RESULTS[1:AUCR,]

#Return only the predictions requested
if(PR>1){PREDICTIONS <- PREDICTIONS[,colnames(PREDICTIONS) %in% as.character(RESULTS[1:PR,1])]} else{
  CN <- colnames(PREDICTIONS)[colnames(PREDICTIONS) %in% as.character(RESULTS[1:PR,1])]
  
  PREDICTIONS <- as.data.frame(PREDICTIONS[,colnames(PREDICTIONS) %in% as.character(RESULTS[1:PR,1])])
  
  colnames(PREDICTIONS) <- CN
}

#Add a descriptive end to the prediction column(s) for clarity
prefix <- "Prob("
suffix <- "=1)"
ending <- paste0(prefix,yvalue,suffix)

for(i in 1:ncol(PREDICTIONS)){colnames(PREDICTIONS)[i] <- paste(colnames(PREDICTIONS)[i],ending)}


#Write the predictions and results to a CSV file.
setwd("C:/R_code")

write.csv(RESULTS,'Model Performance.csv',row.names = F)
write.csv(PREDICTIONS, 'Model Predictions.csv',row.names =F)