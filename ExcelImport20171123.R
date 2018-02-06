
setwd("Z:/Surveys/Crime/2017/Data")

#Remove the below #'s if this is the first time you are running this file on your computer
#packages=c("tidyverse")
#install.packages(packages)

require(readxl)

#Get list of files from Sharepoint folder
files=list.files()

#nominate coordinates of answers to be extracted from each Excel workbook
vars=c("G7","F12","F14","F16","F18","F20","F22","F24","F26","F28","F30","G32","F35","F40","F43","C10","C15","C20","C25","D31","D33","D35","D37","D39","D41","D43","D45","D50","C56","G18","G20","G26","G28","G34","G36","G42","G44","G50","G52","G58","G60","G65","G71","G15","G17","G19","H24","G30","G32","G34","G36","G38","G40","G42","G44","G46","G52","C17","G23","G25","G27","G29","G31","G33","G35","G37","G42","G44","G46","G48","G50","G52","H54","H59","G23","G25","G27","G29","G31","G33","G35","G37","G39","G41","G43","G48","H53","G59","G65","H72")

#create an empty dataframe with specific variable types
file=data.frame(ID=numeric(),
                   name=character(),
                   catfood=character(),
                   catdiy=character(),
                   catfurn=character(),
                   catelec=character(),
                   catcloth=character(),
                   catfoot=character(),
                   catbook=character(),
                   catenter=character(),
                   cathealth=character(),
                   catdepart=character(),
                   catother=character(),
                   outlets=numeric(),
                   turnover=numeric(),
                   employees=numeric(),
                   crimetot=numeric(),
                   cyberperc=numeric(),
                   totnoncyb=numeric(),
                   totcybsec=numeric(),
                   ththeft=numeric(),
                   thinside=numeric(),
                   thfraud=numeric(),
                   thcyber=numeric(),
                   thdata=numeric(),
                   throbb=numeric(),
                   thburg=numeric(),
                   thterr=numeric(),
                   policeresp=character(),
                   areaissue=character(),
                   noinctheft=numeric(),
                   valtheft=numeric(),
                   noincemploy=numeric(),
                   valemploy=numeric(),
                   noincother=numeric(),
                   valother=numeric(),
                   noincrobb=numeric(),
                   valrobb=numeric(),
                   noincburg=numeric(),
                   valburg=numeric(),
                   noincdam=numeric(),
                   validam=numeric(),
                   theftperc=numeric(),
                   gangtheft=character(),
                   vioinj=numeric(),
                   vionoinj=numeric(),
                   vioabuse=numeric(),
                   viotrends=character(),
                   inthitting=numeric(),
                   intglass=numeric(),
                   intknife=numeric(),
                   intgun=numeric(),
                   intstone=numeric(),
                   intsyringe=numeric(),
                   intsword=numeric(),
                   intdogs=numeric(),
                   intother=numeric(),
                   racial=character(),
                   totfraud=numeric(),
                   fracard=numeric(),
                   fracredit=numeric(),
                   fraaccount=numeric(),
                   frarefund=numeric(),
                   fravouch=numeric(),
                   frainside=numeric(),
                   frasupply=numeric(),
                   fraemail=numeric(),
                   afrsreport=character(),
                   afrsbulk=character(),
                   afrsspeed=character(),
                   afrsflow=character(),
                   afrspolice=character(),
                   afrsnouse=character(),
                   afrsother=character(),
                   supplypatt=character(),
                   secdos=character(),
                   secphish=character(),
                   secpharm=character(),
                   secspoof=character(),
                   secdox=character(),
                   secmal=character(),
                   secwhale=character(),
                   secransom=character(),
                   secsocial=character(),
                   secdata=character(),
                   secwebapp=character(),
                   cispvalue=character(),
                   cyberresp=character(),
                   nocyberatt=character(),
                   recstaff=character(),
                   brctool=character(),
                   stringsAsFactors = FALSE)

#This loop names each variable again after read_excel imports data from coordinates... ADD 1 TO "j in 1:#" FOR EACH NEW SUBMISSION
for (j in 1:15) {
  file=c(1)

#reads data from each coordinate in vars list for each specific sheet within workbook, replacing any missing values with NA_integer_ so that the types of variable do not change away from numeric
for (i in 1:15){
  tryCatch({
    a=read_excel(files[j],'1. Company Information',vars[i],col_names=FALSE)
  },error=function(e){})
  if(dim(a)==0){a=NA_integer_}
  file=cbind.data.frame(file, a)
}
for (i in 16:29){
  tryCatch({
    a=read_excel(files[j],'2. Overall Crime',vars[i],col_names=FALSE)
  },error=function(e){})
  if(dim(a)==0){a=NA_integer_}
  file=cbind.data.frame(file, a)
}
for (i in 30:43){
  tryCatch({
    a=read_excel(files[j],'3. Theft & Damage',vars[i],col_names=FALSE)
  },error=function(e){})
  if(dim(a)==0){a=NA_integer_}
  file=cbind.data.frame(file, a)
}
for (i in 44:57){
  tryCatch({
    a=read_excel(files[j],'4. Violence & Abuse',vars[i],col_names=FALSE)
  },error=function(e){})
  if(dim(a)==0){a=NA_integer_}
  file=cbind.data.frame(file, a)
}
for (i in 58:73){
  tryCatch({
    a=read_excel(files[j],'5. Fraud',vars[i],col_names=FALSE)
  },error=function(e){})
  if(dim(a)==0){a=NA_integer_}
  file=cbind.data.frame(file, a)
}
for (i in 74:90){
  tryCatch({
    a=read_excel(files[j],'6. Cyber Security',vars[i],col_names=FALSE)
  },error=function(e){})
  if(dim(a)==0){a=NA_integer_}
  file=cbind.data.frame(file, a)
}
names(file)=c("ID","name","catfood","catdiy","catfurn","catelec","catcloth","catfoot","catbook","catenter","cathealth","catdepart","catother","outlets","turnover","employees","crimetot","cyberperc","totnoncyb","totcybsec","ththeft","thinside","thfraud","thcyber","thdata","throbb","thburg","thterr","policeresp","areaissue","noinctheft","valtheft","noincemploy","valemploy","noincother","valother","noincrobb","valrobb","noincburg","valburg","noincdam","valdam","theftperc","gangtheft","vioinj","vionoinj","vioabuse","viotrends","inthitting","intglass","intknife","intgun","intstone","intsyringe","intsword","intdogs","intother","racial","totfraud","fracard","fracredit","fraaccount","frarefund","fravouch","frainside","frasupply","fraemail","afrsreport","afrsbulk","afrsspeed","afrsflow","afrspolice","afrsnouse","afrsother","supplypatt","secdos","secphish","secpharm","secspoof","secdox","secmal","secwhale","secransom","secsocial","secdata","secwebapp","cispvalue","cyberresp","nocyberatt","recstaff","brctool")
assign(paste0("file", j),file)
}

#binds each file by row into one usable dataframe... ADD FILE# FOR EACH NEW SUBMISSION
output = rbind(file1,file2,file3,file4,file5,file6,file7,file8,file9,file10,file11,file12,file13,file14,file15)

setwd("Z:/Surveys/Crime/2017/Code")

#create .csv file for analysis
write.csv(output, "CrimeSurvey2017ROutput.csv")
