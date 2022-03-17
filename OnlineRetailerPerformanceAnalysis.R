# load packsges
library('openxlsx')
library('dplyr')
library('ggplot2')
library('corrplot')


################## load in csv files ################## 

addsToCart <- read.csv('DataAnalyst_Ecom_data_addsToCart.csv')
sessionCounts <- read.csv('DataAnalyst_Ecom_data_sessionCounts.csv')


################## early exploration ################## 

summary(addsToCart)
summary(sessionCounts)

unique(sessionCounts$dim_browser)
unique(sessionCounts$dim_deviceCategory)

hist(sessionCounts$sessions)
hist(sessionCounts$transactions)
hist(sessionCounts$QTY)

plot(density(log(sessionCounts$sessions)), main = "Desnity Plot of log(Sessions)")
plot(density(log(sessionCounts$transactions)), main = "Desnity Plot of log(Transactions)")
plot(density(log(sessionCounts$QTY)), main = "Desnity Plot of log(QTY)")


################## neccesary cleaning ################## # remove rows where browser isn't recoreded

sessionCountsClean <- sessionCounts %>% filter(dim_browser != 'error' & dim_browser !='(not set)') %>% 
  mutate(dim_date = as.Date(dim_date, format='%m/%d/%y'))


################## first dataframe ################## 

sheet1.MonthDevice <- data.frame()

# checking if aggregate function works as I expect it to
aggregate(sessions ~ dim_deviceCategory + format(dim_date, '%m'), data = sessionCountsClean, FUN = sum)

sessionCountsClean %>% filter(dim_deviceCategory == 'desktop') %>% mutate(m = format(dim_date, '%m')) %>% 
  filter(m == '01') %>% summarise(sum(sessions))

# both answered: 393556     so I expect the function should work


 
s <- aggregate(sessions ~ dim_deviceCategory + format(dim_date, '%m') + format(dim_date, '%y'), data = sessionCountsClean, FUN = sum) # count all sessions
t <- aggregate(transactions ~ dim_deviceCategory + format(dim_date, '%m') + format(dim_date, '%y'), data = sessionCountsClean, FUN = sum) # count all transactions
q <- aggregate(QTY ~ dim_deviceCategory + format(dim_date, '%m') + format(dim_date, '%y'), data = sessionCountsClean, FUN = sum) # count all qunatity
    

# populate dataframe

sheet1.MonthDevice <- cbind(s,t$transactions,q$QTY) %>% 
  # rename columns into human readable format
  rename(Device = dim_deviceCategory,
         Month = 'format(dim_date, "%m")',
         Year = 'format(dim_date, "%y")',
         Sessions = sessions,
         Transactions = 't$transactions',
         QTY = 'q$QTY') %>%
  # create ECR column
  mutate(ECR = Transactions/Sessions)  %>% 
  # convert into date format
  mutate(Month = as.numeric(Month)) %>% 
  mutate(Month = month.name[Month]) %>%
  mutate(Year = as.numeric(Year))


################## second sheet data ##################

current.month = sheet1.MonthDevice$Month[length(sheet1.MonthDevice$Month)]
prev.month = sheet1.MonthDevice$Month[length(sheet1.MonthDevice$Month) - 3]
current.year = sheet1.MonthDevice$Year[length(sheet1.MonthDevice$Year)]
prev.year = sheet1.MonthDevice$Year[length(sheet1.MonthDevice$Year) - 3]

current <- c(sheet1.MonthDevice %>% filter(Month == current.month & Year == current.year) %>% summarise('Total Sessions' = sum(Sessions), 'Total Transactions' = sum(Transactions), 'Total Quantity' = sum(QTY)), 
             sheet1.MonthDevice %>% filter(Month == current.month & Year == current.year & Device == 'mobile') %>% summarise('Mobile Sessions' = sum(Sessions), 'Mobile Transactions' = sum(Transactions), 'Mobile Quantity' = sum(QTY)),
             sheet1.MonthDevice %>% filter(Month == current.month & Year == current.year & Device == 'desktop') %>% summarise('Desktop Sessions' = sum(Sessions), 'Desktop Transactions' = sum(Transactions), 'Desktop Quantity' = sum(QTY)),
             sheet1.MonthDevice %>% filter(Month == current.month & Year == current.year & Device == 'tablet') %>% summarise('Total Sessions' = sum(Sessions), 'Total Transactions' = sum(Transactions), 'Total Quantity' = sum(QTY)))

prev <- c(sheet1.MonthDevice %>% filter(Month == prev.month & Year == prev.year) %>% summarise('Total Sessions' = sum(Sessions), 'Total Transactions' = sum(Transactions), 'Total Quantity' = sum(QTY)), 
             sheet1.MonthDevice %>% filter(Month == prev.month & Year == prev.year & Device == 'mobile') %>% summarise('Mobile Sessions' = sum(Sessions), 'Mobile Transactions' = sum(Transactions), 'Mobile Quantity' = sum(QTY)),
             sheet1.MonthDevice %>% filter(Month == prev.month & Year == prev.year & Device == 'desktop') %>% summarise('Desktop Sessions' = sum(Sessions), 'Desktop Transactions' = sum(Transactions), 'Desktop Quantity' = sum(QTY)),
             sheet1.MonthDevice %>% filter(Month == prev.month & Year == prev.year & Device == 'tablet') %>% summarise('Tablet Sessions' = sum(Sessions), 'Tablet Transactions' = sum(Transactions), 'Tablet Quantity' = sum(QTY)))


absolute <- as.numeric(current) - as.numeric(prev)
relative <- absolute / as.numeric(prev) * 100

################## create workbook ################## 

wb <- createWorkbook(creator = 'Dean Stuart', title = 'Online Retailer Performance Analysis')

addWorksheet(wb, 'Sheet 1')
writeData(wb,1,sheet1.MonthDevice)


addWorksheet(wb, 'Sheet 2')

# changing tpyes for easier formatting in xlsx file
prev <- as.data.frame(prev) # useful to get column names in xlsx file
absolute <- as.list(absolute) # enter list as a row rather then column
relative <- as.list(relative)

writeData(wb,2,prev,startCol = 2)
writeData(wb,2,current,startCol = 2,startRow = 3)
writeData(wb,2,absolute,startCol = 2,startRow = 6)
writeData(wb,2,relative,startCol = 2,startRow = 7)

# write row labels
writeData(wb,2,paste(prev.month , prev.year, sep = " "), startCol = 1, startRow = 2)
writeData(wb,2,paste(current.month , current.year, sep = " "), startCol = 1, startRow = 3)
writeData(wb,2,'Difference', startCol = 1, startRow = 5)
writeData(wb,2,'Absolute', startCol = 1, startRow = 6)
writeData(wb,2,'Relative', startCol = 1, startRow = 7)

# save the workbook as an excel file
saveWorkbook(wb, 'performanceAnalysis.xlsx', overwrite = TRUE)



################## Visualisations ##################

# time v sessions
ggplot(sheet1.MonthDevice %>% filter(Device=='mobile'), aes(x=as.Date(paste(match(Month,month.name),'1', Year, sep = "/"), '%m/%d/%y'), 
                                                            y=Sessions)) + 
  geom_point() +
  xlab('Date') +
  ggtitle('Sessions Over Time on Mobile Devices')

ggplot(sheet1.MonthDevice %>% filter(Device=='desktop'), aes(x=as.Date(paste(match(Month,month.name),'1', Year, sep = "/"), '%m/%d/%y'), 
                                                            y=Sessions)) + 
  geom_point() +
  xlab('Date') +
  ggtitle('Sessions Over Time on Desktops')

ggplot(sheet1.MonthDevice %>% filter(Device=='tablet'), aes(x=as.Date(paste(match(Month,month.name),'1', Year, sep = "/"), '%m/%d/%y'), 
                                                             y=Sessions)) + 
  geom_point() +
  xlab('Date') +
  ggtitle('Sessions Over Time on Tablets')


# time v ECR by color
ggplot(sheet1.MonthDevice, aes(x=as.Date(paste(match(Month,month.name),'1', Year, sep = "/"), '%m/%d/%y'), 
                                                            y=ECR, color = Device)) + 
  geom_point() +
  geom_line() +
  xlab('Date') +
  ggtitle('ECR Over Time (fig. 2)')

# sessions over time by color
ggplot(sheet1.MonthDevice, aes(x=as.Date(paste(match(Month,month.name),'1', Year, sep = "/"), '%m/%d/%y'), 
                               y=Sessions, color = Device)) + 
  geom_point() +
  geom_line() +
  xlab('Date') +
  ggtitle('Sessions over Time (fig. 5)')

ggplot(sheet1.MonthDevice, aes(x=as.Date(paste(match(Month,month.name),'1', Year, sep = "/"), '%m/%d/%y'), 
                               y=Transactions, color = Device)) + 
  geom_point() +
  geom_line() +
  xlab('Date') +
  ggtitle('Transactions over Time (fig. 1)')

ggplot(sheet1.MonthDevice, aes(x=as.Date(paste(match(Month,month.name),'1', Year, sep = "/"), '%m/%d/%y'), 
                               y=QTY, color = Device)) + 
  geom_point() +
  geom_line() +
  xlab('Date') +
  ggtitle('QTY over Time')

# sessions v transactions  
ggplot(sheet1.MonthDevice, aes(x=Sessions, y=Transactions, color = Device)) + 
  geom_point() +
  ggtitle('Sessions v Transactions (fig. 3)') +
  geom_smooth(method='lm', formula = y~x, se=F)

# correlations plots
corrplot(cor(sessionCountsClean %>% filter(dim_deviceCategory == 'mobile') %>% select(sessions,transactions,QTY)), method = 'number', 
         type = 'upper', title = "Correlation Matrix on Mobile")

corrplot(cor(sessionCountsClean %>% filter(dim_deviceCategory == 'desktop') %>% select(sessions,transactions,QTY)), method = 'number', 
         type = 'upper', title = "Correlation Matrix on Desktop")

corrplot(cor(sessionCountsClean %>% filter(dim_deviceCategory == 'tablet') %>% select(sessions,transactions,QTY)), method = 'number', 
         type = 'upper', title = "Correlation Matrix on Tablet")

# session %change over time

p_change <- sheet1.MonthDevice %>%
  group_by(Device) %>%
  mutate(lag = lag(Sessions)) %>%
  mutate(pct.change = (Sessions - lag) / lag) %>% select(Device, Month, Year, pct.change)

ggplot(p_change, aes(x=as.Date(paste(match(Month,month.name),'1', Year, sep = "/"), '%m/%d/%y'), 
                               y=pct.change, color = Device)) + 
  geom_point() +
  geom_line() +
  xlab('Date') +
  ylab('% change from prev. month') +
  ggtitle('Monthly % change in Sessions (fig. 4)')

# linear models

mobile <- lm(Transactions ~ Sessions, data = sheet1.MonthDevice %>% filter(Device == 'mobile'))
summary(mobile)
mobile$coefficients[2]

desktop <- lm(Transactions ~ Sessions, data = sheet1.MonthDevice %>% filter(Device == 'desktop'))
summary(desktop)
desktop$coefficients[2]

tablet <- lm(Transactions ~ Sessions, data = sheet1.MonthDevice %>% filter(Device == 'tablet'))
summary(tablet)
tablet$coefficients[2]

