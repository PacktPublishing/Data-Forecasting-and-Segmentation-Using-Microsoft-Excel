#install.packages("car")   
#library("car")

rng <- EXCEL$Application$get_Range( "A2:B49" )
X <- rng$get_Value()

x=X[,1]
y=X[,2]




x
y


# Compute the linear regression 
fit <- lm(x ~ y )
fit

dwtest (fit)