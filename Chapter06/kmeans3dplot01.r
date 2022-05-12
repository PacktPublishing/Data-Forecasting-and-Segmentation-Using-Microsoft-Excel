  
#install.packages("scatterplot3d")   
library("scatterplot3d")

rng <- EXCEL$Application$get_Range( "E2:G38073" )
X <- rng$get_Value()


scatterplot3d(x=X[,1],y=X[,2],z=X[,3])
