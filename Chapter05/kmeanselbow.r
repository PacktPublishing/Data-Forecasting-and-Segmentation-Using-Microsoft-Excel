#Hoja2

rng <- EXCEL$Application$get_Range( "B2:E35" )
#rng <- EXCEL$Application$get_Range( "B2:B9233" )
#rng <- EXCEL$Application$get_Range( "B2:B103" )
X <- rng$get_Value()
X
set.seed(123)
# Compute and plot wss for k = 2 to k = 15.
k.max <- 15
data <- X
data
wss <- sapply(1:k.max, 
              function(k){kmeans(data, k, nstart=50,iter.max = 15 )$tot.withinss})

nb_clusters = seq(1, length(wss), 1)


class.df <- data.frame (nb_clusters,wss)

plot(class.df) 
