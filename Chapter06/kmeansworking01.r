
 #Hoja2
 rng <- EXCEL$Application$get_Range( "F1:F1" )
numclusters <- rng$get_Value()
numclusters  
   

#rng <- EXCEL$Application$get_Range( "B2:B134" )
#rng <- EXCEL$Application$get_Range( "B2:B9233" )
rng <- EXCEL$Application$get_Range( "B2:E35" )
X <- rng$get_Value()
X

opt_nb_clusters = numclusters

set.seed(124)
kmeans <- kmeans(X, opt_nb_clusters, iter.max = 300, nstart = 50)


#rng <- EXCEL$Application$get_Range( "C2:C134" )
#rng <- EXCEL$Application$get_Range( "C2:C9233" )
rng <- EXCEL$Application$get_Range( "F2:F35" )

rng$put_Value(kmeans$cluster)

kmeans
kmeans$cluster
