from __future__ import division
import xlrd
import numpy as np
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import style
style.use("ggplot")
from sklearn.cluster import KMeans
from xlrd.sheet import ctype_text 

book = xlrd.open_workbook('AGN01.xlsx')
target_column='AGN020202D'
#print book.nsheets
sheet_names= book.sheet_names()


sh = book.sheet_by_index(0)
#print sh.name, sh.nrows, sh.ncols

keys = [sh.cell(0, col_index).value for col_index in xrange(sh.ncols)]
#print keys
i=0
while not sh.cell_value(0,i)==target_column:
	i=i+1
col_having_class=i
#print sh.cell_type(3,2)#returns 1 for string 2 for non string 
d=[]
arr=[]
header=[]
total=0
sum_of_column=0
no_of_categotical_columns=0
for count in xrange(sh.ncols):
	#if sh.cell_type(1,count)==1:
		#print 1
	if sh.cell_type(1,count)==2:
		for row_num in range(1,sh.nrows):
			sum_of_column=sum_of_column+sh.cell_value(row_num,count)
		if 	sum_of_column>sh.nrows:
			for row_num in range(1,sh.nrows):
				d.append(sh.cell_value(row_num,count))
			arr.append(d)
			d=[]
			header.append(count)
			total=total+1  ##counts the total number of columns having non categorical data
			if count==col_having_class:
				col_in_arr_having_class=total
		else:
			no_of_categotical_columns=no_of_categotical_columns+1
		sum_of_column=0
first_data=[]
#print arr
#print header
#print total
#print col_in_arr_having_class
count=0
while count<=total-1:
	#print ("hi")
	if col_in_arr_having_class-1==count:
		count=count+1
	if count>total-1:
		break
	#print count
	correlation = np.corrcoef(arr[col_in_arr_having_class-1],arr[count])[0,1];#print correlation;
	if correlation<0:
		print("%s tends to increase when %s decrease and vice versa."%(sh.cell_value(0,col_having_class),sh.cell_value(0,header[count])))
	if correlation>=0.5:
		b=sh.cell_value(0,header[count]);
		a=sh.cell_value(0,col_having_class);
		print("%s tends to increase when %s increases and is strongly correlated."%(a,b))
	if correlation<0.5 and correlation>0:
		b=sh.cell_value(0,header[count]);
		a=sh.cell_value(0,col_having_class);
		print("%s tends to increase when %s increases and is weakly correlated."%(a,b))
	count=count+1
count=0
print "\n"
	#print count
#print col_in_arr_having_class
#print header
#print sh.nrows
#print sh.cell_value(0,header[1])
#print sh.cell_value(0,header[0])
i=0
#print("result of clustering with respect to %s"%(sh.cell_value(0,col_having_class))
#print col_having_class
while not sh.cell_value(0,i)==target_column:
	i=i+1
col_having_class=i#rename it to something else
no_of_labels_for_two=0
j=0
print("result of clustering with respect to %s"%(sh.cell_value(0,col_having_class)))
print "\n"
while j<total:
	if j>total:
		break
	'''if col_in_arr_having_class-1==j and j==total-1:
		for count in range(0,sh.nrows-1):
			point=[arr[col_in_arr_having_class-1][count] , arr[j][count]]
			first_data.append(point)'''	
	if col_in_arr_having_class-1==j and not j==total-1:
		for count in range(0,sh.nrows-1):
			point=[arr[col_in_arr_having_class-1][count] , arr[j+1][count]]
			first_data.append(point)
	else:
		for count in range(0,sh.nrows-1):
			point=[arr[col_in_arr_having_class-1][count] , arr[j][count]]
			first_data.append(point)
	#print first_data
	X=first_data
	#print X
	#n=int(1/(sh.cell_value(1,col_having_one_hot) - sh.cell_value(2,col_having_one_hot)))
	n=2
	kmeans = KMeans(n_clusters=n)
	kmeans.fit(X)
	centroids = kmeans.cluster_centers_
	labels = kmeans.labels_
	#print(centroids)
	#print(labels)
	colors = ["g.","r.","c.","y."]
	for i in range(len(X)):
		#print("coordinate:",X[i], "label:", labels[i])
		if labels[i]==1:
			no_of_labels_for_two=no_of_labels_for_two+1
		plt.plot(X[i][0], X[i][1], colors[labels[i]], markersize = 10)
		'''if labels[i]==2:
			no_of_labels_for_three=no_of_labels_for_three+1
		plt.plot(X[i][0], X[i][1], colors[labels[i]], markersize = 10)
		if labels[i]==1:
			no_of_labels_for_four=no_of_labels_for_four+1
		plt.plot(X[i][0], X[i][1], colors[labels[i]], markersize = 10)'''

	
	plt.scatter(centroids[:, 0],centroids[:, 1], marker = "x", s=150, linewidths = 5, zorder = 10)
	xlabel=sh.cell_value(0,col_having_class)
	#if j==col_in_arr_having_class-1:
		#j=j+1
	ylabel=sh.cell_value(0,header[j])
	plt.xlabel(xlabel)
	plt.ylabel(ylabel)
	#plt.show()
	#j=j+1
	print("cluster 1 has average %s of %f and cluster 2 has average %s of %f"%(sh.cell_value(0,header[j]),centroids[0][1],sh.cell_value(0,header[j]),centroids[1][1]))
	per= float(no_of_labels_for_two/(sh.nrows-1))
	print("percentage of points of %s in cluster 1 is %f and in cluster 2 is %f"%(sh.cell_value(0,header[j]),(1-per)*100,per*100))
	'''for count in range(1,n+1):
		print("%f percentage of %s has %s of average %f" % (centroids[count-1][0]*100 ,sh.cell_value(count,col_having_class) , sh.cell_value(0,1+header[j]) ,centroids[count-1][1] ))
	if n==2:
		if centroids[0][1] > centroids[1][1]:
			print("%s has higher %s than %s " % (sh.cell_value(1,col_having_class) , sh.cell_value(0,1+header[j]) , sh.cell_value(2,col_having_class)))
		elif centroids[0][1] < centroids[1][1]:
			print("%s has lesser %s than %s " % (sh.cell_value(1,col_having_class) , sh.cell_value(0,1+header[j]) , sh.cell_value(2,col_having_class)))
		else:
			print("%s has same %s as that of %s " % (sh.cell_value(1,col_having_class) , sh.cell_value(0,1+header[j]) , sh.cell_value(2,col_having_class)))
	'''
	j=j+1
	no_of_labels_for_two=0
	first_data=[]
	print "\n"
print("no of categorical column is %d"%(no_of_categotical_columns))
