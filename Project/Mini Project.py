#!/usr/bin/env python
# coding: utf-8

# # Kanishka Mishra     SIT, Pune
# # Email: kanishka.mishra.btech2019@sitpune.edu.in  
# # Mobile Number: 7566744897

# # Task 1: Identifying and importing essential libraries

# In[2]:


import pandas as pd
import matplotlib.pyplot as plt
get_ipython().run_line_magic('matplotlib', 'inline')

import numpy as np


# 
# # Task 2: Data loading and overview
# a.Load the dataset from a file
# b.Look at the first 5 records of the data
# c.Print the columns contained in the data
# d.Print the attributes like count , mean , max , min, standard deviation etc e.Find out are there any missing or null values in the data for all columns

# In[3]:


dataf=pd.read_excel('Superstore.xls')  #Loaded the dataset from the file
dataf  


# In[4]:


df=(dataf.loc[[0,1,2,3,4],['Row ID','Order ID','Order Date','Ship Date','Ship Mode','Customer ID','Customer Name','Segment','Country','City','Postal Code','Region','Product ID','Category','Sub-Category','Product Name','Sales','Quantity','Discount','Profit']]) #First 5 records of the data
df
# print(df)    #printing all the columns of first 5 records


# In[6]:


df1=(dataf.loc[[0,1,2,3,4],['Sales','Profit']])  #sales and profit column of first 5 records
df1


# In[7]:


print(df1.mean()) #mean of first 5 records


# In[27]:


print(df1.std()) #standard deviation of first 5 records


# In[28]:


print(df1.max())   #finding maximum value


# In[29]:


print(df1.min())  #finding minimun value


# In[30]:


print(df1.count())  #count


# In[37]:


print(dataf.isnull())    #Finding all the missing values in the data


# In[40]:


print(dataf.isnull().values.any())   #To check if there are any missing values in the data True:Missing values found False: No missing values


# # Task 3: Find out the per unit price from the data
# a. Convert date time field into pandas time object
# b. Find out the price per unit from the data and create a column for it in the same data

# In[5]:


dataf['Order Date']=pd.to_datetime(dataf['Order Date'])
dataf['Ship Date']=pd.to_datetime(dataf['Ship Date'])
dataf


# In[6]:


#creates the list of price per unit
templist=[]
for i in range (0,9994):
    df3=dataf.iloc[i]
    unit=df3['Sales']/df3['Quantity']
    templist.append(unit)   
#adds a coulmn named 'Price Per unit' into the record
dataf['Price Per Unit']=templist   
dataf
#print(dataf)     


# # Task 4: Find out the monthly revenue and analyze the findings
# a.Create a column with Year and month field only give a suitable name to it
# b.Create a separate dataset from the data which will have two columns , one which is created in step a of this task and second the monthly sales / revenue
# c.Plot this dataset
# 1
# d.Describe your analysis for the monthly revenue observations.

# In[6]:


#Creating a column of year and month
y_m = pd.to_datetime(dataf['Order Date']).dt.to_period('M') 
dataf['Year & Month']=y_m
dataf.head()
dataf


# In[7]:


#Creating a separate datasheet for year and month and monthly sales

#Column for Year and month
dataf1=pd.DataFrame(dataf['Year & Month'])

#Column for Monthly sales
dataf1['Monthly Sales']=dataf['Sales']

# dataf1
group=dataf1.groupby([(dataf1['Year & Month'].dt.year),(dataf1['Year & Month'].dt.month)]).sum()
dframe1=pd.DataFrame(group)
dframe1


# In[9]:


#Plotting the Data

from matplotlib import style
#using ggplot here
style.use("ggplot")

#Quantities to be plotted

x=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

#Sales for the year 2014
y1=[14236.895,4519.892 , 55691.009 , 28295.345  ,23648.287 , 34595.1276,
 33946.393  ,27909.4685, 81777.3508 ,31453.393 , 78628.7167 ,69545.6205]

#Sales for the year 2015
y2=[18174.0756,11951.4110,38726.2520,34195.2085,30131.6865,24797.2920,
    28765.3250,36898.3322,64595.9180,31404.9235,75972.5635,74919.5212]

#Sales for the year 2016
y3=[18542.4910,22978.8150,51715.8750,38750.0390,56987.7280,40344.5340,
    39261.9630,31115.3743,73410.0249,59687.7450,79411.9658,96999.0430]

#Sales for the year 2017
y4=[43971.3740,20301.1334,58872.3528,36521.5361,44261.1102,52981.7257,
    45264.4160,63120.8880,87866.6520,77776.9232,118447.8250,83829.3188]

#Plotting 2014 Monthly sales
plt.plot(x,y1,'g',label='2014', linewidth=2)   #g is the color(green)

#Plotting 2015 Monthly Sales
plt.plot(x,y2,'r',label='2015',linewidth=2)

#Plotting 2014 Monthly sales
plt.plot(x,y3,'b',label='2016', linewidth=2)

#Plotting 2014 Monthly sales
plt.plot(x,y4,"yellow",label='2017', linewidth=2)

#Giving the title for our plot
plt.title('Monthly Sales')

#Giving the names to axes
plt.ylabel('Sales')
plt.xlabel('Months')

#Giving description about the plot curves
plt.legend()

#Defining the the background of our plot
plt.grid(True, color = 'k')   #k gives black color

#Showing what we have plotted
plt.show()


# In[11]:


#Description of the monthly growth analysis

# 1. This graph depicts, compares, and provide a clear monthly growth report for all the years.
# 2. Through this graph we can find which month had the most growth and in which year.
# 3. This graph tells us the pattern of monthly growth in different years ,for instance, in every year october saw a decline in growth


# # Task 5: Find out the yearly revenue and analyze the findings
# a.Create a column with Year field only give a suitable name to it
# b.Create a separate dataset from the data which will have two columns , one which is created in step a of this task and second the yearly sales / revenue
# c.Plot this dataset
# d.Describe your analysis for the yearly revenue observations.

# In[12]:


#Creating a column with only year field and Sales 
year=pd.DataFrame(dataf1['Year & Month'].dt.year)   #Extracting the year from the date time format
# year
year.columns=['Year']
year['Sales']=dataf['Sales']
year


# In[13]:


#Grouping the data of various years together and finding the total sales in the particular year
group1=dataf1.groupby([(dataf1['Year & Month'].dt.year)]).sum()
group1


# In[128]:


#Bar graph

#Quantities for plotting
xx = ['2014','2015','2016','2017']
yy = [484247.4981,470532.5090,609205.5980,733215.2552]

#Plotting these in the form of bar graph
plt.bar(xx ,yy , color = "royalblue",width=0.5)

plt.title("Yearly Sales")

#Setting up name for x axis
plt.xlabel("Years")

#Setting up name for y axis
plt.ylabel("Sales")

#Showing what we have plotted
plt.show()


# In[ ]:


#Description of the analysis

# 1. This bar graph helps us compare the yearly growth in a clear manner.
# 2. This shows how the company's sales has grown through the years.
#3. Over the years the sales has increased.


# # Task 6: Finding out the monthly growth rate and analyse the findings
# a.Create a column in the monthly revenue dataset for monthly growth rate
# b.Find out the maximum monthly growth rate c.Plot the findings
# d.Describe the findings

# In[14]:


#Finding the monthly growth rate using for loop and storing it in a list
ratelist=[]
for i in range (0,47):
    v1=dframe1['Monthly Sales'].values[i+1]
    v2=dframe1['Monthly Sales'].values[i]
    rate=((v1-v2)/v2)*100  #Formula to calculate monthly growth rate
    ratelist.append(rate)
ratelist  #Growth rate list

#Creating a new coloumn named Monthly growth
dframe1['Monthly Growth']=[np.NaN,-68.25226287052057, 1132.1314093345597, -49.19225650948436, -16.42340109300661, 46.29020529055669, -1.875219561265729,
                           -17.78369943457617, 193.00934483936877, -61.53776969747472, 149.9848480575689, -11.55188152778284, -73.867404634056, 
                           -34.239235804653504, 224.03079435557856, -11.70018596170887, -11.8833081541234, -17.70360414442781, 16.001880366614195,
                           28.273649611120327, 75.06460088729987, -51.382495253028225, 141.91290738218171, -1.3860823585345665, -75.25012079228294,
                           23.925178121968667, 125.05892927899032, -25.071288071602865, 47.06495650236645, -29.204873722988232, -2.6833151673037325,
                           -20.74931581999601, 135.92846479111768, -18.692651199468482, 33.045679309881734, 22.146633725568886, -54.668239355722314,
                           -53.83102333804717, 189.9953989760986, -37.96487763268056, 21.191808796892282, 19.702658746232714, -14.565984021920961, 
                           39.4492486106525, 39.203764053509616, -11.483001309757485, 52.291734008835235, -29.226797706078628]
dframe1


# In[197]:


#Finding the maximun monthly growth rate
maxClm = dframe1['Monthly Growth'].max(skipna=True)  #skip na is used to skip the data null/missing values
print(maxClm)   #Maximun Monthly Growth rate in the month of March 2014


# In[206]:


from matplotlib import style

style.use("ggplot")

x=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
#Growth rate for the year 2014
y1=[0,-68.252263,1132.131409,-49.192257,-16.423401,46.290205,-1.875220,
    -17.783699,193.009345,-61.537770,149.984848,-11.551882]

#Growth rate for the year 2015
y2=[-73.867405,-34.239236,224.030794,-11.700186,-11.883308,-17.703604,16.001880,
    28.273650,75.064601,-51.382495,141.912907,-1.386082]

#Growth rate for the year 2016
y3=[-75.250121,23.925178,125.058929,-25.071288,47.064957,-29.204874,-2.683315,
    -20.749316,135.928465,-18.692651,33.045679,22.146634]

#Growth rate for the year 2017
y4=[-54.668239,-53.831023,189.995399,-37.964878,21.191809,19.702659,-14.565984,
    39.449249,39.203764,-11.483001,52.291734,-29.226798]

#Plotting 2014 Monthly growth rate
plt.plot(x,y1,'k',label='2014', linewidth=2)   #g is the color(green)

#Plotting 2015 Monthly growth rate
plt.plot(x,y2,'r',label='2015',linewidth=2)

#Plotting 2014 Monthly growth rate
plt.plot(x,y3,'b',label='2016', linewidth=2)

# #Plotting 2014 Monthly growth rate
plt.plot(x,y4,"yellow",label='2017', linewidth=2)

#Giving the title for our plot
plt.title('Monthly Growth Rate')

#Giving the names to axes
plt.ylabel('Grwoth Rate')
plt.xlabel('Months')

#Giving description about the plot curves
plt.legend()

#Showing what we have plotted
plt.show()


# In[ ]:


#Describing the findings
#Through this graph it is easy to compare the growth rate of different months of different years.
#March 2014 saw the highest growth rate and over the years the growth rate has decreased.


# # Task 7: Finding out the most and least sold product ID
# a.Create a new dataset including the product ids and total quantities sold for that id
# b.Find out the most sold product from the created dataset
# 

# In[231]:


#Creating a new dataset for product ID's and quantities
product_data=pd.DataFrame(dataf['Product ID'])
product_data['Quantities']=dataf['Quantity']
product_data

#Grouping the same ID's and producing the total number of products for a particular ID
prdgroup=product_data.groupby((product_data['Product ID'])).sum()
prdgroup


# In[236]:


#Most Sold Product
print(prdgroup[prdgroup.Quantities == prdgroup.Quantities.max()])   #Most Sold Product


# # Task 8: Finding out the customer who bought most and least from us in terms of quantity
# a.Create a dataset containing name and quantities bought
# b.Find out the customer name and quantity, who bought maximum in quantity

# In[237]:


#Creating a dataset for customer names and quantities they bought
customer=pd.DataFrame(dataf['Customer Name'])
customer['Quantities']=dataf['Quantity']
customer

#Grouping the customer by their names and the quantities bought by a particular customer
cname=customer.groupby((customer['Customer Name'])).sum()
cname


# In[238]:


#Customer name who bought the maximum products
print(cname[cname.Quantities == cname.Quantities.max()])   


# # Task 9: Finding out the customer who bought most and least from us in terms of value
# a.Create a dataset containing name and sales generated by him
# b.Find out the customer name and sales, who bought maximum in value

# In[16]:


#Creating a dataset for customer names and sales 
customer_sales=pd.DataFrame(dataf['Customer Name'])
customer_sales['Sales']=dataf['Sales']
customer_sales

#Grouping the customers by their names and the sum of sales for the particular customer
csale=customer_sales.groupby((customer_sales['Customer Name'])).sum()
csale


# In[17]:


#Customer name who bought the most in terms of sales
print(csale[csale.Sales == csale.Sales.max()])

#Customer name who bought the least in terms of sales
print(csale[csale.Sales == csale.Sales.min()])


# # Task 10: Finding out the majority and minority customer cities on basis of- 
# a. Number of customers
# b. Sales value
# c. Number of quantity sold

# In[18]:


#Creating a dataset with city name
city=pd.DataFrame(dataf['Customer ID'])
city['City']=dataf['City']
city
# print(city['Customer ID'].value_counts())

#Using count for finding the city with majority and minority in terms of customers
print(city['City'].value_counts())


# In[82]:


#Creating a dataset for city and corresponding sales
city['Sales']=dataf['Sales']
city

#Grouping the cities on the basis of their names and finding totals sales for a particular city

cts=city.groupby([(city['City'])]).sum()
cts1=pd.DataFrame(cts)
cts1


# In[83]:


#City with majority sales
print(cts1[cts1.Sales == cts1.Sales.max()])   #City with majority sales


# In[84]:


print(cts1[cts1.Sales == cts1.Sales.min()])   #City with minority sales


# In[91]:


#Creating the dataset with city and quantities
cityQty=pd.DataFrame(dataf['City'])
cityQty['Quantity']=dataf['Quantity']
cityQty

#Grouping the cities on the basis of their names and finding totals quantities bought by a particular city
cQty=cityQty.groupby([(cityQty['City'])]).sum()
cQty1=pd.DataFrame(cQty)
# print(cQty)
print(cQty1[cQty1.Quantity == cQty1.Quantity.max()])    #City with majority Quantity sold
print(cQty1[cQty1.Quantity == cQty1.Quantity.min()])    #City with minority Quantity sold


# # Task 11: Find out the most and least sold product category from the store
# a.Value based b.Quantity based
#  

# In[99]:


#Creating a dataset with category, sales and quantity
category=pd.DataFrame(dataf['Category'])
category['Value']=dataf['Sales']
category['Quantity']=dataf['Quantity']

#Grouping the categories on the basis of their names and finding totals quantities bought and sales by a particular city
categoryV=category.groupby([(category['Category'])]).sum()
categoryV


# In[101]:


print(categoryV[categoryV.Value == categoryV.Value.max()])    #Category with majority sales
print(categoryV[categoryV.Value == categoryV.Value.min()])    #Category with minority sales
#Technology has the most sold Products on the basis of Sales
#Office Supplies has the most sold Products on the basis of Quantity


# In[112]:


#Category with maximum quantity sold
print(categoryV[categoryV.Quantity == categoryV.Quantity.max()])


# In[113]:


#Category with minimum
quantity soldprint(categoryV[categoryV.Quantity == categoryV.Quantity.min()])


# # Task 12: Find out the most and least sold product sub category from the store
# a.Value based b.Quantity based

# In[107]:


#Creating a dataset for sub-category, sales, and quantities
sub_category=pd.DataFrame(dataf['Sub-Category'])
sub_category['Value']=dataf['Sales']
sub_category['Quantity']=dataf['Quantity']
sub_category

#Grouping the sub-categories on the basis of their names and finding totals quantities bought and sales by a particular city
sub_cat=sub_category.groupby([(sub_category['Sub-Category'])]).sum()
sub_cat


# In[110]:


#Sub-category with most sold products on the basis of sales
print(sub_cat[sub_cat.Value == sub_cat.Value.max()])


# In[111]:


#Sub-category with least sold products on the basis of sales
print(sub_cat[sub_cat.Value == sub_cat.Value.min()])


# In[114]:


#Sub-category with most sold products on the basis of quantities
print(sub_cat[sub_cat.Quantity == sub_cat.Quantity.max()])


# In[115]:


#Sub-category with least sold products on the basis of quantities
print(sub_cat[sub_cat.Quantity == sub_cat.Quantity.min()])

