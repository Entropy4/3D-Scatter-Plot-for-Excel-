#this code displays the 3d scatter plot of 3 variables by extracting the data from the specified excel file


from mpl_toolkits.mplot3d import axes3d
import matplotlib.pyplot as plt
import xlrd
import numpy as np


#location of excel file:
#  'r' is prefixed to the location string to pass it as a 'Raw String'
#       this is because single backward-slashes aren't processed properly by regular strings
#       to avoid replacing every single backward-slash with double backward-slash (\ => \\) we prefix the string with 'r'
loc=(r"D:\User\Desktop\Suja\AIRGLOW-MAY-2020\AG-WORK-SHIVA.xls")

#values to be input before running:
sheet_pos=3                                      #sheet position number
start_row=2                                      #starting row number
end_row=98                                       #ending row number
sa_colindex=11                                   #column no. zero index of Sun Angle
sf_colindex=12                                   #column no. zero index of Solar Flux
ca_colindex=13                                   #column no. zero index of FUV_TEC_MIN
##

i_min=start_row-1
i_max=end_row-1


#importing values from excel file:
wb=xlrd.open_workbook(loc)
sheet=wb.sheet_by_index(sheet_pos-1)
sheet.cell_value(0,0)

fig=plt.figure()
ax=plt.axes(projection='3d')

X=[]
Y=[]
Z=[]

for i in range(i_min,i_max):
    itemx=sheet.cell_value(i,sa_colindex)                 
    itemy=sheet.cell_value(i,sf_colindex)
    itemz=sheet.cell_value(i,ca_colindex)
    X.append(itemx)
    Y.append(itemy)
    Z.append(itemz)


#code to display the scatter plot
ax.scatter3D(X, Y, Z, c=Z,cmap='inferno')

ax.set_xlabel("Sun Angle")
ax.set_ylabel("Solar Flux")
ax.set_zlabel("Constant Airglow");

plt.show()
plt.close('all')