from bs4 import BeautifulSoup
import xlsxwriter
import urllib2
import _tkinter

out=[]
q=[]
outlink=[]
inlink=[]
lin=[]
lin1=[]
z=[]
final=[]
temp=[]

url = raw_input("Enter a website to extract the URL's from: ")
r  = urllib2.urlopen("http://" +url)
data = r.read()
soup = BeautifulSoup(data)
for link in soup.find_all('a'):
    z.append(link.get('href'))
    lin.append(link.get('href'))
    lin1.append(link.get('href'))
for i in range(len(z)):
    if(z[i]!=None):
        z[i]=z[i].encode('utf_8')
        lin[i]=lin[i].encode('utf_8')
q.append(z)
#print q
print("Outlink of " + url + " is :-")
print len(z)
i=0
for i in range(z.count(None)): 
    z.remove(None)
    lin.remove(None)
if((z.count(''))>=1):
    z.remove('')
    lin.remove('')

j=0
while(j < len(z)):
    g=[]
    u=z[j]
    while(u==None):
        j=j+1
        u=z[j]
        outlink.append(0)
    try:
        if(u.index('h')!=0):
            r  = urllib2.urlopen("http://" +url+"/"+z[j])
        else:
            r  = urllib2.urlopen(z[j])
        data = r.read()
        soup = BeautifulSoup(data)
        for link in soup.find_all('a'):
            g.append(link.get('href'))
            lin1.append(link.get('href'))
        for i in range(len(g)):
            if(g[i]!=None):
                g[i]=g[i].encode('utf_8')
        q.append(g)
        #print q
    except (ValueError,urllib2.HTTPError,urllib2.URLError,UnicodeDecodeError):
        print ""
    print("Outlink of " + u + " is :-")
    print len(g)
    outlink.append(len(g))
    j=j+1
#print outlink
#print lin

for i in range(len(lin1)):
    if(lin1[i]!=None):
        lin1[i]=lin1[i].encode('utf_8')

for i in range(len(lin)):
    inlink.append(lin1.count(lin[i]))
#print inlink
l=len(inlink)

workbook=xlsxwriter.Workbook('Connection Matrix.xlsx')
worksheet=workbook.add_worksheet()
worksheet.set_column('A:A',75)
worksheet.set_column('B:C',40)
cellformat=workbook.add_format()
cellformat.set_bold()
cellformat.set_font_size(20)
worksheet.write(0,0,'Links In The Page',cellformat)
worksheet.write(0,1,'No. Of Outlinks',cellformat)
worksheet.write(0,2,'No. Of Inlinks',cellformat)
for i in range(len(lin)):
    worksheet.write(i+1,0,lin[i])
    worksheet.write(i+1,1,outlink[i])
    worksheet.write(i+1,2,inlink[i])
    final.append([lin[i],outlink[i],inlink[i]])
    print final[i]

for i in range(l):
    temp.append(inlink[i])
temp.sort()
temp.reverse()

for i in range (10):
    print "LINK:",lin[inlink.index(temp[i])]
    print "OUTLINK:",outlink[inlink.index(temp[i])]
    print "INLINK:",temp[i]

worksheet1=workbook.add_worksheet()
worksheet1.set_column('A:A',75)
worksheet1.set_column('B:C',40)
cellformat=workbook.add_format()
cellformat.set_bold()
cellformat.set_font_size(20)
worksheet1.write(0,0,'Links In The Page',cellformat)
worksheet1.write(0,1,'No. Of Outlinks',cellformat)
worksheet1.write(0,2,'No. Of Inlinks',cellformat)
for i in range(10):
    worksheet1.write(i+1,0,lin[inlink.index(temp[i])])
    worksheet1.write(i+1,1,outlink[inlink.index(temp[i])])
    worksheet1.write(i+1,2,temp[i])

chart1 = workbook.add_chart({'type': 'column'})
chart1.add_series({
    'name':       '=Sheet2!$B$1',
    'categories': '=Sheet2!$A$2:$A$11',
    'values':     '=Sheet2!$B$2:$B$11',
})
chart1.add_series({
    'name':       ['Sheet2', 0, 2],
    'categories': ['Sheet2', 1, 0, 10, 0],
    'values':     ['Sheet2', 1, 2, 10, 2],
})
chart1.set_title({'name_font': {'bold': True}})
chart1.set_title ({'name': 'Results of sample analysis'})
chart1.set_x_axis({'name_font': {'bold': True, 'italic': True}})
chart1.set_x_axis({'name': 'Links'})
chart1.set_y_axis({'name_font': {'bold': True, 'italic': True}})
chart1.set_y_axis({'name': 'inlimks/outlinks'})
chart1.set_style(10)
worksheet1.insert_chart('D2', chart1, {'x_scale':2, 'y_scale': 2})

workbook.close()
