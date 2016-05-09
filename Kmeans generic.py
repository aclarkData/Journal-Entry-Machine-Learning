# -*- coding: utf-8 -*-
"""
Created on Thu Nov 05 13:30:49 2015

@author: aclark
"""
#http://www.gadatascience.com/modeling/kmeans.html

import pyodbc
import pandas as pd
import pandas.io.sql as psql
from pandas import options
options.io.excel.xlsx.writer = 'xlsxwriter'

# Example SQL query is from an Microsoft Dynamics AX 2009 ERP system
cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=Put you server name here;DATABASE=put your database here;Trusted_Connection=yes;')
cursor = cnxn.cursor()
sql = '''
SELECT cast(ledgertrans.accountnum AS INT) AS AccountNum, cast(ledgertrans.transdate AS INT) AS Date, 
amountcur, txt, ledgertrans.Posting, LedgerPostingJournalID, Voucher, Transtype, crediting, createdby, Name FROM ledgertrans 
LEFT JOIN LEDGERTABLE ON LEDGERTABLE.ACCOUNTNUM = LEDGERTRANS.ACCOUNTNUM AND LEDGERTABLE.DATAAREAID = LEDGERTRANS.DATAAREAID 
LEFT JOIN UserInfo ON ID = createdby
WHERE transdate > (GetDate() -31) and txt NOT LIKE '%Reverse%' 
and txt NOT LIKE '%Void%' AND LEDGERTABLE.DATAAREAID = 'your company here' ORDER BY transdate DESC;
'''

je_data = psql.read_sql(sql, cnxn)
cnxn.close()

from sklearn.feature_extraction import DictVectorizer

# Turn categorical features into 1-hot encoded features
categorical_features = je_data[['createdby', 'AccountNum', 'Posting']]
dv = DictVectorizer()
cat_matrix = dv.fit_transform(categorical_features.T.to_dict().values())

# Collect the other numerical features
from scipy.sparse import hstack
other_features = je_data[['amountcur','crediting','Date']]
data_matrix = hstack([cat_matrix, other_features])
data_matrix

from sklearn.preprocessing import scale
data_matrix = scale(data_matrix.todense())

from sklearn.cluster import KMeans

# define how many clusters and how many iterations. 
clustering_model = KMeans(n_clusters = 25, n_init= 10)
clustering_model.fit(data_matrix)

clusters = clustering_model.predict(data_matrix)

results = pd.DataFrame({ 'cluster' : clusters, 'amountcur' : je_data['amountcur'], 'Date' : je_data['Date'], 'crediting' : je_data['crediting'], 
'Name' : je_data['Name'],'createdby' : je_data['createdby'],'AccountNum' : je_data['AccountNum'],  'txt' : je_data['txt'], 
'LedgerPostingJournalID' : je_data['LedgerPostingJournalID'],'Posting' : je_data['Posting'], 'Voucher' : je_data['Voucher']})
cluster_counts = results.groupby('cluster')['amountcur'].value_counts()     
results.hist()
     
#Group results by cluster
bycluster = results.groupby('cluster')

pd.set_option('display.precision',3)
#http://stackoverflow.com/questions/22105452/pandas-what-is-the-equivalent-of-sql-group-by-having

#extract data per cluster where the the cluster has fewer than 20 values
kclusters=bycluster.filter(lambda x: len(x) < 20)

#export results
kclusters.to_excel(r'put your path here\kmeans.xlsx', sheet_name='Sheet1')  
 
 
#email:
#http://stackoverflow.com/questions/882712/sending-html-email-using-python

import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# me == my email address
# you == recipient's email address
me = "your email"
you = "who you are emailing"

# Create message container - the correct MIME type is multipart/alternative.
msg = MIMEMultipart('alternative')
msg['Subject'] = "Kmeans"
msg['From'] = me
msg['To'] = you

# Create the body of the message (a plain-text and an HTML version).
text = "Hi!\nHow are you?\nFollow the path below to see the latest clustering results."
html = """\
<html>
  <head></head>
  <body>
    <p>Hi!<br>
       How are you?<br>
       Follow the path below to see the latest clustering results.<br>
       Click on "kmeans" for the results. 
    </p>
  </body>
</html>
"""

# Record the MIME types of both parts - text/plain and text/html.
part1 = MIMEText(text, 'plain')
part2 = MIMEText(html, 'html')

# Attach parts into message container.
# According to RFC 2046, the last part of a multipart message, in this case
# the HTML message, is best and preferred.
msg.attach(part1)
msg.attach(part2)
# Send the message via local SMTP server.
mail = smtplib.SMTP('smtp.gmail.com', 587)

mail.ehlo()

mail.starttls()

mail.login('your username', 'your password')
mail.sendmail(me, you, msg.as_string())
mail.quit()
 
#schedule by making a batch file and then scheduling through windows scheduler. See example batch file below:
# start 'your python path here' 'path to file here'\AstecKmeans.py 
