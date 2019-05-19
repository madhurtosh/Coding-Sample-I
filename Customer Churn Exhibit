import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import confusion_matrix
from sklearn.metrics import roc_curve
from IPython.display import display, HTML
import sys
import xlwings as xw
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from timeit import Timer


def churn(country):
    df = pd.read_excel("Churn.xls")
    
    df = df[df['Country'] == country]
    df = df.drop(["Phone", "Area Code", "Country"], axis=1)
    features = df.drop(["Churn"], axis=1).columns
    
    df_train, df_test = train_test_split(df, test_size=0.25)

    
    clf = RandomForestClassifier(n_estimators=30)
    clf.fit(df_train[features], df_train["Churn"])
    
    predictions = clf.predict(df_test[features])
    probs = clf.predict_proba(df_test[features])
    #return probs_test, len(probs_test)
    #display(predictions)
    
    score = clf.score(df_test[features], df_test["Churn"])
    
    #print("Accuracy: ", score)
    df_test = df_test[df_test["Churn"] == 1]["Customer ID"]
    df_test.index = np.arange(0, len(df_test))
    return df_test

def rfm(country):
    data = pd.read_csv("Retail Data for Analysis.csv",encoding='cp1252')
    data_cust = data[['CustomerID','Country']].drop_duplicates()
    data_cust = data_cust.groupby(['Country'])['CustomerID'].count().reset_index().sort_values(by='CustomerID',ascending=False)
    
    data = data[data['Country'] == country]
    data['InvoiceDate']=pd.to_datetime(data['InvoiceDate'])
    data_monetory = data.groupby(['CustomerID'])['UnitPrice'].sum().reset_index()
    data_monetory['montry_rank']=data_monetory['UnitPrice'].rank(ascending=False)
    data_freq = data[['CustomerID','InvoiceNo']].drop_duplicates()
    data_freq = data_freq.groupby(['CustomerID'])['InvoiceNo'].count().reset_index()
    data_freq['freq_rank']=data_freq['InvoiceNo'].rank(ascending=False)
    
    data_recency = data.groupby(['CustomerID'])['InvoiceDate'].max().reset_index()
    data_recency['recency_rank'] = data_recency['InvoiceDate'].rank(ascending=False)
    
    data_all_rank = pd.merge(data_monetory,data_freq,on='CustomerID')
    data_all_rank = pd.merge(data_all_rank,data_recency,on='CustomerID')
    data_all_rank['final_rank'] = data_all_rank['montry_rank'] + data_all_rank['freq_rank'] + data_all_rank['recency_rank']
    data_all_rank = data_all_rank.sort_values(by='final_rank').iloc[:30]
    
    data_recency_not = data_recency
    data_recency_not['recency_rank'] = data_recency['InvoiceDate'].rank(ascending=True)
    data_recency_not = data_recency_not[data_recency_not['recency_rank']<=0.3*data_recency_not.shape[0]]
    data_all_rank_not_recent=pd.merge(data_recency_not,data_freq,on='CustomerID')
    data_all_rank_not_recent=pd.merge(data_all_rank_not_recent,data_monetory,on='CustomerID')
    data_all_rank_not_recent['final_rank'] = data_all_rank_not_recent['montry_rank'] + data_all_rank_not_recent['freq_rank']
    data_all_rank_not_recent=data_all_rank_not_recent.sort_values(by='final_rank').iloc[:30]
    data_all_rank_not_recent.index = np.arange(0, len(data_all_rank_not_recent))
    data_all_rank.index = np.arange(0, len(data_all_rank))
    return data_all_rank_not_recent["CustomerID"], data_all_rank["CustomerID"]

def create_excel(list_df, names_df, country):
    wb = xw.Book()
    #wb.activate(steal_focus=True)

    for index in range(len(list_df)):
        sht = wb.sheets.add(names_df[index])
        #sht = wb.sheets[names_df[index]]
        sht.range('A1').value = list_df[index]
        
    wb.save(r"C:\TrendsMarketplace\Output\CustomerData_"+ country +".xlsx")
    wb.close()
    return r"C:\TrendsMarketplace\Output\CustomerData_"+ country +".xlsx"

def send_email(excel_file, country):
    creds = pd.read_csv("Email Credentials.json", header=None, squeeze=True)
    recpts = pd.read_csv("Email Recipients.csv", squeeze=True)
    
    
    #recipients = ','.join(recpts)
    recipients = list(recpts)
    # creates SMTP session
    s = smtplib.SMTP('smtp.gmail.com', 587)

    # start TLS for security
    s.starttls()

    # Authentication
    s.login(creds[0], creds[1])

    # Create email body
    SUBJECT = "Monthly Customer Report: "+country

    msg = MIMEMultipart()
    msg['Subject'] = SUBJECT
    msg['From'] = creds[0]
    msg['To'] = ', '.join(recipients)
    body = MIMEText("Hi,\n\nPlease refer to the attached file for top performers, non-recent performers and possible churns.\n\nRegards,\nAnalytics Team")
    msg.attach(body)
    
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(excel_file, "rb").read())
    encoders.encode_base64(part)

    part.add_header('Content-Disposition', 'attachment; filename="'+excel_file+'"')

    msg.attach(part)

    # sending the mail
    s.sendmail(creds[0], recipients, msg.as_string())

# Main
country = sys.argv[1]
df_churn = churn(country)
df_not_recent, df_recent = rfm(country)
excel_file = create_excel([df_churn, df_not_recent, df_recent], ["Possible Churns", "Non Recent Performers", "Top Performers"], country)
send_email(excel_file, country)


# Time the function
def convert_timing(country):
    timer1 = Timer(lambda: test_time(country)).timeit(number=1)
    return (timer1)
