import win32com.client
from nltk.tokenize import RegexpTokenizer
from nltk.corpus import stopwords
from collections import Counter
from MessageFinder import MessageFinder
import pandas as pd
import xlsxwriter
import string
import numpy as np
import datetime

stop_words = set(stopwords.words("english"))
stop_words.update(["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "one", "two", "customer", "item", "would",
                   "order", "orders", "also", "know", "amazon", "stated", "could", "like", "need", "issue",
                   "cx", "items"])

# Calls MessageFinder function from the MessageFinder script, returns a message object for specified folder
messages = MessageFinder('carriaus@amazon.com', 'Inbox', 'AmazonFresh Inventory Feedback')

# Get the latest message
message = messages.GetNext()
body_content = message.body

# Tokenize the E-mail to find the final location of "Feedback:" to find the true body of the e-mail
tokenizer = RegexpTokenizer('\s+',gaps=True)
body_list_full = tokenizer.tokenize(body_content)

count = 0

all_words = []
all_dates = []
all_body = []

for message in messages:
    count = count + 1
    # Get the latest message
    message = messages.GetNext()
    
    try:
        body_content = message.body
    except AttributeError:
        break
    
    # Tokenize the E-mail to find the final location of "Feedback:" to find the true body of the e-mail
    tokenizer = RegexpTokenizer('\s+',gaps=True)
    body_list_full = tokenizer.tokenize(body_content)
    
    feedback_loc = 0
    counter = 1
    for word in body_list_full:
        if word == "Feedback:":
            feedback_loc = counter
            counter = counter + 1
        else:
            counter = counter + 1

    body_list = body_list_full[feedback_loc:]

    body_list = [''.join(c for c in s if c not in string.punctuation) for s in body_list]
    body_list = [s for s in body_list if s]
    DT = message.SentOn

    week_num = datetime.date(DT.year, DT.month, DT.day).isocalendar()[1]

    filtered_words = []

    for word in body_list:
        if word.lower() not in stop_words:
            filtered_words.append(word.lower())
            
    for word in filtered_words:
        all_words.append(word)
    
    date_list = [str(week_num)] * len(filtered_words)
    
    body_text = []
    body_text.append(body_content)

    body_text_list = body_text * len(filtered_words)
    
    for date in date_list:
        all_dates.append(date)

    for text in body_text_list:
        all_body.append(text)
        
##    if count == 10:
##        break

# Create a dictionary with two keys, one for each list
all_dict = {"Week": all_dates, "Word": all_words, "Body": all_body}

# Add this dictionary to dataframe with Key's as column headers
df_body = pd.DataFrame(all_dict, columns=["Week", "Word", "Body"])

df = df_body.drop(["Body"], axis=1)

# Get count of Date-Word combinations and put in new dataframe
# Note: This will output the data in Multi-index format so will need to convert to dataframe
df_count = pd.DataFrame(df.groupby(["Week", "Word"])["Word"].count())

# Ref: https://stackoverflow.com/questions/20110170/turn-pandas-multi-index-into-column
# Temporarily renaming 'Word' to 'W', so when we reset the index, we won't have duplicate columns
df_count.index = df_count.index.set_names(["Week", "W"])
df_count.reset_index(inplace=True)

# Renaming columns to proper names
df_count.columns = ["Week", "Word", "Count"]

# Add Mean and Standard Deviation to each row
df_mean = pd.DataFrame(df_count.groupby(["Word"])["Count"].mean())
df_mean.reset_index(inplace=True)
df_mean.columns = ["Word", "Mean"]
df_std = pd.DataFrame(df_count.groupby(["Word"])["Count"].std())
df_std.reset_index(inplace=True)
df_std.columns = ["Word", "Stnd_Dev"]

df_final = pd.merge(df_count, df_mean, on="Word")
df_final = pd.merge(df_final, df_std, on="Word")

# drop row if all there are any NaN values in any row
df_final = df_final.dropna()

# for each row calculate how the count of the word on the day varies from the average
# and how many standard deviations this is.
# if greater than one standard deviation, print output (means data is an outlier in >68%
# Ref: https://en.wikipedia.org/wiki/68%E2%80%9395%E2%80%9399.7_rule

df_final["DeltaToMean"] = df_final["Count"].sub(df_final["Mean"], axis=0)

df_final["Stnd_Dev_Count"] = np.where(df_final["DeltaToMean"] == 0, df_final["DeltaToMean"], df_final["DeltaToMean"]/df_final["Stnd_Dev"])

df_final = df_final.drop(['DeltaToMean'], axis=1)

df_final = df_final[df_final["Stnd_Dev_Count"] > 1.5]


print(df_final)

df_output = pd.merge(df_body, df_final, on=["Word", "Week"])

# Create a Pandas Excel writer using XlsxWriter as the engine.
# Ref: http://xlsxwriter.readthedocs.io/working_with_pandas.html
writer = pd.ExcelWriter('data_master.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df_output.to_excel(writer, sheet_name='data', index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()

