import pandas as pd
import praw
import time
import warnings
import json
warnings.filterwarnings('ignore')
from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
from random import randint
import smtplib, ssl
from email.mime.text import MIMEText

# You have to install praw, by doing "pip install praw" in command line

class Reddit:
    def __init__(self, config):
        self.reddit = praw.Reddit(
            client_id=config["CLIENT_ID"],
            client_secret=config["CLIENT_SECRET"],
            username=config["USERNAME"],
            password=config["PASSWORD"],
            user_agent="Script by u/ps3far33",
        )

    def post(self, subreddit, title, text, link, flare):
        if not pd.isna(link):
            if not pd.isna(flare):
                template_id=None
                for template in self.reddit.subreddit(subreddit).flair.link_templates:
                    if template['text'].strip().lower() == flare.strip().lower():
                        template_id = template['id']
                        template_text = template['text']
                        break
                if template_id:
                    submission = self.reddit.subreddit(subreddit).submit(title=title, url=link, flair_id=template_id, flair_text=template_text)
                else:
                    print('Flare is not corresponding anyone on the list')
                    return
            else:
                submission = self.reddit.subreddit(subreddit).submit(title=title, url=link)
        else:
            if not pd.isna(flare):
                template_id=None
                for template in self.reddit.subreddit(subreddit).flair.link_templates:
                    if template['text'].strip().lower() == flare.strip().lower():
                        template_id = template['id']
                        template_text = template['text']
                        break
                if template_id:
                    submission = self.reddit.subreddit(subreddit).submit(title=title, selftext=text, flair_id=template_id, flair_text=template_text)
                else:
                    print('Flare is not corresponding anyone on the list')
                    return
            else:
                submission = self.reddit.subreddit(subreddit).submit(title=title, selftext=text)
        
            
        return submission

    
def proceed_input(row, config):
    reddit = Reddit(config)
    subreddit = row.Subreddit
    title = row.Title
    text = row.Text
    link = row.Link
    flare = row.Flare
    sub = reddit.post(subreddit, title=title, text=text, link=link, flare=flare)
    print(f"Posting to {subreddit}")
def main():
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    directory_excel = askopenfilename() # show an "Open" dialog box and return the path to the selected file
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    directory_config = askopenfilename() # show an "Open" dialog box and return the path to the selected file
    with open(directory_config) as f:
        data = f.read()
    config = json.loads(data)
    hour_delay = config['Hour Delay']*60*60
    print(f"You have to wait {hour_delay} in hours as the set hour delay")
    time.sleep(hour_delay)
    print('--------- Posting starts Now --------------')
    df_input = pd.read_excel(directory_excel)
    #First Posting
    not_posted = []
    indices = []
    for i in range(df_input.shape[0]):
        try:
            proceed_input(df_input.iloc[i], config)
            time_to_sleep = randint(int(config['time_down_limit'])*60, int(config['time_upper_limit']*60))
            time.sleep(time_to_sleep)
        except:
            not_posted.append(df_input.iloc[i])
            indices.append(i)
            print(f'The sumbission {i+1} is skipped, not posted (error)')
            pass
    #Second Posting
    '''print('------------ Retry failed posts ---------------')
    not_posted_again = 0
    for i in range(len(not_posted)):
        try:
            proceed_input(not_posted[i], config)
            time_to_sleep = randint(int(config['time_down_limit'])*60, int(config['time_upper_limit']*60))
            time.sleep(time_to_sleep)
        except:
            print(f'The sumbission {indices[i]} is skipped for the second time, not posted (error)')
            not_posted_again+=1
            pass'''
    
    df_new = df_input.iloc[indices][['Subreddit','Title','Text','Link','Flare']]
    df_new.to_excel('post_with_issues.xlsx')
    sender = config['email']
    receivers = [config['email']]
    body_of_email = f"The queue is finished, the number of failed post is {len(not_posted)}"

    msg = MIMEText(body_of_email, "html")
    msg["Subject"] = "Reddit Bot: Queue finished"
    msg["From"] = sender
    msg["To"] = ','.join(receivers)

    s = smtplib.SMTP_SSL(host = 'smtp.gmail.com', port = 465)
    s.login(user = sender, password = config['password_mail'])
    s.sendmail(sender, receivers, msg.as_string())
    s.quit()
    if len(not_posted) == df_input.shape[0]:
        print('Account might be banned, all posts failed')
        body_of_email = f"All posts failed, the account might be banned"
        msg = MIMEText(body_of_email, "html")
        msg["Subject"] = "Reddit Bot: Account might be banned"
        msg["From"] = sender
        msg["To"] = ','.join(receivers)

        s = smtplib.SMTP_SSL(host = 'smtp.gmail.com', port = 465)
        s.login(user = sender, password = config['password_mail'])
        s.sendmail(sender, receivers, msg.as_string())
        s.quit()
if __name__ == "__main__":
    main()
