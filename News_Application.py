# Importing required modules
import requests
import json
import win32com.client

# Instantiate the SAPI.spVoice dispatcher to enable text-to-speech functionality
speaker = win32com.client.Dispatch("SAPI.spVoice")

# List of valid country codes
countries = ['ae','ar','at','au','be','bg','br','ca','ch','cn','co','cu','cz','de','eg','fr','gb','gr','hk','hu','id','ie','il','in','it',
             'jp','kr','lt','lv','ma','mx','my','ng','nl','no','nz','ph','pl','pt','ro','rs','ru','sa','se','sg','si','sk','th','tr','tw',
             'ua','us','ve','za']

# List of valid news topics
topics = ['Business', 'Entertainment', 'General', 'Health', 'Science', 'Technology']

# API Key for News API
API_KEY = 'fcaaa5e72e894f85856c8bd6ceb6c96f'

# Function to get news headlines from a specific country
def country_news(country,readOrListen):

    try:
        # Send a GET request to the News API to retrieve top headlines for a specific country
        response = requests.get(f"https://newsapi.org/v2/top-headlines?language=en&country={country}&apiKey={API_KEY}")
        countryNews = json.loads(response.text)

        # Check if the country code is valid
        if country in countries:

            # Iterate over the top 10 headlines
            for i in range(10):

                # Extract article details
                print(headLine:=f"\nHeadline : {i+1}")
                print(title:=f'Title : {countryNews["articles"][i]["title"]}')
                print(description:=f'Description : {countryNews["articles"][i]["description"]}')
                print(f'Read More : {countryNews["articles"][i]["url"]}')
                print(f'Source : {countryNews["articles"][i]["source"]["name"]}')
                print("")

                # If readOrListen is set to 'Listen', use text-to-speech to read out the article details
                if readOrListen == 'Listen':
                    speaker.Speak(headLine)
                    speaker.Speak(title)
                    speaker.Speak(description)

        # Print an error message if the country code is invalid
        else: print(">> No news related to this country")

    # Handle exceptions that occur during the request
    except requests.RequestException as e:
          print(f"An error occured : {e}")


# Function to get news headlines on a specific topic       
def topic_news(topic,readOrListen):

    try:

        # Send a GET request to the News API to retrieve top headlines for a specific topic
        response = requests.get(f"https://newsapi.org/v2/top-headlines?language=en&category={topic}&apiKey={API_KEY}")
        topicNews = json.loads(response.text)

        # Check if the topic is valid
        if topic in topics:

            # Iterate over the top 10 headlines
            for i in range(10):

                # Extract article details
                print(headLine:=f"\nHeadline : {i+1}")
                print(title:=f'Title : {topicNews["articles"][i]["title"]}')
                print(description:=f'Description : {topicNews["articles"][i]["description"]}')
                print(f'Read More : {topicNews["articles"][i]["url"]}')
                print(f'Source : {topicNews["articles"][i]["source"]["name"]}')

                # If readOrListen is set to 'Listen', use text-to-speech to read out the article details
                if readOrListen == 'Listen':
                    speaker.Speak(headLine)
                    speaker.Speak(title)
                    speaker.Speak(description)

        # Print an error message if the topic is invalid
        else: print(">> No news related to this topic")

    # Handle exceptions that occur during the request
    except requests.RequestException as e:
          print(f"An error occured : {e}")


# Function to get news headlines on a specific topic from a specific country
def country_topic_news(country,topic,readOrListen):

    try:

        # Send a GET request to the News API to retrieve top headlines for a specific country and topic
        response = requests.get(f"https://newsapi.org/v2/top-headlines?language=en&country={country}&category={topic}&apiKey={API_KEY}")
        countryTopicNews = json.loads(response.text)

        # Check if the topic is valid
        if topic in topics:

            # Iterate over the top 10 headlines
            for i in range(10):

                # Extract article details
                print(headLine:=f"\nHeadline : {i+1}")
                print(title:=f'Title : {countryTopicNews["articles"][i]["title"]}')
                print(description:=f'Description : {countryTopicNews["articles"][i]["description"]}')
                print(f'Read More : {countryTopicNews["articles"][i]["url"]}')
                print(f'Source : {countryTopicNews["articles"][i]["source"]["name"]}')
                print("")

                # If readOrListen is set to 'Listen', use text-to-speech to read out the article details
                if readOrListen == 'Listen':
                    speaker.Speak(headLine)
                    speaker.Speak(title)
                    speaker.Speak(description)

        # Print an error message if the topic is invalid
        else: print(">> No news related to this country or topic")

    # Handle exceptions that occur during the request
    except requests.RequestException as e:
          print(f"An error occured : {e}")
    

# Main menu for user input
print("\n>> Press 1 to get the 'Top 10' headlines from a specific country")
print(">> Press 2 to get the 'Top 10' headlines on a specific topic")
print(">> Press 3 to get the 'Top 10' headlines on a specific topic from a specific country")
userInput = input("\n>> Enter Here : ")

# Possible topics for user reference
possibleTopics = ">> Possible topics : Business, Entertainment, General, Health, Science, Technology"
    
# Option 1: Get news by country    
if userInput == "1":

    country = input(">> Enter country : ")
    country = country[0:2].lower() # Ensure country code is 2 characters and lowercase
    readOrListen = input(">> Do you want to read or listen ? :").title()

    if (readOrListen =='Read') or (readOrListen =='Listen'):   
        country_news(country,readOrListen)

    else: print(">> Enter read or listen")


# Option 2: Get news by topic
elif userInput == "2":

    print(possibleTopics)
    topic = input(">> Enter topic : ")

    readOrListen = input(">> Do you want to read or listen ? :").title()
    if (readOrListen =='Read') or (readOrListen =='Listen'):   
            topic_news(topic.title(),readOrListen)

    else: print(">> Enter read or listen")

# Option 3: Get news by country and topic
elif userInput == "3":

    country = input(">> Enter country : ")
    country = country[0:2].lower() # Ensure country code is 2 characters and lowercase

    if country in countries:

        print(possibleTopics)
        topic = input(">> Enter topic : ")

        readOrListen = input(">> Do you want to read or listen ? :").title()
        if (readOrListen =='Read') or (readOrListen =='Listen'):   
            country_topic_news(country,topic.title(),readOrListen)

        else: print(">> Enter read or listen!")

    else: print(">> No news for this country")

# If user inputs a unavailable option
else: print(">> Invalid input!")
         
