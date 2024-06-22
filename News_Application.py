# Importing required modules
import requests
import json
import win32com.client

# Instantiate the SAPI.spVoice dispatcher to enable text-to-speech functionality
speaker = win32com.client.Dispatch("SAPI.spVoice")
speaker.Voice = speaker.GetVoices().item(1)
speaker.Rate = 0

# A dictionary mapping country codes to country names
available_countries = {
    
    'United Arab Emirates': 'ae', 'Argentina': 'ar', 'Austria': 'at', 'Australia': 'au', 'Belgium': 'be', 'Bulgaria': 'bg', 'Brazil': 'br', 'Canada': 'ca',
    'Switzerland': 'ch', 'China': 'cn', 'Colombia': 'co', 'Cuba': 'cu', 'Czech Republic': 'cz', 'Germany': 'de', 'Egypt': 'eg', 'France': 'fr', 'United Kingdom': 'gb',
    'Greece': 'gr', 'Hong Kong': 'hk', 'Hungary': 'hu', 'Indonesia': 'id', 'Ireland': 'ie', 'Israel': 'il', 'India': 'in', 'Italy': 'it', 'Japan': 'jp', 'South Korea': 'kr',
    'Lithuania': 'lt', 'Latvia': 'lv', 'Morocco': 'ma', 'Mexico': 'mx', 'Malaysia': 'my', 'Nigeria': 'ng', 'Netherlands': 'nl', 'Norway': 'no', 'New Zealand': 'nz',
    'Philippines': 'ph', 'Poland': 'pl', 'Portugal': 'pt', 'Romania': 'ro', 'Serbia': 'rs', 'Russia': 'ru', 'Saudi Arabia': 'sa', 'Sweden': 'se', 'Singapore': 'sg',
    'Slovenia': 'si', 'Slovakia': 'sk', 'Thailand': 'th', 'Turkey': 'tr', 'Taiwan': 'tw', 'Ukraine': 'ua', 'United States': 'us','Venezuela': 've', 'South Africa': 'za'
}

# List of valid news topics
topics = ['Business', 'Entertainment', 'General', 'Health', 'Science', 'Technology']

# Possible topics for user reference
possibleTopics = ">> Possible topics : Business, Entertainment, General, Health, Science, Technology"

# API Key for News API
API_KEY = 'fcaaa5e72e894f85856c8bd6ceb6c96f'

# Function to get news headlines from a specific country or topic
def fetch_and_display_news(country = "", topic = "", readOrListen = "", userInput=""):
  
  """
  Fetch and display news based on the user's choice.
  
  Args:
  country (str): Country code for news.
  topic (str): News topic/category.
  readOrListen (str): User's preference to read or listen to the news.
  userInput (str): User's choice for fetching news (country, topic, or both).
  
  Returns:
  None
  """
  
  try:
      
      # Construct API request URL based on user input
      if userInput == "1":
            # Send a GET request to the News API to retrieve top headlines for a specific country
            response = requests.get(f"https://newsapi.org/v2/top-headlines?language=en&country={country}&apiKey={API_KEY}")
        
      elif userInput == "2":
            # Send a GET request to the News API to retrieve top headlines for a specific topic
            response = requests.get(f"https://newsapi.org/v2/top-headlines?language=en&category={topic}&apiKey={API_KEY}")

      elif userInput == "3":
            # Send a GET request to the News API to retrieve top headlines for a specific country and topic
            response = requests.get(f"https://newsapi.org/v2/top-headlines?language=en&country={country}&category={topic}&apiKey={API_KEY}")
        
      # Load response data
      news = json.loads(response.text)

      # Iterate over the top 10 headlines
      for headlineNo, news in enumerate (news['articles'], 1):

        # Extract article details
        title = news["title"]
        description = news["description"]
        readMore = news["url"]
        source = news["source"]["name"]
            
        # Print article details
        line = '__'
        print(line*100)
        print(f"\n>> Headline : {headlineNo}")
        print(f">> Title : {title}")
        print(f">> Description: {description}")
        print(f">> Read More : {readMore}")
        print(f">> Source : {source}")
        
        # Terminates the loop after fetching 10 headlines
        if headlineNo == 10:
            break 

        # If readOrListen is set to 'Listen', use text-to-speech to read out the article details
        if readOrListen == 'Listen':
            speaker.Speak(f'Headline : {headlineNo}')
            speaker.Speak(f'Title : {title}')
            speaker.Speak(f'Description : {description}')
            
      print(line*100)

  # Handle exceptions that occur during the request
  except requests.RequestException as e:
          print(f"An error occured : {e}")

def display_main_menu():

    """
    Display the main menu for user input.
    
    Returns:
    None                     
    """

    line = '='
    print(f"\n{line*10} Welcome to the News Fetcher {line*10}\n")
    print(">> Choose an option to get the latest news")
    print(">> Press 1 to get the 'Top 10' headlines from a specific country")
    print(">> Press 2 to get the 'Top 10' headlines on a specific topic")
    print(">> Press 3 to get the 'Top 10' headlines on a specific topic from a specific country")

def taking_user_choice():

    """
    Get the user's choice from the main menu.
    
    Returns:
    str: The user's choice (1, 2, or 3) or None if maximum attempts are exceeded.
    """

    attempts = 0
    while attempts < 3:

        userInput = input("\n>> Enter your choice (1, 2, or 3) : ")

        if userInput in ['1','2','3']:
            return userInput
            
        else: print(">> Invalid input! Please enter (1, 2, or 3)")
        attempts += 1

    print("\n>> Maximum attempts exceeded")
    return None

def get_country():

    """
    Get the country code from the user.
    
    Returns:
    str: The valid country code or None if maximum attempts are exceeded.
    """

    attempts = 0
    while attempts < 3:

        country = input("\n>> Enter country : ").title().strip()

        if country in available_countries: 
            country = country[0:2].lower() # Ensure country code is 2 characters and lowercase
            return country

        else: print(">> Sorry! no news available for this country enter a different country")
        attempts += 1

    print("\n>> Maximum attempts exceeded")
    return None

def get_read_or_listen():

    """
    Get the user's preference to read or listen to the news.
    
    Returns:
    str: 'Read' or 'Listen' or None if maximum attempts are exceeded.
    """

    attempts = 0
    while attempts < 3:

        readOrListen = input("\n>> Would you like to 'Read' or 'Listen' to the news? : ").title().strip()
    
        if readOrListen in ['Read','Listen'] :
            return readOrListen

        else: print(">> Please enter 'Read' or 'Listen")
        attempts += 1

    print("\n>> Maximum attempts exceeded!")
    return None

    
def get_topic():

    """
    Get the news topic from the user.
    
    Returns:
    str: The valid news topic or None if maximum attempts are exceeded.
    """
    attempts = 0
    while attempts < 3:

        print(f'\n{possibleTopics}')
        topic = input(">> Enter topic : ").title().strip()

        if topic in topics:
            return topic
            
        else: print(">> Sorry! no news available around this topic")
        attempts += 1
    
    else:
        print("\n>> Maximum attempts exceeded!")
        return None

def get_news(userInput):

    """
    Get the news based on the user's input choice.
    
    Args:
    user_input (str): The user's choice for fetching news by country, topic, or both.

    Returns:
    None
    """

    if userInput == "1":
        country = get_country()
        
        if country:
            readOrListen = get_read_or_listen()

            if readOrListen:
                fetch_and_display_news(country = country, readOrListen = readOrListen, userInput = userInput)
    
    elif userInput == "2":
        topic = get_topic()

        if topic:
            readOrListen = get_read_or_listen()

            if readOrListen:
                fetch_and_display_news(topic = topic, readOrListen = readOrListen, userInput = userInput)
    
    elif userInput == "3":
        country = get_country()

        if country:
            topic = get_topic()

            if topic:
                readOrListen = get_read_or_listen()

                if readOrListen:
                    fetch_and_display_news(country = country, topic = topic, readOrListen = readOrListen, userInput = userInput)

def start_news():

    """
    Main function to run the news fetching program.
    
    Returns:
    None
    """

    display_main_menu()
    userInput = taking_user_choice()

    if userInput:
        get_news(userInput)


if __name__ == "__main__":
    start_news()
