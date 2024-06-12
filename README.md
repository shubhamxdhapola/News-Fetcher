This code is a Python script that fetches and displays the latest news headlines using the News API. It offers three main functionalities:

1. **Fetch top 10 headlines from a specific country**.
2. **Fetch top 10 headlines on a specific topic**.
3. **Fetch top 10 headlines on a specific topic from a specific country**.

The script interacts with the user via a command-line interface, prompting them to choose an option, enter a country code (if applicable), and select whether they want to read the news or listen to it using text-to-speech.

Key components of the script:
- **News API Integration**: Makes HTTP requests to the News API to retrieve news headlines.
- **User Interaction**: Provides prompts to gather user preferences.
- **Text-to-Speech**: Utilizes the `win32com.client` library to read news headlines aloud.
- **Error Handling**: Manages potential issues during the API request process.

The script ensures user input is valid and offers three attempts for each input field before terminating with a message if the user fails to provide valid input.
