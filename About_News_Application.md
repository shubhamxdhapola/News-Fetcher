'''
This Python script provides a command-line interface for fetching and reading news headlines using the News API. The script allows users to get the top 10 news headlines either by country, by topic, or by both country and topic. Additionally, it offers an option to either read the headlines on the console or listen to them using text-to-speech functionality. Here is a brief explanation of the script:

### Imports and Setup

- `requests`, `json`: These modules are used to handle HTTP requests and to parse JSON data from the News API.
- `win32com.client`: This module is used to access Windows COM functionality, specifically for text-to-speech.

### Initialization

- `speaker`: An instance of `SAPI.spVoice` is created to enable text-to-speech.
- `countries` and `topics`: Lists of valid country codes and news topics are defined to validate user input.
- `API_KEY`: The API key for accessing the News API is stored as a variable.

### Functions

1. **`country_news(country, readOrListen)`**:
   - Fetches top headlines for a specific country.
   - Validates the country code and prints the top 10 headlines.
   - If `readOrListen` is set to 'Listen', it uses text-to-speech to read out the headlines.

2. **`topic_news(topic, readOrListen)`**:
   - Fetches top headlines for a specific topic.
   - Validates the topic and prints the top 10 headlines.
   - If `readOrListen` is set to 'Listen', it uses text-to-speech to read out the headlines.

3. **`country_topic_news(country, topic, readOrListen)`**:
   - Fetches top headlines for a specific country and topic.
   - Validates both the country code and topic, and prints the top 10 headlines.
   - If `readOrListen` is set to 'Listen', it uses text-to-speech to read out the headlines.

### Error Handling

- Each function contains a `try` block to handle network requests. If a request fails, an exception is caught, and an error message is printed.

### User Interface

- The script presents a menu with three options:
  1. Get top headlines by country.
  2. Get top headlines by topic.
  3. Get top headlines by country and topic.
- Based on user input, the script prompts for additional details (country code, topic, read or listen) and calls the appropriate function.
- It ensures inputs are valid and prompts users appropriately.

### Example Usage

1. If a user selects option 1, they are prompted to enter a country code and choose between reading or listening to the news. The script then fetches and displays or reads out the top 10 news headlines for that country.
2. If a user selects option 2, they choose a topic and whether they want to read or listen. The script fetches and displays or reads out the top 10 headlines for that topic.
3. If a user selects option 3, they enter both a country code and a topic, then choose to read or listen. The script fetches and displays or reads out the top 10 headlines for the specified country and topic.

This script leverages the News API and text-to-speech capabilities to provide a flexible and interactive way to stay updated with current news.
'''