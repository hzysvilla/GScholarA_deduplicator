# GScholarA_deduplicator

**GScholarA_deduplicator** is a Python script designed to help users eliminate duplicate academic paper recommendations from their Gmail inbox. It automatically fetches unread academic emails, extracts paper details, and saves the unique results to an Excel file. This tool is perfect for anyone frustrated by seeing the same paper recommendations repeatedly.

## Features

- Fetch unread academic paper emails from Gmail.
- Extract relevant paper details: title, authors, and snippet.
- Detect and remove duplicate entries based on paper title and authors.
- Save unique paper information to an Excel file with improved formatting.

## Prerequisites

1. **Google Gmail API Setup**:
    - This script interacts with Gmail through Google's API. You need to set up the Gmail API and obtain `credentials.json` and `token.json` files.
    
    ### Steps to Get `credentials.json` and `token.json`:
    
    - Go to [Google Developer Console](https://console.developers.google.com/).
    - Create a new project or select an existing one.
    - Enable the **Gmail API**:
        1. In the left menu, navigate to **APIs & Services > Library**.
        2. Search for **Gmail API** and click **Enable**.
    - Set up **OAuth 2.0 credentials**:
        1. In the left menu, navigate to **APIs & Services > Credentials**.
        2. Click **Create Credentials** and choose **OAuth client ID**.
        3. Select **Desktop App** as the application type.
        4. Download the `credentials.json` file.
    - The first time you run the script, it will automatically request access to Gmail and generate the `token.json` file.

2. **Install Required Python Libraries**:

   Before running the script, make sure you have the following Python libraries installed:

   ```bash
   pip install google-auth google-auth-oauthlib google-auth-httplib2 termcolor google-api-python-client openpyxl beautifulsoup4
   ```

## How It Works

1. **Authenticate with Gmail**: 
   The script uses OAuth 2.0 authentication to access Gmail. When you run it for the first time, you will be prompted to sign in and grant access.

2. **Fetch Emails**: 
   It queries your Gmail account for unread messages from `scholaralerts-noreply@google.com` (Google Scholar Alerts). 

3. **Parse Email Content**: 
   The HTML content of the email is extracted and parsed using `BeautifulSoup` to gather details about the academic papers.

4. **De-duplicate Papers**: 
   The script identifies and removes duplicate papers based on the title and authors. It counts the occurrences of each unique paper.

5. **Save to Excel**: 
   The script saves the unique papers, along with their details (title, authors, snippet), into an Excel file. The resulting file is saved in a `history/` folder, named with a timestamp and email count.

## How to Use

1. **Clone the Repository**:
   Clone this repository to your local machine:

   ```bash
   git clone https://github.com/hzysvilla/GScholarA_deduplicator.git
   cd GScholarA_deduplicator
   ```

2. **Configure Gmail API**:
    - Follow the steps above to create and download the `credentials.json` file.
    - Place the `credentials.json` file in the root directory of this project.

3. **Run the Script**:
   Run the script by executing the following command:

   ```bash
   python GScholarA_deduplicator.py
   ```

   The first time you run it, the script will prompt you to authorize Gmail access. After that, it will automatically fetch unread academic paper emails, remove duplicates, and save the results to an Excel file.

4. **Check Results**:
   The unique paper details will be saved in the `history/` folder as an Excel file with a name containing the timestamp and the number of emails processed.

## Configuration

- **SCOPES**: The script uses the `https://www.googleapis.com/auth/gmail.modify` scope to modify Gmail messages (marking them as read after processing).
- **EMAIL_QUERY**: The query string is set to fetch emails from `scholaralerts-noreply@google.com` that are unread. You can modify this if you need to fetch emails from a different source.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
