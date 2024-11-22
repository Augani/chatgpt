import openai
import PyPDF2
import pandas as pd
import json
import os
import dotenv

dotenv.load_dotenv()

# Set your OpenAI API key
openai.api_key = os.getenv('OPENAI_API_KEY')

# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_path):
    text = ''
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text += page.extract_text()
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}")
    return text

# Function to extract data from an Excel file
def extract_data_from_excel(excel_path):
    try:
        df = pd.read_excel(excel_path)
        return df.to_string()
    except Exception as e:
        print(f"Error reading Excel file {excel_path}: {e}")
        return ''

# Function to read news articles from a JSON file
def extract_articles_from_json(json_path):
    compiled_text = ''
    try:
        with open(json_path, 'r', encoding='utf-8') as file:
            articles = json.load(file)
        for article in articles:
            title = article.get('title', 'No Title')
            content = article.get('content', '')
            compiled_text += f'\n\nTitle: {title}\nContent:\n{content}'
    except Exception as e:
        print(f"Error reading JSON file {json_path}: {e}")
    return compiled_text

# Function to extract data from a CSV file
def extract_data_from_csv(csv_path):
    try:
        df = pd.read_csv(csv_path)
        return df.to_string()
    except Exception as e:
        print(f"Error reading CSV file {csv_path}: {e}")
        return ''


# Collect data from multiple sources
def collect_data(news_json_files, pdf_paths, excel_paths):
    compiled_text = ''

    # Extract text from news articles JSON files
    for json_file in news_json_files:
        print(f'Extracting articles from JSON file {json_file}')
        articles_text = extract_articles_from_json(json_file)
        compiled_text += f'\n\nArticles from {json_file}:\n{articles_text}'

    # Extract text from PDFs
    for pdf in pdf_paths:
        print(f'Extracting text from PDF {pdf}')
        pdf_text = extract_text_from_pdf(pdf)
        compiled_text += f'\n\nContent from PDF {pdf}:\n{pdf_text}'

    # Extract data from Excel files
    for excel in excel_paths:
        print(f'Extracting data from Excel file {excel}')
        excel_text = extract_data_from_excel(excel)
        compiled_text += f'\n\nData from Excel {excel}:\n{excel_text}'

    return compiled_text

# Function to generate an analysis article using OpenAI's API
def generate_article(compiled_text, style_instructions, sample_outline=None):
    # Construct the prompt
    prompt = (
        "Based on the following data from news articles, PDFs, and Excel sheets, "
        "write a comprehensive analysis article.\n"
        "The article should follow these style and format guidelines:\n"
        f"{style_instructions}\n"
    )

    if sample_outline:
        prompt += f"\nHere is an outline/example to follow:\n{sample_outline}\n"

    prompt += f"\nData:\n{compiled_text}\n\nPlease provide a detailed analysis."

    prompt += f"\n\nPlease focus on the Nigerian Stock Exchange (NSE) market."


    # Increased max tokens for the prompt
    max_prompt_tokens = 7000  # Increased from 3500

    prompt = prompt[-max_prompt_tokens:]

    # Split the content into multiple chunks if needed
    chunks = [compiled_text[i:i + max_prompt_tokens] 
             for i in range(0, len(compiled_text), max_prompt_tokens)]
    
    full_article = ""
    
    for i, chunk in enumerate(chunks):
        current_prompt = prompt.replace(compiled_text, chunk)
        if i > 0:
            current_prompt = f"Continue the analysis with the following additional data:\n{chunk}"
        
        response = openai.chat.completions.create(
            model="gpt-3.5-turbo-16k",  # Using 16k model for longer context
            messages=[
                {"role": "system", "content": "You are an expert analyst and writer. Create detailed, comprehensive analysis incorporating as much relevant information as possible from the source materials."},
                {"role": "user", "content": current_prompt}
            ],
            max_tokens=4000,  # Increased from 1500
            temperature=0.7
        )
        
        full_article += response.choices[0].message.content + "\n\n"
    
    return full_article

# Main function
def main():
    # Define your data source directories
    news_articles_dir = './input/articles'
    pdf_dir = './input/research_files'
    excel_dir = './input/market_data'

    # Verify that directories exist
    for directory in [news_articles_dir, pdf_dir, excel_dir]:
        if not os.path.exists(directory):
            print(f"Directory {directory} does not exist.")
            return

    # Get list of all JSON files in the news articles directory
    news_json_files = [
        os.path.join(news_articles_dir, f)
        for f in os.listdir(news_articles_dir)
        if f.endswith('.json')
    ]

    # Get list of all PDF files in the research directory
    pdf_paths = [
        os.path.join(pdf_dir, f)
        for f in os.listdir(pdf_dir)
        if f.endswith('.pdf')
    ]

    # Get list of all Excel files in the excel directory
    excel_paths = [
        os.path.join(excel_dir, f)
        for f in os.listdir(excel_dir)
        if f.endswith('.xls') or f.endswith('.xlsx')
    ]

    # Check if there are files to process
    if not news_json_files and not pdf_paths and not excel_paths:
        print("No files found to process.")
        return

    # If you have a sample article for style reference, specify its path
    sample_article_pdf = './input/sample.pdf'  # Update this path if you have a sample article
    if os.path.exists(sample_article_pdf):
        sample_article_text = extract_text_from_pdf(sample_article_pdf)
        # You might analyze this text to generate style instructions
    else:
        sample_article_text = None

    # Define style instructions
    style_instructions = (
        "The article should have a formal and analytical tone.\n"
        "Use headings and subheadings to structure the content.\n"
        "Include an introduction, several data analysis sections, and a conclusion.\n"
        "Use bullet points and numbered lists where appropriate.\n"
        "Incorporate industry-specific terminology as seen in the sample article.\n"
        "The article should be in the style of a Chief Economist.\n"
        "The article should only be written based on the sections provided in the sample outline.\n"
        "Make sure the article isn't too long, recommended length is about 4-6 pages."
    )

    # Optionally, create a sample outline (shortened due to token limits)
    sample_outline = (
        # "1. Introduction\n"
        # "   - Brief overview of the main topics.\n"
        # "2. Data Analysis\n"
        # "   - Analysis of news articles.\n"
        # "   - Insights from PDFs.\n"
        # "   - Key findings from Excel data.\n"
        # "3. Conclusion\n"
        # "   - Summarize the key points.\n"
        # "   - Implications and recommendations."

        "SecondSTAX Africa View\n"
        "From the Chief Economist | Monday, October 21, 2024\n"
        "1. Weekly Recap\n"
        "- Paragraph summarizing the Ghanaian market performance\n"
        "- Paragraph summarizing the Kenyan market performance\n"
        "- Paragraph summarizing the Nigerian market performance\n"
        "2. Indices\n"
        "3. Treasuries\n"
        "4. Forex\n"
        "5. Commodities\n"
        "6. Economic Indicators\n"
        "7. Market Movers\n"
        "8. Recommendations Table\n"
        "9. News Moving the Markets\n"
        "10. SSX Markets Last Week\n"
        "11. Disclosure & Disclaimer Policy\n"

    )

    # Collect and compile data
    compiled_text = collect_data(news_json_files, pdf_paths, excel_paths)

    # Generate the analysis article
    article = generate_article(compiled_text, style_instructions, sample_outline)

    # Output the article
    print("\nGenerated Article:\n")
    print(article)

    # Optionally, save the article to a file
    with open('analysis_article.txt', 'w', encoding='utf-8') as file:
        file.write(article)

if __name__ == '__main__':
    main()