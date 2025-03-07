

import os
import json
import re
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openai import OpenAI
import io
import fitz  # For PDF handling
import docx2txt
import nltk
from nltk.tokenize import word_tokenize
from dotenv import load_dotenv

load_dotenv()


# Function to extract text from PDF using fitz
def extract_text_from_pdf(file_content):
    text = ''
    try:
        pdf_document = fitz.open(stream=file_content, filetype="pdf")
        for page_num in range(pdf_document.page_count):
            page = pdf_document.load_page(page_num)
            text += page.get_text()
        pdf_document.close()
    except Exception as e:
        print(f"Error extracting text from .pdf file: {e}")
    return text

# Function to extract text from .docx using docx2txt
def extract_text_from_docx(file_content):
    try:
        if file_content.startswith(b'\x50\x4b\x03\x04'):
            text = docx2txt.process(io.BytesIO(file_content))
            return text.strip()
        else:
            print("Unsupported file type (.docx)")
            return None
    except Exception as e:
        print(f"Error extracting text from .docx file: {e}")
        return None

# Function to analyze CV content
def analyze_cv_from_content(file_content):
    try:
        if file_content.startswith(b'%PDF'):
            extracted_text = extract_text_from_pdf(file_content)
        elif file_content.startswith(b'PK'):
            extracted_text = extract_text_from_docx(file_content)
        else:
            print("Unsupported file type")
            return None

        # Define JSON structure for extraction
        json_structure = {
            "basic_info": {
                "first_name": "",
                "last_name": "",
                "full_name": "",
                "gender": "",
                "location": "",
                "birth_date": "",
                "linkedin_url": "",
                "designation": "",
                "industry_function": "",
                "university": "",
                "education_level": "",
                "graduation_year": "",
                "graduation_month": "",
                "majors": "",
                "GPA": "",
            },
            "education": [{
                "degree_certificate": "",
                "institution": "",
                "passing_year": "",
                "completion_year": "",
            }],
            "work_experience": [{
                "job_title": "",
                "company": "",
                "location_job": "",
                "duration": "",
                "job_summary": "",
            }],
            "project_experience": [{
                "project_name": "",
                "project_description": "",
            }],
        }

        prompt = (
            "Your task is to read, analyze, and extract information from the CV. "
            "Important details are enclosed within triple backticks in a nested loop format - "
            f"```{extracted_text}```. "
            f"The required information is structured as follows: "
            f"```{json_structure}```. "
            "Fields to extract: "
            "Basic Info - first_name, last_name, full_name, gender(provide only if mentioned), location (pick only the city name from the permanent address of the candidate, if available, do not refer to preferred location or company's address, only take reference from residential or present address), birth_date, "
            "designation , industry_function(current industry functionality only), "
            "university (extract only the institute's name from the highest education degree achieved, excluding any certification), education_level (pick the highest education degree name mentioned under education or similar section in the CV), "
            "graduation_year, linkedin_url, graduation_month, majors, GPA(provide percentage ,CGPA,CQPI etc. for highest educational degree, if mentioned). "
            "Education - degree_certificate, institution, passing_year(only provide numerical year of completion of said degree,for example '2012',or so),"
            "completion_year(only provide numerical year). "
            "Work Experience - job_title, company, location_job (only pick the city name from the respective job address, if available), duration (extract only the time duration of the jobs, if available), job_summary (pls write only job description from the respective job only), "
            "Project Experience - Extract project details accurately from the 'Project Experience' section in CVs, "
            "For graduation_year : Retrieve the numerical completion year for the highest educational degree attained, focusing solely on year details,"
            "excluding any non-year information like percentages, divisions, or other text not representing a calendar year."
            "Differentiate 'project_name' to capture distinct project names and 'project_description' for respective detailed project descriptions, "
            "Consider variations in formatting where project details might start with keywords or phrases, ensuring 'project_name' specifically represents the project's name, "
            "Use your own intelligence to pick the name that may not be direct in some cases, "
            "Please consider variations in format and content across CVs while extracting information, "
            "For designation: extract only the current designation of the present job, look for keywords like 'till date' ,'since, or 'present' in the job duration to identify the current job , "
            "Indentify keywords like 'role','working as',job-role' etc in present job to get the designation, also there will be no end date mentioned for current job, "
            "For birth_date: Please extract and prioritize the date of birth (DOB) if available. Do not prioritize or consider age information; focus solely on retrieving the valid date of birth in the correct format if it is provided. "
            "For education: Retrieve all non-certification educational degrees incuding  'Masters', 'Bachelors', 'School', 'Executive', 'Professional degrees(CA, CS, ICWA)' - sorted in reverse chronology, excluding 'education_level' mentioned in basic_info, "
            " Look for fields like 'qualifications','education' or similar to fetch all degrees listed,if available,"
            "prioritizing the latest one first based on the completion year and passing year (if available). "
            "For work_experience: include all works or job experiences mentioned across the cv, prioritizing the latest job first based on duration or timeline, current job should come first. "
            "For project_experience: Look for sections that detail project experiences. "
            "Extract the project names and descriptions within these sections. "
            "Exclude achievements or other non-project related information. "
            "The projects can be labeled in various ways, such as 'Project Experience,' 'Worked on Projects,' 'Project Highlights,' or similar variations. "
            "Ensure the extracted information pertains solely to projects and not other types of experiences."
        )

        # Regular expressions for extracting basic info
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}(?:\.[A-Za-z]{2,})?\b'
        phone_pattern = r'(?:(?:\+\d{1,2}\s?)?\(?\d{3}\)?[/.\s-]?\d{3}[/.\s-]?\d{4}\b)|(?:\bphn\d+[/.\s-]?\d+\b)'

        basic_info = {"email": "", "phone_number": "", "portfolio_website_url": "", "github_main_page_url": ""}
        
        # Extract emails and phone numbers using regex
        email_matches = re.findall(email_pattern, extracted_text)
        phone_matches = re.findall(phone_pattern, extracted_text)

        basic_info["email"] = email_matches if email_matches else ""
        basic_info["phone_number"] = phone_matches if phone_matches else ""

        # Extract URLs
        url_pattern = r'\b(website|linkedin|github)\b[^A-Za-z\n]+(https?://\S+)\b'
        for match in re.finditer(url_pattern, extracted_text):
            key = match.group(1).lower() + "_url"
            basic_info[key] = match.group(2)

        # Extract GitHub URL
        github_pattern = r'https?://[\w./-]+(?:\n[\w./-]+)*github\.com/[\w./-]+'
        githubs = re.findall(github_pattern, extracted_text)
        if githubs:
            basic_info["github_main_page_url"] = githubs[0]

        # Prepare prompt for GPT
        prompt += f"```{extracted_text}```"
        
        # openAI to get structured data
        # Load environment variables from the .env file

        # Get the API key from the environment variable
        api_key = os.getenv("OPENAI_API_KEY")

        api_key = os.getenv("OPENAI_API_KEY")
        client = OpenAI(api_key=api_key)

        messages = [{"role": "system", "content": "You are an expert who can extract exact information from CVs"},
                    {"role": "user", "content": prompt}]
        response = client.chat.completions.create(model="gpt-3.5-turbo", messages=messages, temperature=0)
        # print(f"response************ : {response}")

        if response.choices:
            content = response.choices[0].message.content
            data_dict = json.loads(content)

            # Merge regex extracted info with GPT response
            merged_info = data_dict.copy() if data_dict else {}
            merged_info["basic_info"].update(basic_info)

            # Convert to JSON with ordered basic info
            json_output_merged = json.dumps(merged_info, indent=4)

            # Append to Excel file
            current_date_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            model_name = response.model
            prompt_tokens = response.usage.prompt_tokens
            completion_tokens = response.usage.completion_tokens
            total_tokens = response.usage.total_tokens

            data = {'Project Name': 'CV_Parsing', 'Date and Time': [current_date_time],
                    'Prompt Tokens': [prompt_tokens], 'Completion Tokens': [completion_tokens], 'Total Tokens': [total_tokens],
                    'Model Name': [model_name]}
            df = pd.DataFrame(data)
            append_to_excel(df)

            return json_output_merged

    except Exception as e:
        print(f"Error: {e}")
        return None

def append_to_excel(df):
    file_path = "..xlsx"
    existing_df = pd.read_excel(file_path)
    combined_df = pd.concat([existing_df, df], ignore_index=True)
    combined_df.to_excel(file_path, index=False)
    print(f"DataFrame appended to {file_path}")

if __name__ == "__main__":
    file_path = "Abhishek_kumar_resume.pdf"
    with open(file_path, 'rb') as file:
        file_content = file.read()
    analyze_cv_from_content(file_content)
